# -*- coding: utf-8 -*-
"""main.py —— 程序入口（命令行入口）

一条命令生成当天的《全国农产品批发价格日报》Excel：
    python main.py

常用参数：
    python main.py --categories 蔬菜 果品   只抓指定分类
    python main.py --list-categories        查看网站上有哪些分类
    python main.py --output 我的报表.xlsx   自定义输出文件名
    python main.py --no-detail              跳过逐品种报价（快速模式）

执行流程（编排层，具体干活在 fetcher / parser / report 里）：
    1. 拉品种目录树 —— 拿到分类和品种 id
    2. 拉今日涨跌幅排行（TOP20）
    3. 拉价格指数（日序列 + 分类指数）
    4. 拉今日报送市场名单
    5. 并发拉取所选分类下全部品种的今日报价（最耗时的一步，有进度显示）
    6. 全部整理后写入一个多工作表的 Excel 文件
"""
import argparse
import concurrent.futures
import sys
import threading
import time
from collections import OrderedDict
from datetime import datetime

import cache
import config
import fetcher
import parser as parse_mod
import report as report_mod


def build_cli():
    """定义命令行参数。

    用 argparse 标准库：用户在终端输入的参数会被自动解析，
    不认识的参数会给出提示（比如拼错的 --categries）。
    """
    cli = argparse.ArgumentParser(
        prog="python main.py",
        description="抓取农业农村部全国农产品批发市场价格数据，生成 Excel 日报",
    )
    cli.add_argument(
        "--categories",
        nargs="+",
        metavar="分类名",
        help="只抓取指定的大类（可多个，用空格隔开），默认：蔬菜 果品 畜禽产品 水产品",
    )
    cli.add_argument(
        "--list-categories",
        action="store_true",
        help="只列出网站上现有的全部大类名称，然后退出",
    )
    cli.add_argument(
        "--output",
        metavar="文件名",
        help="自定义输出 Excel 文件名（默认：农产品批发价格日报_日期.xlsx）",
    )
    cli.add_argument(
        "--no-detail",
        action="store_true",
        help="跳过逐品种报价抓取（只生成排行/指数/市场名单，速度更快）",
    )
    cli.add_argument(
        "--refresh",
        action="store_true",
        help="忽略今天的本地缓存，全部品种重新抓一遍",
    )
    return cli


def format_report_summary(sheets):
    """Format only workbooks sheets that actually exist, with compact numbering."""
    present = [(name, count, unit) for name, count, unit in sheets if count]
    return [
        f"{number}. {name} —— {count} {unit}"
        for number, (name, count, unit) in enumerate(present, start=1)
    ]


def configure_console_encoding():
    """Prevent unsupported progress symbols from crashing Windows GBK terminals."""
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if callable(reconfigure):
            try:
                reconfigure(errors="backslashreplace")
            except (OSError, ValueError):
                pass


def load_categories():
    """从网站拉取品种树，返回 {大类名: [品种...]} 和 {代码前缀: 大类名}。

    品种树是所有后续步骤的基础（品种 id 从这里拿），
    所以它如果失败，整个程序就没有继续的意义——直接退出并提示。
    """
    try:
        tree_payload = fetcher.get_json(config.API_VARIETIES_TREE)
        tree = parse_mod.parse_variety_tree(tree_payload)
        code_map = parse_mod.build_code_category_map(tree_payload)
    except (fetcher.FetchError, ValueError) as exc:
        # 品种树拿不到 = 后续全废，给出明确提示并退出（退出码 1 = 异常退出）
        print(f"[错误] 无法获取品种目录，程序终止。原因：{exc}")
        sys.exit(1)
    return tree, code_map


def pick_categories(tree, requested_names):
    """根据用户指定的分类名，从品种树里挑出要抓的分类。

    用户可能拼错分类名，这里做友好处理：
    如果指定了不存在的分类，列出可用的分类并退出，而不是抓一堆错的。
    """
    if not requested_names:
        requested_names = config.DEFAULT_CATEGORY_NAMES

    chosen = OrderedDict()
    for name in requested_names:
        if name in tree:
            chosen[name] = tree[name]
        else:
            print(f"[错误] 网站上没有叫「{name}」的分类。可用分类：{', '.join(tree.keys())}")
            sys.exit(1)
    return chosen


def fetch_growth_ranking(code_map):
    """拉今日涨跌幅排行。失败时给出警告并返回空列表（不致命）。"""
    try:
        payload = fetcher.post_json(config.API_GROWTH_RANKING, {"unitType": ""})
        rows, unit = parse_mod.parse_growth_ranking(payload, code_map)
        # 从第一条数据里顺便拿到"数据日期"（涨跌幅排行里的 date 数组）
        date_list = (payload.get("data") or {}).get("date") or []
        data_date = date_list[0] if date_list else None
        return rows, unit, data_date
    except (fetcher.FetchError, ValueError) as exc:
        print(f"[提示] 涨跌幅排行获取失败（已跳过）：{exc}")
        return [], "", None


def fetch_indexes():
    """拉价格指数（日序列 + 分类指数）。失败时警告并返回空列表。"""
    day_rows, level_rows = [], []
    try:
        day_rows = parse_mod.parse_price_index_day(
            fetcher.post_json(config.API_PRICE_INDEX_DAY)
        )
    except (fetcher.FetchError, ValueError) as exc:
        print(f"[提示] 价格指数（日序列）获取失败（已跳过）：{exc}")
    try:
        level_rows = parse_mod.parse_index_by_level(
            fetcher.post_json(config.API_INDEX_BY_LEVEL)
        )
    except (fetcher.FetchError, ValueError) as exc:
        print(f"[提示] 分类价格指数获取失败（已跳过）：{exc}")
    return day_rows, level_rows


def fetch_today_markets():
    """拉今日报送市场名单。失败时警告并返回空列表。"""
    try:
        return parse_mod.parse_today_markets(
            fetcher.post_json(config.API_TODAY_MARKETS, {"provinceCode": ""})
        )
    except (fetcher.FetchError, ValueError) as exc:
        print(f"[提示] 今日报送市场名单获取失败（已跳过）：{exc}")
        return []


# 抓取结果的三种"状态"标记（避免在函数间传一堆布尔值）
QUOTE_OK = "ok"            # 抓到数据（含"今天没报价"这种正常结果）
QUOTE_BANNED = "banned"    # 因为限流没抓到（等解封后可重试）
QUOTE_FAILED = "failed"    # 其他原因失败（网络抖动等）

# 中途存档间隔：每完成多少个抓取任务，就把已抓到的部分写进缓存一次
SAVE_EVERY_N = 30


def fetch_variety_quotes(chosen_categories, cached_quotes, data_date, refresh=False):
    """并发抓取所选分类下全部品种的今日报价（带缓存 + 限流自动续抓）。

    整体策略：
      1. 缓存里有结果（含"今天没报价"）的品种直接跳过，不重复请求网站；
      2. 抓取节奏放慢（见 config），尽量避免触发限流；
      3. 真被限流了：立刻停手（不再发新请求），倒计时等网站解封，然后自动续抓；
      4. 每抓完一批就把结果存进缓存，中途中断也不丢。

    返回：{品种id: 报价字典或 None}，以及统计信息（成功/失败/限流放弃的数量）。
    """
    # 把 {分类: [品种]} 摊平成任务列表，跳过缓存里已有的品种
    tasks = []
    cached_count = 0
    for category, varieties in chosen_categories.items():
        for variety in varieties:
            if not refresh and variety["id"] in cached_quotes:
                cached_count += 1
            else:
                tasks.append((category, variety))

    quotes = dict(cached_quotes)  # 最终结果从缓存内容起步
    failed_names = []             # 记录最终失败的品种名（结尾统一汇报）

    total = len(tasks)
    print(f"\n开始抓取品种今日报价（并发 {config.MAX_WORKERS} 线程）...")
    print(f"  共 {total + cached_count} 个品种：缓存命中 {cached_count} 个，需要抓取 {total} 个")

    pending = tasks          # 待抓清单（可能被限流截断后继续）
    ban_rounds = 0           # 已经历的"限流 -> 等待 -> 续抓"轮数
    start_time = time.time()
    done_count = 0
    batch_banned = []        # 先初始化，防止"全部命中缓存"时下面的收尾逻辑报错
    batch_failed = []

    while pending:
        # 跑一批：返回 (成功结果, 因限流失败的, 其他原因失败的, 限流等待秒数)
        # save_every 传入"只保存成功部分"的回调，实现批内中途存档
        batch_ok, batch_banned, batch_failed, ban_wait = _run_quote_batch(
            pending, save_every=lambda part: cache.save(data_date, {**quotes, **part})
        )
        quotes.update(batch_ok)
        done_count += len(batch_ok) + len(batch_banned) + len(batch_failed)

        # 每批结束立刻存档：即使程序被中断，已抓到的部分也不会丢
        cache.save(data_date, quotes)

        # 本轮没有失败：全部完成
        if not batch_banned and not batch_failed:
            break

        # 有因限流失败的：等待解封后只重试这些品种（前提是还没超过最大等待轮数）
        if batch_banned and ban_rounds < config.MAX_RATE_LIMIT_WAITS:
            ban_rounds += 1
            print()
            print(f"[提示] 网站限流，本轮有 {len(batch_banned)} 个品种被拦截。")
            print(f"   已自动停手，等待 {ban_wait} 秒后继续（第 {ban_rounds}/{config.MAX_RATE_LIMIT_WAITS} 次等待）...")
            if not _countdown_wait(ban_wait):
                # 用户在等待时按了 Ctrl+C：尊重用户，生成已有数据
                print("已中断等待，将用当前已抓到的数据生成报表。")
                failed_names.extend(v["name"] for _, v in batch_banned)
                break
            # 解封了：被拦截的 + 同一批里其他原因失败的，一起补抓
            pending = batch_banned + batch_failed
            continue

        # 走到这里：要么限流等待次数用完，要么只是普通失败。不再纠缠，收尾。
        failed_names.extend(v["name"] for _, v in batch_banned + batch_failed)
        if batch_banned:
            print(f"\n[提示] 限流等待次数已用完，剩余 {len(batch_banned)} 个品种本次放弃（可用 --refresh 稍后补抓）。")
        elif batch_failed:
            print(f"\n[提示] 有 {len(batch_failed)} 个品种抓取失败（网络原因），本次跳过。")
        break

    # 收尾：把最终结果再完整存一次缓存
    cache.save(data_date, quotes)

    print()
    elapsed = time.time() - start_time
    print(f"品种报价抓取完成：已获取 {len(quotes)} 个品种（含缓存），本次处理 {done_count} 个抓取任务，用时 {elapsed:.0f} 秒")
    if failed_names:
        print(f"本次未抓到的品种（{len(failed_names)} 个）：{'、'.join(failed_names[:8])}"
              + ("..." if len(failed_names) > 8 else ""))

    # 返回 (报价字典, 失败品种id集合)，方便 build_quote_rows 判断
    failed_ids = {v["id"] for _, v in batch_banned + batch_failed} if (batch_banned or batch_failed) else set()
    return quotes, failed_ids


def _run_quote_batch(tasks, save_every=None):
    """跑一批品种抓取任务（线程池）。

    限流保护：一旦某个线程发现被限流，立刻竖起"停止牌"（ban_flag），
    其他还在排队的任务看到停止牌就直接返回"被拦截"，不再发请求——
    避免在被封的 3 分钟里还傻乎乎地发几百个注定失败的请求。

    参数 save_every：可选，每成功处理 N 个任务就调用一次（用来中途存档，
    防止长任务被 Ctrl+C 打断后前功尽弃）。

    返回：(成功结果dict, 被拦截列表, 失败列表, 限流等待秒数或None)
    """
    ban_flag = threading.Event()  # 停止牌：set() 之后所有排队任务立即放弃
    ban_wait = [None]             # 用列表包装，方便在闭包里修改
    ok = {}
    banned = []
    failed = []

    def work(item):
        """线程池里跑的单个任务。"""
        category, variety = item
        # 开工前先看停止牌：已被限流就直接放弃，别浪费请求
        if ban_flag.is_set():
            return QUOTE_BANNED
        time.sleep(config.POLITE_DELAY)  # 礼貌间隔：慢一点，稳一点
        if ban_flag.is_set():
            return QUOTE_BANNED
        try:
            quote = fetcher.fetch_variety_price(variety["id"])
            return QUOTE_OK, quote  # quote 可能为 None（该品种今天没人报价，属正常）
        except fetcher.RateLimitedError as exc:
            # 被限流：竖起停止牌，让其他任务全部停止
            ban_flag.set()
            if ban_wait[0] is None or exc.wait_seconds > ban_wait[0]:
                ban_wait[0] = exc.wait_seconds
            return QUOTE_BANNED
        except fetcher.FetchError:
            return QUOTE_FAILED

    with concurrent.futures.ThreadPoolExecutor(max_workers=config.MAX_WORKERS) as pool:
        future_map = {pool.submit(work, t): t for t in tasks}
        processed = 0
        for future in concurrent.futures.as_completed(future_map):
            category, variety = future_map[future]
            processed += 1
            try:
                result = future.result()
            except Exception:
                # 理论上到不了这里（work 内部已兜底），保留一层保护
                failed.append((category, variety))
                continue
            if isinstance(result, tuple) and result[0] == QUOTE_OK:
                ok[variety["id"]] = result[1]
            elif result == QUOTE_BANNED:
                banned.append((category, variety))
            else:
                failed.append((category, variety))

            # 中途存档：每完成 SAVE_EVERY_N（30）个任务，就把已抓到的部分写进缓存。
            # 注意：save_every 是"保存动作"（一个函数），取余要用常量 SAVE_EVERY_N。
            if save_every and processed % SAVE_EVERY_N == 0:
                save_every(ok)

            # 进度显示：每完成 25 个（或最后 1 个）打印一整行。
            # 不用 \r 覆盖刷新——它在部分终端渲染不出来，会显得像卡死。
            if processed % 25 == 0 or processed == len(tasks):
                print(
                    f"  进度：{processed}/{len(tasks)}"
                    f"（成功 {len(ok)}，无报价/被拦截 {len(banned) + len(failed)}）"
                )

    return ok, banned, failed, ban_wait[0]


def _countdown_wait(seconds):
    """倒计时等待限流解除；用户在等待时按 Ctrl+C 可以中断。

    为什么每 10 秒打印一整行而不是刷新同一行？
    之前用 \r（回车符）覆盖同一行刷新，在部分终端（比如某些 PowerShell）
    里渲染不出来，看起来像程序卡死了。改成"每 10 秒打印一行 + 最后 5 秒逐秒
    打印"，任何终端都能看到进度，总共也就 20 来行，不算刷屏。

    返回：True 表示完整等完了；False 表示被用户中断。
    """
    try:
        while seconds > 0:
            if seconds % 10 == 0 or seconds <= 5:
                print(f"   剩余 {seconds:>3} 秒（按 Ctrl+C 可中断，直接生成已有数据）")
            time.sleep(1)
            seconds -= 1
        return True
    except KeyboardInterrupt:
        print()
        return False


def build_quote_rows(chosen_categories, quotes, failed_ids):
    """把每个品种的报价数据汇总成报表行。

    每个品种一行：分类 / 品种名 / 报价市场数 / 最低价 / 最高价 / 平均价 / 数据日期。
    品种的三种状态：
      - 有报价：正常统计一行；
      - 今天没人报价（quote 为 None）：跳过（没数据可统计）；
      - 抓取失败（在 failed_ids 里）：跳过，失败名单由调用方另行提示。
    """
    rows = []
    for category, varieties in chosen_categories.items():
        for variety in varieties:
            quote = quotes.get(variety["id"])
            if quote is None or variety["id"] in failed_ids:
                continue
            summary = parse_mod.summarize_quotes(quote)
            if summary is None:
                continue
            row = {"分类": category, "品种名称": variety["name"]}
            row.update(summary)
            row["数据日期"] = quote.get("date", "")
            rows.append(row)
    # 按 分类 -> 品种名 排序，同一类的品种排在一起，报表更好看
    rows.sort(key=lambda r: (r["分类"], r["品种名称"]))
    return rows


def generate_report(output_path, data_date, ranking_rows, quote_rows, day_rows, level_rows, market_rows, quote_note):
    """把各种数据组织成多工作表结构，交给 report.py 写文件。"""
    sheets = OrderedDict()

    # Sheet1 涨跌幅排行：品种 + 价格 + 涨跌幅，附上数据来源说明
    if ranking_rows:
        sheets["涨跌幅排行TOP20"] = (
            ranking_rows,
            "来源：全国农产品批发市场价格信息系统「涨跌排行」（涨幅前10 + 跌幅前10）",
        )

    # Sheet2 分类品种批发价：本工具的核心数据（各分类全部品种的今日报价统计）
    if quote_rows:
        sheets["分类品种批发价"] = (quote_rows, quote_note)

    # Sheet3 价格指数日序列
    if day_rows:
        sheets["批发价格指数(日)"] = (
            day_rows,
            "全国农产品批发价格 200 指数（以 2015 年为基期 100）",
        )

    # Sheet4 分类价格指数：今日 vs 昨日
    if level_rows:
        sheets["分类价格指数"] = (level_rows, "各分类今日指数值及昨日指数值")

    # Sheet5 今日报送市场
    if market_rows:
        sheets["今日报送市场"] = (market_rows, f"今日共 {len(market_rows)} 家市场报送了价格")

    # 核心数据一个都没有 = 本次运行没有产出，明确提示而不是悄悄生成空文件
    if not sheets:
        print("[错误] 所有数据都没有抓到，未生成报表。请检查网络后重试。")
        sys.exit(1)

    try:
        report_mod.build_excel(output_path, data_date or _today(), sheets)
    except PermissionError:
        # Windows 上最常见的场景：目标文件正在 Excel 里打开着，被系统锁住了
        print(f"[错误] 无法写入报表文件「{output_path}」。")
        print("   原因：该文件被其他程序占用（很可能正开着 Excel 看它）。")
        print("   处理：关闭 Excel 里的这个文件后重新运行即可；")
        print("         数据已在本地缓存，重新运行只需要几秒钟。")
        sys.exit(1)


def _today():
    """今天的日期字符串（备用：接口拿不到数据日期时用本地日期）。"""
    return datetime.now().strftime("%Y-%m-%d")


# ==================== 运行锁 ====================
# 为什么需要锁？这个网站有严格的限流配额，如果同时跑两个抓取任务，
# 两个任务会互相抢配额、互相把对方"限流"，最后双双卡在"等解封->重试->再被封"
# 的循环里。所以用锁文件保证同一时间只有一个任务在跑。

def _pid_alive(pid):
    """判断某个进程号是否还活着（Windows 用系统 API 查询）。"""
    try:
        import ctypes

        # 0x1000 = PROCESS_QUERY_LIMITED_INFORMATION：只查询、不动进程
        handle = ctypes.windll.kernel32.OpenProcess(0x1000, False, pid)
        if not handle:
            return False
        ctypes.windll.kernel32.CloseHandle(handle)
        return True
    except (OSError, AttributeError):
        # 非 Windows 系统或查询失败：保守起见当作"还活着"
        return True


def acquire_run_lock():
    """运行前加锁：如果已有任务在跑，给出明确提示并退出。"""
    import os

    lock_path = os.path.join(config.CACHE_DIR, ".run.lock")
    os.makedirs(config.CACHE_DIR, exist_ok=True)

    if os.path.exists(lock_path):
        try:
            with open(lock_path, "r", encoding="utf-8") as file:
                old_pid = int(file.read().strip())
            # 旧锁对应的进程还活着 = 真有任务在跑，不能抢
            if _pid_alive(old_pid):
                print(f"[错误] 检测到另一个抓取任务正在运行（进程号 {old_pid}）。")
                print("   两个任务同时爬会互相触发网站限流，所以这里拦住了。")
                print("   处理：等它跑完，或先在任务管理器里结束它，再重新运行。")
                sys.exit(1)
            # 进程已死 = 上次异常退出留下的旧锁，可以安全覆盖
        except (ValueError, OSError):
            pass  # 锁文件损坏：当作没有锁

    # 写入自己的进程号，表示"我占着坑了"
    with open(lock_path, "w", encoding="utf-8") as file:
        file.write(str(os.getpid()))


def release_run_lock():
    """运行结束（无论成功失败）释放锁。"""
    import os

    lock_path = os.path.join(config.CACHE_DIR, ".run.lock")
    try:
        os.remove(lock_path)
    except OSError:
        pass  # 锁文件本来就不存在，忽略


def main():
    """程序入口：按顺序编排上面所有步骤。"""
    configure_console_encoding()
    cli = build_cli()
    args = cli.parse_args()

    # 加运行锁：同一时间只允许一个抓取任务（防止互抢配额互相限流）
    acquire_run_lock()

    print("=" * 60)
    print("全国农产品批发市场价格日报生成器")
    print("数据来源：农业农村部 · 全国农产品批发市场价格信息系统")
    print("=" * 60)

    # 第 0 步：拿到品种树（基础数据）
    tree, code_map = load_categories()

    # --list-categories：只想看有哪些分类，列完直接退出
    if args.list_categories:
        print("\n网站上现有的大类：")
        for name, varieties in tree.items():
            print(f"  {name}（{len(varieties)} 个品种）")
        return

    # 第 1 步：确定要抓的分类
    chosen = pick_categories(tree, args.categories)
    print(f"\n本次抓取分类：{'、'.join(chosen.keys())}")

    # 第 2 步：拉排行榜 + 指数 + 市场名单（快，一起做）
    print("\n正在获取涨跌幅排行、价格指数、报送市场名单...")
    ranking_rows, unit, data_date = fetch_growth_ranking(code_map)
    day_rows, level_rows = fetch_indexes()
    market_rows = fetch_today_markets()
    if ranking_rows:
        print(f"  [完成] 涨跌幅排行 {len(ranking_rows)} 条（数据日期：{data_date}）")
    if day_rows:
        print(f"  [完成] 价格指数日序列 {len(day_rows)} 天")
    if market_rows:
        print(f"  [完成] 今日报送市场 {len(market_rows)} 家")

    # 缓存按"数据日期"分天存放；拿不到数据日期就用本地日期兜底
    run_date = data_date or _today()

    # 第 3 步：并发抓取各品种今日报价（最耗时的一步）
    quotes, failed_ids = {}, set()
    if not args.no_detail:
        cached_quotes = cache.load(run_date)
        quotes, failed_ids = fetch_variety_quotes(chosen, cached_quotes, run_date, args.refresh)

    # 第 4 步：汇总品种报价
    quote_rows = build_quote_rows(chosen, quotes, failed_ids)
    quoted = len(quote_rows)
    total_varieties = sum(len(v) for v in chosen.values())
    quote_note = (
        f"来源：各批发市场今日报送的品种报价汇总。"
        f"本次抓取 {len(chosen)} 个分类共 {total_varieties} 个品种，"
        f"其中 {quoted} 个品种今日有市场报价。"
    )
    if quote_rows:
        print(f"  [完成] 品种报价汇总 {quoted} 行")

    # 第 5 步：写 Excel 文件
    output_path = args.output or config.OUTPUT_FILE_TEMPLATE.format(date=data_date or _today())
    generate_report(
        output_path, data_date, ranking_rows, quote_rows,
        day_rows, level_rows, market_rows, quote_note,
    )

    # 收尾：打印本次运行小结
    print("\n" + "=" * 60)
    summary_lines = format_report_summary([
        ("涨跌幅排行TOP20", len(ranking_rows), "条"),
        ("分类品种批发价", len(quote_rows), "条（本次核心数据）"),
        ("批发价格指数(日)", len(day_rows), "天"),
        ("分类价格指数", len(level_rows), "条"),
        ("今日报送市场", len(market_rows), "家"),
    ])
    print(f"报表包含 {len(summary_lines)} 个工作表：")
    for line in summary_lines:
        print(f"  {line}")
    print("=" * 60)


if __name__ == "__main__":
    try:
        main()
    finally:
        # 无论正常结束、报错退出还是 Ctrl+C，都释放运行锁
        release_run_lock()

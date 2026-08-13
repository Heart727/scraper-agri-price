# -*- coding: utf-8 -*-
"""config.py —— 全局配置

为什么单独一个配置文件？
把所有"以后可能会变"的东西集中放在这里：
网站接口地址、解密密钥、抓取范围、请求频率等。
将来接口地址变了，或者想换一批分类抓，只需要改这个文件，
不用去翻核心逻辑代码。这就是"配置和代码分离"的好处。

（这套接口是 2026-08 从 pfsc.agri.cn 网站前端 JS 源码里
找到并逐一实测通过的，属于网站自己公开使用的接口。）
"""

# ==================== 网站接口地址 ====================
# 数据来源：农业农村部"全国农产品批发市场价格信息系统"（pfsc.agri.cn）
BASE_URL = "https://pfsc.agri.cn"

# 1) 今日全国涨跌幅排行 TOP20
#    返回：品种名、今日均价、昨日均价、涨跌额、涨跌幅(%) 等
API_GROWTH_RANKING = BASE_URL + "/price_portal/index/growthRanking"

# 2) 品种目录树
#    返回：大类 -> 子类 -> 品种 的完整目录（大类含：蔬菜、果品、畜禽产品、水产品等）
#    用途：拿到每个品种的 id（调报价接口要用），以及"品种 -> 大类"的对应关系
API_VARIETIES_TREE = BASE_URL + "/price_portal/sys-user-relation/getVarietiesTree"

# 3) 指定品种的今日各市场报价（AES 加密返回）
#    注意：品种 id 必须放在"查询字符串"里传（网站前端就是这么调的），
#    放在请求体里会无效。返回数据整体是加密的，需要用下面的密钥解密。
API_VARIETY_PRICE = BASE_URL + "/price_portal/index/getMarketReportPriceChart"

# 4) 全国农产品批发价格指数（日序列，约一个月）
#    返回：每日的农产品指数、粮油指数、菜篮子指数
API_PRICE_INDEX_DAY = BASE_URL + "/price_portal/pi-info-day/getPortalPiInfoDay"

# 5) 分类价格指数（各分类今日指数值 + 昨日值，可算涨跌）
API_INDEX_BY_LEVEL = BASE_URL + "/price_portal/pi-info-day/getIndexByLevel"

# 6) 今日已报送报价的市场名单（约 190 家）
API_TODAY_MARKETS = BASE_URL + "/api/priceQuotationController/getTodayMarketByProvinceCode"

# ==================== AES 解密配置 ====================
# 报价接口返回的数据是 AES-CBC 加密的密文。
# 解密规则（来自网站前端 JS，属于公开信息）：
#   - 密文的前 16 个字符是 IV（初始向量）
#   - 剩余部分是 Base64 编码的密文
#   - 用下面这个 32 字节的固定密钥解密
AES_KEY = b"7s9K$pG2xQ8zR5mB7vA3sD9fH2jW40cV"

# ==================== 请求参数 ====================
# 浏览器标识：很多网站会拒绝没有 User-Agent 的请求（以为是机器人）
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"
)

TIMEOUT = 30        # 单个请求的超时时间（秒），超过就放弃并提示
RETRIES = 2         # 失败后自动重试的次数（总共最多试 1 + 2 = 3 次）

# ---- 抓取节奏（重要！）----
# 实测发现：该网站约 3 分钟内只容忍几百次请求，超了会封禁 3 分钟
# （返回 code=403「访问频率过高」）。所以这里刻意放慢：
#   2 个线程并发 + 每线程每次请求前等 0.9 秒 ≈ 每秒 1.7 个请求左右，
# 贴着网站的容忍度走。宁可慢一点，也不要被封。
# 想调快/调慢改这两个数字即可（越快越容易被封，越慢越安全）。
POLITE_DELAY = 0.9  # 每个线程两次请求之间的间隔（秒）
MAX_WORKERS = 2     # 抓取品种报价时同时工作的线程数

# 一次运行中，被限流后最多自动"等待解封 + 续抓"几轮（每轮等约 3 分钟）。
# 超过这个次数就放弃剩余品种，用已经抓到的数据生成报表，避免无限等待。
MAX_RATE_LIMIT_WAITS = 3

# ---- 本地缓存 ----
# 同一天反复运行时，已抓过的品种直接读缓存，不再请求网站。
# 缓存按天存放：cache/2026-08-13.json
CACHE_DIR = "cache"

# ==================== 抓取范围 ====================
# 默认抓取这 4 个大类的全部品种报价（大类名必须和网站上显示的一致）。
# 用 --categories 参数可以换成其他分类，用 --list-categories 可以查看全部分类。
DEFAULT_CATEGORY_NAMES = ["蔬菜", "果品", "畜禽产品", "水产品"]

# ==================== 输出文件 ====================
# 输出文件名模板：{date} 会被替换成数据日期（比如 2026-08-13）
OUTPUT_FILE_TEMPLATE = "农产品批发价格日报_{date}.xlsx"

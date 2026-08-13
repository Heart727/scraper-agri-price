# -*- coding: utf-8 -*-
"""parser.py —— 数据解析层：把接口返回的原始数据整理成"能直接用"的结构

网络层（fetcher.py）只负责把原始 JSON 拿回来，
这个文件负责把 JSON 里嵌套很深的字段拆出来、整理干净。
比如：从品种树里抽出"分类 -> 品种列表"、给报价算最高/最低/平均价。
"""
from collections import OrderedDict

import config


# ==================== 品种树解析 ====================

def parse_variety_tree(payload):
    """把品种树的原始 JSON 解析成容易使用的结构。

    原始结构（嵌套三层）：
      大类（蔬菜 AE） -> 子类（叶菜类等） -> attributelist（品种列表）
    品种列表里每个品种有：id（数字字符串）、varietyName（品种名）。

    返回（OrderedDict，保持网站原始顺序）：
      { "蔬菜": [{"id": "77", "name": "大白菜"}, ...],
        "果品": [{"id": ..., "name": ...}, ...], ... }
    """
    # 一路往下钻，拿到 10 个大类节点；结构固定为 content[0].content.children
    try:
        categories = payload["content"][0]["content"]["children"]
    except (KeyError, IndexError, TypeError) as exc:
        # 网站改版导致结构对不上时，抛出带说明的错误而不是崩溃
        raise ValueError(f"品种树的数据结构和预期不一致，网站可能改版了（{exc}）")

    result = OrderedDict()
    for category in categories:
        name = category.get("label")  # 大类名，如"蔬菜"
        varieties = []
        # 大类下的每个子类（如"叶菜类"）都挂着一个品种列表 attributelist
        for sub in category.get("children") or []:
            for item in sub.get("attributelist") or []:
                varieties.append(
                    {"id": str(item.get("id")), "name": item.get("varietyName")}
                )
        # 只保留有品种的大类，空分类没有抓取价值
        if name and varieties:
            result[name] = varieties

    # 最后做个安全检查：目录里必须有内容，否则后面全是无用功
    if not result:
        raise ValueError("品种树解析结果为空，请确认接口返回正常")
    return result


def build_code_category_map(payload):
    """建立"品种代码前缀 -> 大类名"的映射表。

    用途：涨跌幅排行接口返回的品种代码形如 "AE02006"，
    前缀 AE 代表蔬菜、AF 代表果品、AL 代表畜禽产品。
    有了这张表，就能给排行数据自动加上"所属大类"这一列。

    返回：{"AE": "蔬菜", "AF": "果品", ...}
    """
    try:
        categories = payload["content"][0]["content"]["children"]
    except (KeyError, IndexError, TypeError):
        return {}
    return {str(c.get("code")): c.get("label") for c in categories if c.get("code")}


# ==================== 涨跌幅排行解析 ====================

def parse_growth_ranking(payload, code_map):
    """把涨跌幅排行的原始 JSON 解析成一行行的表格数据。

    原始结构：data 里有几个平行的数组（names 品种名、avgPrice 今日均价、
    lastAvgPrice 昨日均价、priceHbs 涨跌幅……），
    需要按"同一位置属于同一个品种"的规则把它们并成一行。

    返回：list[dict]，每行包含 品种名/所属大类/今日均价/昨日均价/涨跌额/涨跌幅
    """
    data = payload.get("data") or {}
    names = data.get("names") or []
    codes = data.get("codes") or []
    avg_price = data.get("avgPrice") or []
    last_price = data.get("lastAvgPrice") or []
    diff_price = data.get("difAvgPrice") or []
    change_pct = data.get("priceHbs") or []
    unit = data.get("meteringUnit", "元/公斤")  # 计量单位，默认元/公斤

    rows = []
    for i, name in enumerate(names):
        code = codes[i] if i < len(codes) else ""
        # 代码前两位是大类代码（如 AE），查不到就标"其他"
        category = code_map.get(code[:2], "其他")
        rows.append(
            {
                "品种名称": name,
                "所属大类": category,
                "今日均价(元/公斤)": _to_number(avg_price, i),
                "昨日均价(元/公斤)": _to_number(last_price, i),
                "涨跌额(元/公斤)": _to_number(diff_price, i),
                "涨跌幅(%)": _to_number(change_pct, i),
            }
        )
    return rows, unit


def _to_number(array, index):
    """安全地从数组里取第 index 个值并转成数字。

    接口偶尔会缺字段或给空字符串，这里统一兜底：
    取不到或转不成数字就返回 None（Excel 里显示为空，不报错）。
    """
    try:
        value = array[index]
        return float(value)
    except (IndexError, TypeError, ValueError):
        return None


# ==================== 品种报价解析 ====================

def summarize_quotes(quote_data):
    """把某品种的"各市场报价列表"汇总成一行统计。

    原始数据：190 家市场里，报了该品种价格的有 N 家，
    每家一个价格。这里汇总成：报价市场数 / 最低价 / 最高价 / 平均价。
    平均价 = 全部市场价格的简单算术平均（与网站口径一致）。
    """
    prices = [float(p) for p in (quote_data.get("prices") or []) if p is not None]
    if not prices:
        # 理论上不会走到这里（上游已过滤空数据），但保留保护
        return None
    return {
        "报价市场数": len(prices),
        "最低价(元/公斤)": round(min(prices), 2),
        "最高价(元/公斤)": round(max(prices), 2),
        "平均价(元/公斤)": round(sum(prices) / len(prices), 2),
    }


# ==================== 指数解析 ====================

def parse_price_index_day(payload):
    """解析"农产品批发价格指数"日序列。

    返回：list[dict]，每行 = 一天，含 日期/农产品指数/粮油指数/菜篮子指数。
    网站按日期倒序返回（最新在前），这里按日期正序排列方便看趋势。
    """
    rows = []
    for item in payload.get("content") or []:
        rows.append(
            {
                "日期": item.get("publishDate", "")[:10],  # 去掉时间部分只留日期
                "农产品批发价格200指数": _to_number([item.get("agriculture")], 0),
                "粮油产品批发价格指数": _to_number([item.get("grainAndOil")], 0),
                "菜篮子产品批发价格指数": _to_number([item.get("vegetableBasket")], 0),
            }
        )
    # 按日期从小到大排（"2026-08-01" 这种格式字符串排序即时间排序）
    rows.sort(key=lambda r: r["日期"])
    return rows


def parse_index_by_level(payload):
    """解析"分类价格指数"（今日指数值 vs 昨日指数值）。

    返回：list[dict]，每行 = 一个分类，含 分类名/今日指数/昨日指数。
    涨跌幅可以由报告层计算：(今日-昨日)/昨日*100。
    """
    rows = []
    # 原始结构是嵌套数组：content -> [大类数组, 子类数组, ...]，展平后逐条处理
    for group in payload.get("content") or []:
        for item in group or []:
            name = item.get("indexName")
            if name:
                rows.append(
                    {
                        "分类": name,
                        "今日指数": _to_number([item.get("indexValue")], 0),
                        "昨日指数": _to_number([item.get("nextRecordValue")], 0),
                    }
                )
    return rows


def parse_today_markets(payload):
    """解析"今日报送市场名单"。

    返回：list[dict]，每行 = 一个市场（市场名 + 代码）。
    注：该接口只返回市场名单，不返回具体价格。
    """
    rows = []
    for item in payload.get("content") or []:
        name = item.get("marketName") or item.get("szsmMarketName")
        if name:
            rows.append({"市场名称": name, "市场代码": item.get("marketCode", "")})
    return rows

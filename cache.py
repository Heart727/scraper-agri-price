# -*- coding: utf-8 -*-
"""cache.py —— 本地缓存：同一天重复运行时，已抓过的品种不再请求网站

为什么需要缓存？两个原因：
1. 网站有严格的限流（约 3 分钟内只能访问几百次），重复请求既慢又容易被封。
2. 每天的数据是当天有效，所以缓存"按天"存一份，日期变了自动失效。

缓存文件格式（JSON）：
    cache/2026-08-13.json
    {
      "saved_at": "2026-08-13 10:30:00",
      "quotes": {
        "77":  {"date": "2026-08-13", "markets": [...], "prices": [...]},
        "80":  null,          // null 表示"该品种今天没有市场报价"（也是结果，缓存下来）
        ...
      }
    }
"""
import json
import os
from datetime import datetime

import config


def _cache_path(date_str):
    """算出某天的缓存文件路径（cache/2026-08-13.json）。"""
    return os.path.join(config.CACHE_DIR, f"{date_str}.json")


def load(date_str):
    """读取某天的缓存。

    参数 date_str：数据日期（如 "2026-08-13"）
    返回：{品种id: 报价字典或 None}；没有缓存文件时返回空字典 {}。
    注：报价为 None 表示"该品种今天确实没有报价"，和"没抓过"要区分开——
    区分方式是用 `品种id in 缓存` 判断（见 main.py 的用法）。
    """
    path = _cache_path(date_str)
    if not os.path.exists(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as file:
            data = json.load(file)
        return data.get("quotes") or {}
    except (json.JSONDecodeError, OSError):
        # 缓存文件损坏：当它不存在（重新抓一遍，慢一点但结果正确）
        return {}


def save(date_str, quotes):
    """把抓到的品种报价写入当天的缓存文件。

    每抓完一批就存一次，这样即使中途被 Ctrl+C 打断，
    已经抓到的部分也留在磁盘上，下次运行直接续用。
    """
    path = _cache_path(date_str)
    # 确保 cache 目录存在（不存在就创建）
    os.makedirs(config.CACHE_DIR, exist_ok=True)
    data = {
        "saved_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "quotes": quotes,
    }
    # 先写临时文件再改名，避免写到一半断电导致缓存文件损坏
    temp_path = path + ".tmp"
    with open(temp_path, "w", encoding="utf-8") as file:
        json.dump(data, file, ensure_ascii=False)
    os.replace(temp_path, path)

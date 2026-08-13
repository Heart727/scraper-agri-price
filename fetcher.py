# -*- coding: utf-8 -*-
"""fetcher.py —— 网络层：发请求、解密、重试、错误处理

职责单一：这个文件只负责"跟网站打交道"。
其他文件（main.py / parser.py / report.py）想拿数据，
只需要调用这里封装好的几个函数，不用关心 HTTP 细节。

错误处理原则（也是本项目的要求）：
出错时给出明确的中文提示，绝不"静默失败"——
也就是说，不会出现"程序没报错，但数据悄悄没了"的情况。
"""
import base64
import json
import time

import requests
from Crypto.Cipher import AES
from Crypto.Util.Padding import unpad

import config


class FetchError(Exception):
    """自定义异常：所有"拿数据失败"的情况都抛这个。

    好处：main.py 里只需要捕获这一种异常，就能兜住
    断网、超时、接口报错、数据为空等所有网络层问题。
    """

    def __init__(self, message):
        # 中文提示统一加上"抓取失败"前缀，用户一眼就能看懂
        super().__init__(f"抓取失败：{message}")


class RateLimitedError(FetchError):
    """限流异常：请求太频繁，被网站临时封禁（code=403）。

    比普通 FetchError 多带一个信息：wait_seconds（还要等多少秒才能再访问）。
    调用方拿到这个信息后可以"等待一段时间再继续"，而不是傻乎乎地硬闯。
    """

    def __init__(self, message, wait_seconds):
        super().__init__(message)
        self.wait_seconds = wait_seconds


def _make_session():
    """创建一个配置好的 HTTP 会话。

    为什么用"会话"（Session）而不是每次新开连接？
    因为 Session 会复用 TCP 连接，连续请求几百次时能明显提速。
    统一挂上 User-Agent，避免被网站当成机器人拒绝。
    """
    session = requests.Session()
    session.headers.update(
        {
            "User-Agent": config.USER_AGENT,
            "Referer": config.BASE_URL + "/",  # 网站有防盗链，不带 Referer 会被拒绝
        }
    )
    return session


def _request_with_retry(method, url, **kwargs):
    """带重试的请求核心函数。

    流程：尝试发送请求 -> 失败就稍等片刻重试 -> 超过重试次数就抛异常。
    任何一次成功都会直接返回；只有全部失败才报错。
    每次重试之间留一点时间，既礼貌又给网络抖动留缓冲。
    """
    session = _make_session()
    last_error = None

    # 总共尝试 1 + RETRIES 次（比如 RETRIES=2 就是最多试 3 次）
    for attempt in range(config.RETRIES + 1):
        try:
            response = session.request(method, url, timeout=config.TIMEOUT, **kwargs)
            # HTTP 状态码 4xx/5xx 也算失败，主动抛出来走重试
            response.raise_for_status()
            return response
        except requests.RequestException as exc:
            # 记录最后一次错误信息，重试结束后用中文包装再抛出
            last_error = exc
            if attempt < config.RETRIES:
                # 等待时间随重试次数翻倍（第 1 次等 1 秒，第 2 次等 2 秒）
                time.sleep(2 ** attempt)

    # 走到这里说明全部重试都失败了，给出明确的中文提示
    hint = "请检查：1) 网络是否正常；2) 网站 https://pfsc.agri.cn 是否能打开"
    raise FetchError(f"请求 {url} 连续失败 {config.RETRIES + 1} 次（{last_error}）。{hint}")


def _parse_json_response(response, url):
    """把 HTTP 响应解析成 JSON 字典，并检查业务状态码。

    网站的接口有两种返回格式：
      - 新接口：{"code": 0, "msg": "success", "data": ...}，code=0 表示成功
      - 老接口：{"code": 200, "message": "请求处理成功", "content": ...}，code=200 表示成功
    这里统一处理，业务失败也抛 FetchError，绝不把错误数据当正常数据用。
    """
    try:
        payload = response.json()
    except ValueError as exc:
        # 返回的不是 JSON（比如被防火墙拦了，返回了 HTML）
        raise FetchError(f"接口 {url} 返回的不是 JSON 数据，可能被网站拦截了（{exc}）")

    code = payload.get("code")
    if code == 403:
        # 403 有两种含义：访问太频繁（限流）或其他原因。
        # 限流时网站的提示里会写"频率"，并且附带还要等多久，
        # 把等待秒数解析出来交给调用方，调用方就能"等一等再继续"。
        message = payload.get("msg") or payload.get("message") or "未知原因"
        if "频率" in message:
            raise RateLimitedError(
                f"访问频率过高，被网站临时限制（{message}）",
                _parse_wait_seconds(message),
            )
        raise FetchError(f"接口 {url} 返回业务错误：code=403，{message}")
    if code not in (0, 200):
        message = payload.get("msg") or payload.get("message") or "未知原因"
        raise FetchError(f"接口 {url} 返回业务错误：code={code}，{message}")
    return payload


def _parse_wait_seconds(message):
    """从限流提示里解析"还要等多少秒"。

    网站的提示形如："系统限制您0小时2分钟59秒后才可再次访问"。
    用正则把 小时/分钟/秒 三个数字抠出来换算成总秒数。
    解析不出来就默认 180 秒（3 分钟），宁多等不硬闯。
    """
    import re

    match = re.search(r"(\d+)\s*小时\s*(\d+)\s*分钟\s*(\d+)\s*秒", message)
    if match:
        hours, minutes, seconds = (int(g) for g in match.groups())
        return hours * 3600 + minutes * 60 + seconds
    return 180


def get_json(url):
    """发 GET 请求并返回 JSON（品种树等接口用）。"""
    response = _request_with_retry("GET", url)
    return _parse_json_response(response, url)


def post_json(url, data=None):
    """发 POST 请求（参数放请求体），返回 JSON（涨跌幅排行、指数等接口用）。"""
    response = _request_with_retry("POST", url, json=data or {})
    return _parse_json_response(response, url)


def post_query(url, params):
    """发 POST 请求（参数放查询字符串），返回 JSON。

    注意这个区别：报价接口只有把品种 id 放在"查询字符串"里才生效
    （网址问号后面的部分），放在请求体里会被网站忽略。
    这是从网站前端 JS 源码里确认的行为。
    """
    response = _request_with_retry("POST", url, params=params, json={})
    return _parse_json_response(response, url)


def decrypt_aes(raw):
    """解密报价接口返回的密文，得到 JSON 字符串。

    解密规则（来自网站前端 JS）：
      - 密文前 16 个字符 = IV（初始向量）
      - 剩下部分 = Base64 编码的 AES-CBC 密文
      - 填充方式为 PKCS7，密钥见 config.AES_KEY

    参数 raw：接口返回的密文字符串
    返回：解密后的明文字符串（内容是 JSON）
    """
    try:
        iv = raw[:16].encode("utf-8")                       # 前 16 个字符当 IV
        ciphertext = base64.b64decode(raw[16:])             # 其余部分 Base64 还原成字节
        cipher = AES.new(config.AES_KEY, AES.MODE_CBC, iv)  # 按 CBC 模式建解密器
        plain = unpad(cipher.decrypt(ciphertext), 16)       # 解密并去掉 PKCS7 填充
        return plain.decode("utf-8")
    except Exception as exc:
        # 解密失败说明网站可能改了加密方式，给出明确提示而不是崩溃
        raise FetchError(f"报价数据解密失败，网站可能更新了加密方式（{exc}）")


def fetch_variety_price(variety_id):
    """获取某个品种"今天"在各批发市场的报价。

    参数 variety_id：品种 id（从品种树里拿到的数字字符串）
    返回：字典 {"date": 数据日期, "markets": [市场名...], "prices": [价格...]}
         如果该品种今天没有市场报价，返回 None（这不算错误，属正常情况）。
    """
    payload = post_query(
        config.API_VARIETY_PRICE,
        {"marketIDs": "", "provinceCodes": "", "varietyID": variety_id},
    )

    # 今天没人报价的品种，网站会返回 data 为 null，属于正常业务情况
    if payload.get("data") is None:
        return None

    decrypted = json.loads(decrypt_aes(payload["data"]))
    return {
        "date": decrypted.get("date", ""),
        "markets": decrypted.get("x", []),
        "prices": decrypted.get("y", []),
    }

from __future__ import annotations

from datetime import datetime, timedelta, timezone
from email.utils import parsedate_to_datetime

import requests


CHINA_TZ = timezone(timedelta(hours=8), "Asia/Shanghai")
STARTUP_CUTOFF = datetime(2026, 6, 13, 0, 0, 0, tzinfo=CHINA_TZ)
NETWORK_TIME_URLS = (
    "https://www.baidu.com",
    "https://www.microsoft.com",
    "https://www.cloudflare.com",
)


class StartupBlocked(Exception):
    """Raised when temporary startup time limiting blocks the program."""


def parse_http_date(value: str) -> datetime:
    parsed = parsedate_to_datetime(value)
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=CHINA_TZ)
    return parsed


def get_network_time(timeout: float = 5.0) -> datetime | None:
    for url in NETWORK_TIME_URLS:
        try:
            response = requests.head(url, timeout=timeout, allow_redirects=True)
            header = response.headers.get("Date")
            if header:
                return parse_http_date(header)
        except requests.RequestException:
            continue
    return None


def enforce_startup_time_limit(
    *,
    now_fetcher=get_network_time,
    cutoff: datetime = STARTUP_CUTOFF,
) -> datetime:
    current_time = now_fetcher()
    if current_time is None:
        raise StartupBlocked("无法获取网络时间，程序不能启动。")

    current_time = current_time.astimezone(CHINA_TZ)
    cutoff = cutoff.astimezone(CHINA_TZ)
    if current_time >= cutoff:
        raise StartupBlocked("程序已超过临时使用期限，不能启动。")
    return current_time

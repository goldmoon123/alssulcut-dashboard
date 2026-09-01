import re
from datetime import datetime


def duration_to_seconds(duration):
    pattern = re.compile(
        r"PT"
        r"(?:(\d+)H)?"
        r"(?:(\d+)M)?"
        r"(?:(\d+)S)?"
    )

    match = pattern.fullmatch(duration)

    if not match:
        return 0

    hours = int(match.group(1) or 0)
    minutes = int(match.group(2) or 0)
    seconds = int(match.group(3) or 0)

    return hours * 3600 + minutes * 60 + seconds


def format_duration(seconds):
    minutes = seconds // 60
    remaining_seconds = seconds % 60

    return f"{minutes}:{remaining_seconds:02d}"


def format_watch_time(minutes):
    minutes = float(minutes or 0)

    if minutes >= 60:
        return f"{minutes / 60:,.1f}시간"

    return f"{minutes:,.0f}분"


def format_number(number):
    number = float(number or 0)

    if number >= 100000000:
        return f"{number / 100000000:.1f}억"

    if number >= 10000:
        return f"{number / 10000:.1f}만"

    return f"{number:,.0f}"


def parse_youtube_datetime(value):
    if not value:
        return None

    return datetime.fromisoformat(
        value.replace("Z", "+00:00")
    )


def format_youtube_datetime(value):
    dt = parse_youtube_datetime(value)

    if not dt:
        return "-"

    return dt.strftime("%Y-%m-%d %H:%M")
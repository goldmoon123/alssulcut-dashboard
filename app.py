import os
from datetime import date, datetime, timedelta, timezone
from zoneinfo import ZoneInfo
from io import BytesIO
from openpyxl.worksheet.table import Table, TableStyleInfo
import calendar
import re
import time
import requests

import pandas as pd
import streamlit as st

from dotenv import load_dotenv
from google_auth_oauthlib.flow import Flow
from google.oauth2.credentials import Credentials

from youtube_api import (
    create_youtube_api,
    get_channel_info,
    get_all_channel_videos,
    split_videos_by_status,
)

from analytics import (
    create_analytics_api,
    get_period_summary,
    get_daily_channel_data,
    get_video_performance_for_day,
    get_video_analytics,
    get_daily_video_analytics,
    get_previous_period,
    calculate_change,
)

from utils import (
    format_watch_time,
)

KST = ZoneInfo("Asia/Seoul")


# =========================================================
# 1. 기본 설정
# =========================================================

load_dotenv()


def get_config(name, default=None):
    value = os.getenv(name)
    if value:
        return value

    try:
        return st.secrets.get(name, default)
    except Exception:
        return default


CLIENT_ID = get_config("GOOGLE_CLIENT_ID")
CLIENT_SECRET = get_config("GOOGLE_CLIENT_SECRET")

# 로컬에서는 localhost, 배포된 Streamlit에서는 Secrets의 주소 사용
REDIRECT_URI = get_config(
    "GOOGLE_REDIRECT_URI",
    "http://localhost:8501",
)

# V6.5: 스냅샷 조회는 서버(Streamlit)에서만 수행합니다.
# Secret Key는 절대 코드/GitHub에 넣지 않고 Streamlit Secrets 또는 환경변수로만 주입합니다.
SUPABASE_URL = get_config("SUPABASE_URL")
SUPABASE_SECRET_KEY = (
    get_config("SUPABASE_SECRET_KEY")
    or get_config("SUPABASE_SERVICE_ROLE_KEY")
)



def _supabase_rest_headers():
    if not SUPABASE_URL or not SUPABASE_SECRET_KEY:
        return None
    return {
        "apikey": SUPABASE_SECRET_KEY,
        "Authorization": f"Bearer {SUPABASE_SECRET_KEY}",
        "Content-Type": "application/json",
        "Accept": "application/json",
    }


def _supabase_table_get(table_name, params):
    headers = _supabase_rest_headers()
    if not headers:
        return {"ok": False, "reason": "not_configured", "rows": []}

    try:
        response = requests.get(
            f"{SUPABASE_URL.rstrip('/')}/rest/v1/{table_name}",
            headers=headers,
            params=params,
            timeout=15,
        )
        if response.status_code not in (200, 206):
            return {
                "ok": False,
                "reason": "http_error",
                "status": response.status_code,
                "rows": [],
            }
        rows = response.json()
        return {
            "ok": True,
            "reason": None,
            "rows": rows if isinstance(rows, list) else [],
        }
    except Exception:
        return {"ok": False, "reason": "request_error", "rows": []}


def _supabase_table_upsert(table_name, payload, on_conflict):
    headers = _supabase_rest_headers()
    if not headers:
        return {"ok": False, "reason": "not_configured"}

    try:
        response = requests.post(
            f"{SUPABASE_URL.rstrip('/')}/rest/v1/{table_name}",
            headers={
                **headers,
                "Prefer": "resolution=merge-duplicates,return=representation",
            },
            params={"on_conflict": on_conflict},
            json=payload,
            timeout=15,
        )
        if response.status_code not in (200, 201):
            return {
                "ok": False,
                "reason": "http_error",
                "status": response.status_code,
                "message": response.text[:500],
            }
        return {
            "ok": True,
            "reason": None,
            "rows": response.json() if response.text else [],
        }
    except Exception as exc:
        return {
            "ok": False,
            "reason": "request_error",
            "message": str(exc),
        }


def _supabase_table_insert(table_name, payload):
    headers = _supabase_rest_headers()
    if not headers:
        return {"ok": False, "reason": "not_configured"}

    try:
        response = requests.post(
            f"{SUPABASE_URL.rstrip('/')}/rest/v1/{table_name}",
            headers={**headers, "Prefer": "return=representation"},
            json=payload,
            timeout=15,
        )
        if response.status_code not in (200, 201):
            return {
                "ok": False,
                "reason": "http_error",
                "status": response.status_code,
                "message": response.text[:500],
            }
        return {
            "ok": True,
            "reason": None,
            "rows": response.json() if response.text else [],
        }
    except Exception as exc:
        return {
            "ok": False,
            "reason": "request_error",
            "message": str(exc),
        }


def _supabase_table_update(table_name, filters, payload):
    headers = _supabase_rest_headers()
    if not headers:
        return {"ok": False, "reason": "not_configured"}

    try:
        response = requests.patch(
            f"{SUPABASE_URL.rstrip('/')}/rest/v1/{table_name}",
            headers={**headers, "Prefer": "return=representation"},
            params=filters,
            json=payload,
            timeout=15,
        )
        if response.status_code not in (200, 204):
            return {
                "ok": False,
                "reason": "http_error",
                "status": response.status_code,
                "message": response.text[:500],
            }
        return {
            "ok": True,
            "reason": None,
            "rows": response.json() if response.text else [],
        }
    except Exception as exc:
        return {
            "ok": False,
            "reason": "request_error",
            "message": str(exc),
        }


def _supabase_table_delete(table_name, filters):
    headers = _supabase_rest_headers()
    if not headers:
        return {"ok": False, "reason": "not_configured"}

    try:
        response = requests.delete(
            f"{SUPABASE_URL.rstrip('/')}/rest/v1/{table_name}",
            headers={**headers, "Prefer": "return=minimal"},
            params=filters,
            timeout=15,
        )
        if response.status_code not in (200, 204):
            return {
                "ok": False,
                "reason": "http_error",
                "status": response.status_code,
                "message": response.text[:500],
            }
        return {"ok": True, "reason": None}
    except Exception as exc:
        return {
            "ok": False,
            "reason": "request_error",
            "message": str(exc),
        }



def _snapshot_point_near(rows, target_dt, tolerance_minutes=90, prefer_before=True):
    """
    기준 시점에 가장 가까운 스냅샷을 찾습니다.
    - 기본은 target 이전 스냅샷 우선
    - 허용 오차를 넘으면 None
    """
    if not rows:
        return None

    tolerance = timedelta(minutes=tolerance_minutes)

    if prefer_before:
        before = [r for r in rows if r["captured_at"] <= target_dt]
        if before:
            candidate = before[-1]
            if (target_dt - candidate["captured_at"]) <= tolerance:
                return candidate

    candidate = min(
        rows,
        key=lambda r: abs((r["captured_at"] - target_dt).total_seconds()),
    )
    if abs(candidate["captured_at"] - target_dt) <= tolerance:
        return candidate
    return None


def _snapshot_age_milestones(video, rows, now_utc=None):
    """
    V7 — 업로드 후 기준 시점별(1h/6h/12h/24h/48h/72h) 실제 스냅샷 조회수.
    데이터가 없으면 0이 아니라 None으로 둡니다.
    """
    now_utc = now_utc or datetime.now(timezone.utc)
    published = _parse_utc(video.get("published_raw"))
    if not published or not rows:
        return {}

    milestones = [1, 6, 12, 24, 48, 72]
    result = {}

    for hour in milestones:
        target = published + timedelta(hours=hour)

        # 아직 해당 나이에 도달하지 않은 영상은 미도달
        if now_utc < target:
            result[hour] = {
                "status": "not_reached",
                "view_count": None,
                "captured_at": None,
                "delay_minutes": None,
            }
            continue

        point = _snapshot_point_near(
            rows,
            target,
            tolerance_minutes=90,
            prefer_before=True,
        )
        if point is None:
            result[hour] = {
                "status": "missing",
                "view_count": None,
                "captured_at": None,
                "delay_minutes": None,
            }
            continue

        delay = abs((point["captured_at"] - target).total_seconds()) / 60
        result[hour] = {
            "status": "ok",
            "view_count": int(point.get("view_count") or 0),
            "captured_at": point["captured_at"],
            "delay_minutes": round(delay, 1),
        }

    return result


def _snapshot_interval_gain(video, rows, start_hour, end_hour, now_utc=None):
    """
    업로드 후 start_hour~end_hour 사이 증가량.
    경계 스냅샷이 없으면 None.
    """
    now_utc = now_utc or datetime.now(timezone.utc)
    published = _parse_utc(video.get("published_raw"))
    if not published or not rows:
        return None

    start_target = published + timedelta(hours=start_hour)
    end_target = published + timedelta(hours=end_hour)

    if now_utc < end_target:
        return None

    start_point = _snapshot_point_near(rows, start_target, 90, True)
    end_point = _snapshot_point_near(rows, end_target, 90, True)
    if not start_point or not end_point:
        return None

    gain = int(end_point.get("view_count") or 0) - int(start_point.get("view_count") or 0)
    if gain < 0:
        return {
            "gain": None,
            "status": "counter_adjusted",
            "start": start_point,
            "end": end_point,
        }

    return {
        "gain": gain,
        "status": "ok",
        "start": start_point,
        "end": end_point,
    }


def _snapshot_milestone_table(video, rows, now_utc=None):
    milestone_data = _snapshot_age_milestones(video, rows, now_utc)
    rows_out = []

    labels = {
        1: "1시간",
        6: "6시간",
        12: "12시간",
        24: "24시간",
        48: "48시간",
        72: "72시간",
    }

    for hour in [1, 6, 12, 24, 48, 72]:
        item = milestone_data.get(hour, {})
        status = item.get("status")

        if status == "ok":
            value = f"{int(item['view_count']):,}회"
            note = f"기준시점 ±{item.get('delay_minutes', 0):.0f}분"
        elif status == "not_reached":
            value = "⏳ 미도달"
            note = "아직 해당 업로드 나이에 도달하지 않음"
        else:
            value = "데이터 부족"
            note = "기준시점 근처 스냅샷 없음"

        rows_out.append({
            "기준 시점": labels[hour],
            "누적 조회수": value,
            "상태": note,
        })

    return rows_out


def _connected_youtube_channel_id(youtube_client):
    """현재 OAuth로 연결된 YouTube 채널 ID를 직접 확인합니다."""
    response = youtube_client.channels().list(
        part="id",
        mine=True,
        maxResults=1,
    ).execute()
    items = response.get("items", [])
    if not items:
        return None
    return items[0].get("id")


@st.cache_data(ttl=300, show_spinner=False)
def _fetch_channel_snapshots(channel_id, since_iso):
    """
    현재 OAuth 채널의 최근 스냅샷을 Supabase REST API에서 읽습니다.
    - 서버 전용 Secret Key 사용
    - channel_id는 사용자가 입력하지 않고 YouTube OAuth에서 얻은 값만 사용
    - 1000행 단위로 페이지 처리
    """
    if not SUPABASE_URL or not SUPABASE_SECRET_KEY:
        return {
            "ok": False,
            "reason": "not_configured",
            "rows": [],
            "message": "Supabase 읽기 설정이 아직 연결되지 않았습니다.",
        }

    if not channel_id:
        return {
            "ok": False,
            "reason": "channel_missing",
            "rows": [],
            "message": "현재 연결된 YouTube 채널 ID를 확인하지 못했습니다.",
        }

    url = f"{SUPABASE_URL.rstrip('/')}/rest/v1/video_snapshots"
    headers = {
        "apikey": SUPABASE_SECRET_KEY,
        "Accept": "application/json",
    }
    params = {
        "select": "video_id,captured_at,view_count,like_count,comment_count",
        "channel_id": f"eq.{channel_id}",
        "captured_at": f"gte.{since_iso}",
        "order": "captured_at.asc",
    }

    page_size = 1000
    offset = 0
    rows = []
    max_rows = 30000

    try:
        while offset < max_rows:
            page_headers = {
                **headers,
                "Range-Unit": "items",
                "Range": f"{offset}-{offset + page_size - 1}",
            }
            response = requests.get(
                url,
                headers=page_headers,
                params=params,
                timeout=15,
            )
            if response.status_code not in (200, 206):
                return {
                    "ok": False,
                    "reason": "http_error",
                    "rows": [],
                    "message": f"Supabase 스냅샷 조회 실패 (HTTP {response.status_code})",
                }

            page = response.json()
            if not isinstance(page, list):
                return {
                    "ok": False,
                    "reason": "invalid_response",
                    "rows": [],
                    "message": "Supabase 응답 형식을 확인하지 못했습니다.",
                }

            rows.extend(page)
            if len(page) < page_size:
                break
            offset += page_size

        return {
            "ok": True,
            "reason": None,
            "rows": rows,
            "message": None,
            "truncated": len(rows) >= max_rows,
        }
    except Exception:
        return {
            "ok": False,
            "reason": "request_error",
            "rows": [],
            "message": "Supabase 스냅샷을 불러오는 중 연결 오류가 발생했습니다.",
        }


def _parse_utc(value):
    if not value:
        return None
    try:
        dt = datetime.fromisoformat(str(value).replace("Z", "+00:00"))
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt.astimezone(timezone.utc)
    except Exception:
        return None


def _group_snapshot_rows(rows, allowed_video_ids=None):
    allowed = set(allowed_video_ids or [])
    grouped = {}
    for row in rows:
        video_id = row.get("video_id")
        if not video_id or (allowed and video_id not in allowed):
            continue
        captured_at = _parse_utc(row.get("captured_at"))
        if not captured_at:
            continue
        try:
            view_count = int(row.get("view_count") or 0)
            like_count = int(row.get("like_count") or 0)
            comment_count = int(row.get("comment_count") or 0)
        except Exception:
            continue
        grouped.setdefault(video_id, []).append({
            "captured_at": captured_at,
            "view_count": view_count,
            "like_count": like_count,
            "comment_count": comment_count,
        })

    for video_id in grouped:
        # 같은 시각 중복이 있으면 마지막 값만 사용
        dedup = {}
        for row in grouped[video_id]:
            dedup[row["captured_at"]] = row
        grouped[video_id] = sorted(dedup.values(), key=lambda x: x["captured_at"])
    return grouped


def _snapshot_window_hours(video, now_utc):
    """영상 나이에 따라 비교 구간을 자동 선택합니다."""
    published = _parse_utc(video.get("published_raw"))
    if not published:
        return 3
    age_hours = max((now_utc - published).total_seconds() / 3600, 0)
    if age_hours < 6:
        return 1
    if age_hours < 48:
        return 3
    if age_hours < 24 * 7:
        return 6
    return 24


def _nearest_snapshot_at_or_before(rows, target, max_gap_hours):
    candidate = None
    for row in rows:
        if row["captured_at"] <= target:
            candidate = row
        else:
            break
    if candidate is None:
        return None
    gap_hours = (target - candidate["captured_at"]).total_seconds() / 3600
    if gap_hours > max_gap_hours:
        return None
    return candidate


def _snapshot_activity_floor(current_views):
    """
    아주 작은 조회수 변화가 급격한 비율 변화로 과대평가되지 않게 하는 최소 활동량.
    채널 공통 고정값 하나가 아니라 현재 누적 조회수에 따라 완만하게 조정합니다.
    """
    current_views = max(int(current_views or 0), 0)
    return max(3, min(50, int(round(current_views * 0.0001))))


def _snapshot_growth_state(video, snapshot_rows, now_utc=None):
    """
    V6.5-1 기본 성장상태.
    '현재 속도 vs 직전 속도'만 판정하며 급상승/재상승은 별도 단계에서 처리합니다.
    """
    now_utc = now_utc or datetime.now(timezone.utc)
    rows = snapshot_rows or []
    window_hours = _snapshot_window_hours(video, now_utc)

    base = {
        "state": "⏳ 데이터 축적 중",
        "window_hours": window_hours,
        "recent_gain": None,
        "previous_gain": None,
        "recent_velocity": None,
        "previous_velocity": None,
        "ratio": None,
        "confidence": "낮음",
        "latest_at": None,
        "note": None,
    }

    if len(rows) < 3:
        return base

    latest = rows[-1]
    base["latest_at"] = latest["captured_at"]
    freshness_hours = (now_utc - latest["captured_at"]).total_seconds() / 3600
    if freshness_hours > 8:
        base["state"] = "⏳ 스냅샷 갱신 대기"
        base["note"] = f"마지막 스냅샷이 약 {freshness_hours:.1f}시간 전입니다."
        return base

    recent_target = latest["captured_at"] - timedelta(hours=window_hours)
    previous_target = latest["captured_at"] - timedelta(hours=window_hours * 2)
    max_gap_hours = min(max(window_hours * 0.5, 0.75), 8)

    recent_start = _nearest_snapshot_at_or_before(rows, recent_target, max_gap_hours)
    previous_start = _nearest_snapshot_at_or_before(rows, previous_target, max_gap_hours)
    if not recent_start or not previous_start:
        return base

    recent_hours_actual = (latest["captured_at"] - recent_start["captured_at"]).total_seconds() / 3600
    previous_hours_actual = (recent_start["captured_at"] - previous_start["captured_at"]).total_seconds() / 3600
    if recent_hours_actual <= 0 or previous_hours_actual <= 0:
        return base

    recent_gain = latest["view_count"] - recent_start["view_count"]
    previous_gain = recent_start["view_count"] - previous_start["view_count"]
    recent_velocity = recent_gain / recent_hours_actual
    previous_velocity = previous_gain / previous_hours_actual

    base.update({
        "recent_gain": recent_gain,
        "previous_gain": previous_gain,
        "recent_velocity": recent_velocity,
        "previous_velocity": previous_velocity,
    })

    # YouTube가 조회수를 사후 보정해 누적값이 감소하는 드문 경우는 일반 성장판정에서 분리합니다.
    if recent_gain < 0 or previous_gain < 0:
        base["state"] = "→ 유지"
        base["note"] = "조회수 누적값 보정이 감지되어 속도 판정을 보류했습니다."
        return base

    current_views = max(int(latest["view_count"]), 0)
    # 아주 작은 변화(예: 2→4)를 상승으로 과대평가하지 않기 위한 최소 활동량.
    activity_floor = _snapshot_activity_floor(current_views)
    if recent_gain + previous_gain <= activity_floor:
        base["state"] = "💤 정체"
    elif previous_velocity <= 0:
        base["state"] = "↗ 상승" if recent_velocity > 0 else "💤 정체"
    else:
        ratio = recent_velocity / previous_velocity
        base["ratio"] = ratio
        if ratio >= 1.35:
            base["state"] = "↗ 상승"
        elif ratio <= 0.65:
            base["state"] = "↘ 하락"
        else:
            base["state"] = "→ 유지"

    # 판정 신뢰도: 시간 경계와 가까운 스냅샷을 확보했는지 + 최신성 기준
    recent_gap = abs((recent_target - recent_start["captured_at"]).total_seconds()) / 3600
    previous_gap = abs((previous_target - previous_start["captured_at"]).total_seconds()) / 3600
    boundary_quality = max(recent_gap, previous_gap) / max(window_hours, 1)
    if freshness_hours <= 1.5 and boundary_quality <= 0.25 and len(rows) >= 5:
        base["confidence"] = "높음"
    elif freshness_hours <= 8 and boundary_quality <= 0.5:
        base["confidence"] = "보통"
    else:
        base["confidence"] = "낮음"

    return base


def _snapshot_special_event(video, snapshot_state, snapshot_rows, now_utc=None):
    """
    V6.5-2 특별 성장 이벤트 V1.
    기본 성장상태와 분리해서 🚀 급상승 / 🔥 재상승만 감지합니다.

    원칙
    - 비율 하나만으로 판정하지 않음
    - 절대 증가량(활동량 기준)을 함께 확인
    - 신뢰도 낮음이면 특별 이벤트를 확정하지 않음
    - 오래된 영상의 재상승은 직전 구간이 거의 정체였는지 추가 확인
    """
    result = {
        "event": None,
        "label": None,
        "confidence": "낮음",
        "reason": None,
    }

    if not snapshot_state or snapshot_state.get("recent_gain") is None:
        return result

    if snapshot_state.get("confidence") == "낮음":
        result["reason"] = "비교 구간의 스냅샷이 아직 충분하지 않습니다."
        return result

    rows = snapshot_rows or []
    if not rows:
        return result

    now_utc = now_utc or datetime.now(timezone.utc)
    published = _parse_utc(video.get("published_raw"))
    age_hours = (
        max((now_utc - published).total_seconds() / 3600, 0)
        if published else 0
    )

    recent_gain = max(int(snapshot_state.get("recent_gain") or 0), 0)
    previous_gain = max(int(snapshot_state.get("previous_gain") or 0), 0)
    recent_velocity = max(float(snapshot_state.get("recent_velocity") or 0), 0.0)
    previous_velocity = max(float(snapshot_state.get("previous_velocity") or 0), 0.0)
    ratio = snapshot_state.get("ratio")
    window_hours = int(snapshot_state.get("window_hours") or 1)
    current_views = max(int(rows[-1].get("view_count") or 0), 0)

    activity_floor = _snapshot_activity_floor(current_views)
    meaningful_gain = max(activity_floor * 4, 10)

    # 🔥 재상승: 업로드 후 3일 이상 지난 영상이 직전 구간에는 거의 멈췄다가
    # 최근 구간에서 의미 있는 증가로 다시 살아난 경우.
    if age_hours >= 72:
        was_quiet = previous_gain <= activity_floor
        woke_up = recent_gain >= meaningful_gain
        velocity_jump = (
            recent_velocity > 0
            if previous_velocity <= 0
            else recent_velocity >= previous_velocity * 3.0
        )
        if was_quiet and woke_up and velocity_jump:
            result.update({
                "event": "resurge",
                "label": "🔥 재상승",
                "confidence": snapshot_state.get("confidence", "보통"),
                "reason": (
                    f"직전 {window_hours}시간 +{previous_gain:,}회에서 "
                    f"최근 {window_hours}시간 +{recent_gain:,}회로 다시 증가"
                ),
            })
            return result

    # 🚀 급상승: 현재 기본상태가 '상승'이고,
    # 직전 대비 속도 2배 이상 + 최소 활동량을 충분히 넘긴 경우.
    if snapshot_state.get("state") == "↗ 상승":
        strong_ratio = ratio is not None and ratio >= 2.0
        strong_gain = recent_gain >= meaningful_gain
        if strong_ratio and strong_gain:
            result.update({
                "event": "surge",
                "label": "🚀 급상승",
                "confidence": snapshot_state.get("confidence", "보통"),
                "reason": (
                    f"최근 {window_hours}시간 속도가 직전 구간의 {ratio:.2f}배 · "
                    f"최근 +{recent_gain:,}회"
                ),
            })

    return result


# HTTP 허용은 localhost 개발 때만 사용
if REDIRECT_URI.startswith("http://localhost"):
    os.environ["OAUTHLIB_INSECURE_TRANSPORT"] = "1"
else:
    os.environ.pop("OAUTHLIB_INSECURE_TRANSPORT", None)

SCOPES = [
    "https://www.googleapis.com/auth/youtube.readonly",
    "https://www.googleapis.com/auth/yt-analytics.readonly",
]

st.set_page_config(
    page_title="Shorts Scope",
    page_icon="📊",
    layout="wide",
)

st.markdown("""
<style>
@media (max-width: 700px) {
    div[data-testid="stHorizontalBlock"] > div {
        min-width: 0;
    }
}
</style>
""", unsafe_allow_html=True)

# =========================================================
# 모바일 화면 최적화
# =========================================================
st.markdown(
    """
    <style>
    @media (max-width: 768px) {
        .block-container {
            padding-top: 1.1rem !important;
            padding-left: 0.85rem !important;
            padding-right: 0.85rem !important;
            padding-bottom: 4rem !important;
        }

        h1 { font-size: 2rem !important; line-height: 1.18 !important; }
        h2 { font-size: 1.55rem !important; line-height: 1.2 !important; }
        h3 { font-size: 1.25rem !important; line-height: 1.2 !important; }

        [data-testid="stMetric"] {
            width: 100% !important;
            min-width: 0 !important;
            padding: 0.65rem 0.75rem !important;
            border: 1px solid rgba(128,128,128,.18);
            border-radius: 12px;
            min-height: 96px;
        }
        [data-testid="stMetricLabel"] {
            width: 100% !important;
            min-width: 0 !important;
            font-size: 0.82rem !important;
        }
        [data-testid="stMetricValue"] {
            width: 100% !important;
            min-width: 0 !important;
            font-size: 1.65rem !important;
        }
        [data-testid="stMetricDelta"] { font-size: 0.75rem !important; }

        /* 모바일에서는 metric 묶음을 1열로 쌓아 제목이 세로로 찢어지지 않게 함 */
        [data-testid="stHorizontalBlock"]:has([data-testid="stMetric"]) {
            display: grid !important;
            grid-template-columns: 1fr !important;
            gap: 0.65rem !important;
            width: 100% !important;
        }

        [data-testid="stHorizontalBlock"]:has([data-testid="stMetric"]) > div[data-testid="column"] {
            width: 100% !important;
            min-width: 0 !important;
            max-width: 100% !important;
            flex: none !important;
        }

        /* 달력: 7칸을 무조건 한 줄에 유지 */
        [data-testid="stHorizontalBlock"]:has(> div:nth-child(7)):not(:has(> div:nth-child(8))) {
            display: grid !important;
            grid-template-columns: repeat(7, minmax(0, 1fr)) !important;
            gap: 3px !important;
        }
        [data-testid="stHorizontalBlock"]:has(> div:nth-child(7)):not(:has(> div:nth-child(8))) > div {
            min-width: 0 !important;
            width: auto !important;
        }
        [data-testid="stHorizontalBlock"]:has(> div:nth-child(7)):not(:has(> div:nth-child(8))) button {
            min-height: 58px !important;
            padding: 3px 1px !important;
            border-radius: 7px !important;
        }
        [data-testid="stHorizontalBlock"]:has(> div:nth-child(7)):not(:has(> div:nth-child(8))) button p {
            font-size: 0.64rem !important;
            line-height: 1.15 !important;
            white-space: pre-line !important;
        }

        /* 달력 이전/현재/다음 3칸은 가로 유지 */
        [data-testid="stHorizontalBlock"]:has(> div:nth-child(3)):not(:has(> div:nth-child(4))) {
            flex-wrap: nowrap !important;
            gap: 0.35rem !important;
        }
        [data-testid="stHorizontalBlock"]:has(> div:nth-child(3)):not(:has(> div:nth-child(4))) > div {
            min-width: 0 !important;
        }

        .stButton > button {
            font-size: 0.82rem;
        }

        [data-testid="stVegaLiteChart"],
        [data-testid="stArrowVegaLiteChart"] {
            margin-top: -0.25rem !important;
            margin-bottom: 0.5rem !important;
        }

        hr { margin: 0.85rem 0 !important; }
        p { line-height: 1.42; }
        .stCaption, [data-testid="stCaptionContainer"] { font-size: 0.82rem !important; }

        /* 모바일에서 섹션을 조금 더 촘촘하게 */
        [data-testid="stVerticalBlock"] { gap: 0.65rem !important; }

        [data-testid="stImage"] img { border-radius: 10px !important; }
        [data-testid="stDataFrame"] { font-size: 0.78rem !important; }
        [data-baseweb="tab-list"] { gap: 0.15rem !important; }
        [data-baseweb="tab"] { padding-left: 0.45rem !important; padding-right: 0.45rem !important; }


        /* metric 제목이 ... 으로 잘리지 않도록 */
        [data-testid="stMetricLabel"] p {
            white-space: normal !important;
            overflow: visible !important;
            text-overflow: clip !important;
            line-height: 1.15 !important;
        }

        /* 긴 metric 값이 ... 으로 잘리지 않도록 */
        [data-testid="stMetricValue"] > div {
            overflow: visible !important;
            text-overflow: clip !important;
            white-space: nowrap !important;
        }

        /* 3개짜리 핵심 지표는 모바일에서 2열 + 마지막 한 칸 전체 폭 */
        [data-testid="stHorizontalBlock"]:has([data-testid="stMetric"]):has(> div:nth-child(3)):not(:has(> div:nth-child(4))) {
            display: grid !important;
            grid-template-columns: 1fr !important;
            gap: 0.65rem !important;
            width: 100% !important;
        }
        [data-testid="stHorizontalBlock"]:has([data-testid="stMetric"]):has(> div:nth-child(3)):not(:has(> div:nth-child(4))) > div:last-child {
            grid-column: 1 / -1;
        }

        /* 긴 표/확장영역은 화면 밖으로 삐져나오지 않게 */
        [data-testid="stDataFrame"],
        [data-testid="stExpander"] {
            max-width: 100% !important;
            overflow-x: auto !important;
        }

        /* 메뉴/라디오 문구가 모바일에서 한 줄 강제로 잘리지 않게 */
        div[role="radiogroup"] label p {
            white-space: normal !important;
            line-height: 1.25 !important;
        }
    }

    .status-two-grid {
        display: grid;
        grid-template-columns: repeat(2, minmax(0, 1fr));
        gap: 0.75rem;
        margin-top: 0.6rem;
    }
    .status-card {
        border: 1px solid rgba(128,128,128,.20);
        border-radius: 12px;
        padding: 0.8rem 0.9rem;
        min-height: 88px;
    }
    .status-label {
        font-size: 0.95rem;
        margin-bottom: 0.35rem;
    }
    .status-value {
        font-size: 1.8rem;
        line-height: 1.1;
    }
    @media (max-width: 768px) {
        .status-two-grid { gap: 0.55rem; }
        .status-card {
            min-height: 92px;
            padding: 0.65rem 0.7rem;
        }
        .status-label { font-size: 0.82rem; }
        .status-value { font-size: 1.65rem; }
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# =========================================================
# V7.3 — Modern SaaS UI Design System
# 기능 로직은 건드리지 않고 전체 시각 체계만 통일합니다.
# =========================================================
st.markdown(
    """
    <style>
    :root {
        --ss-bg: #f6f7fb;
        --ss-surface: #ffffff;
        --ss-surface-soft: #fafbfe;
        --ss-text: #171923;
        --ss-muted: #667085;
        --ss-border: #e7e9f0;
        --ss-primary: #ff4d5a;
        --ss-primary-dark: #e83e4b;
        --ss-shadow: 0 8px 24px rgba(20, 24, 35, 0.06);
        --ss-radius: 16px;
        --ss-radius-sm: 11px;
    }

    html, body, [class*="css"] {
        font-family:
            Inter, Pretendard, -apple-system, BlinkMacSystemFont,
            "Segoe UI", "Noto Sans KR", sans-serif;
    }

    .stApp {
        background: var(--ss-bg);
        color: var(--ss-text);
    }

    .block-container {
        max-width: 1440px;
        padding-top: 2.1rem;
        padding-bottom: 5rem;
    }

    /* 제목 */
    h1, h2, h3 {
        color: var(--ss-text) !important;
        letter-spacing: -0.025em;
    }
    h1 {
        font-size: 2.25rem !important;
        font-weight: 800 !important;
        line-height: 1.15 !important;
    }
    h2 {
        font-size: 1.65rem !important;
        font-weight: 750 !important;
        margin-top: 1.35rem !important;
    }
    h3 {
        font-size: 1.22rem !important;
        font-weight: 700 !important;
    }

    .stCaption,
    [data-testid="stCaptionContainer"],
    [data-testid="stCaptionContainer"] p {
        color: var(--ss-muted) !important;
        line-height: 1.5 !important;
    }

    /* 상단 브랜드 카드 */
    .ss-hero {
        background:
            radial-gradient(circle at 90% 15%, rgba(255,77,90,.15), transparent 34%),
            linear-gradient(135deg, #ffffff 0%, #fbfbfe 100%);
        border: 1px solid var(--ss-border);
        border-radius: 22px;
        padding: 1.55rem 1.7rem;
        box-shadow: var(--ss-shadow);
        margin-bottom: 1rem;
    }
    .ss-brand-row {
        display: flex;
        align-items: center;
        gap: .7rem;
        margin-bottom: .45rem;
    }
    .ss-logo {
        width: 42px;
        height: 42px;
        display: grid;
        place-items: center;
        border-radius: 13px;
        background: linear-gradient(135deg, #ff4d5a, #ff7a63);
        color: white;
        font-size: 1.2rem;
        box-shadow: 0 7px 18px rgba(255,77,90,.22);
    }
    .ss-brand {
        font-size: 1.42rem;
        font-weight: 800;
        letter-spacing: -.03em;
    }
    .ss-hero-title {
        font-size: 1.08rem;
        font-weight: 650;
        color: #353a49;
        margin-top: .2rem;
    }
    .ss-hero-sub {
        color: var(--ss-muted);
        font-size: .92rem;
        margin-top: .35rem;
    }
    .ss-chip-row {
        display: flex;
        flex-wrap: wrap;
        gap: .45rem;
        margin-top: .9rem;
    }
    .ss-chip {
        padding: .34rem .62rem;
        border: 1px solid var(--ss-border);
        border-radius: 999px;
        background: rgba(255,255,255,.86);
        color: #444b5f;
        font-size: .78rem;
        font-weight: 600;
    }

    /* Metric 카드 */
    [data-testid="stMetric"] {
        background: var(--ss-surface);
        border: 1px solid var(--ss-border);
        border-radius: var(--ss-radius);
        padding: 1rem 1.05rem !important;
        min-height: 112px;
        box-shadow: 0 4px 14px rgba(20,24,35,.035);
        transition: transform .15s ease, box-shadow .15s ease;
    }
    [data-testid="stMetric"]:hover {
        transform: translateY(-1px);
        box-shadow: 0 8px 22px rgba(20,24,35,.07);
    }
    [data-testid="stMetricLabel"] p {
        color: var(--ss-muted) !important;
        font-size: .84rem !important;
        font-weight: 650 !important;
    }
    [data-testid="stMetricValue"] {
        color: var(--ss-text) !important;
        font-weight: 780 !important;
    }

    /* 버튼 */
    .stButton > button,
    [data-testid="stFormSubmitButton"] > button,
    .stDownloadButton > button,
    [data-testid="stLinkButton"] a {
        min-height: 42px;
        border-radius: 11px !important;
        border: 1px solid #dfe2ea !important;
        font-weight: 700 !important;
        box-shadow: none !important;
        transition: transform .12s ease, box-shadow .12s ease, border-color .12s ease;
    }
    .stButton > button:hover,
    [data-testid="stFormSubmitButton"] > button:hover,
    .stDownloadButton > button:hover,
    [data-testid="stLinkButton"] a:hover {
        transform: translateY(-1px);
        border-color: #cfd3dd !important;
        box-shadow: 0 6px 14px rgba(20,24,35,.07) !important;
    }
    button[kind="primary"],
    [data-testid="stFormSubmitButton"] button[kind="primary"] {
        background: linear-gradient(135deg, var(--ss-primary), #ff6a62) !important;
        border-color: transparent !important;
        color: white !important;
    }

    /* 입력 */
    [data-baseweb="input"] > div,
    [data-baseweb="textarea"] > div,
    [data-baseweb="select"] > div,
    [data-baseweb="base-input"] {
        background: var(--ss-surface) !important;
        border-radius: 11px !important;
        border-color: #dfe2ea !important;
    }
    [data-baseweb="select"] > div:hover,
    [data-baseweb="input"] > div:hover {
        border-color: #c7cbd6 !important;
    }

    /* 라디오 메뉴: 탭형 네비게이션 */
    div[role="radiogroup"] {
        gap: .45rem !important;
        flex-wrap: wrap !important;
    }
    div[role="radiogroup"] label {
        background: var(--ss-surface);
        border: 1px solid var(--ss-border);
        border-radius: 999px;
        padding: .2rem .35rem;
        min-height: 38px;
        transition: background .12s ease, border-color .12s ease;
    }
    div[role="radiogroup"] label:hover {
        border-color: #cfd3dd;
        background: #fbfbfd;
    }
    div[role="radiogroup"] label p {
        font-size: .88rem !important;
        font-weight: 700 !important;
    }

    /* 탭 */
    [data-baseweb="tab-list"] {
        background: #eef0f5;
        padding: .28rem;
        border-radius: 12px;
        gap: .2rem !important;
    }
    [data-baseweb="tab"] {
        border-radius: 9px !important;
        min-height: 38px;
        font-weight: 700;
    }
    [aria-selected="true"][data-baseweb="tab"] {
        background: white !important;
        box-shadow: 0 2px 7px rgba(20,24,35,.07);
    }

    /* 알림 박스 */
    [data-testid="stAlert"] {
        border-radius: 13px !important;
        border-width: 1px !important;
        box-shadow: none !important;
    }

    /* Expander */
    [data-testid="stExpander"] {
        background: var(--ss-surface);
        border: 1px solid var(--ss-border) !important;
        border-radius: 14px !important;
        overflow: hidden;
    }
    [data-testid="stExpander"] summary {
        font-weight: 700 !important;
    }

    /* 표 */
    [data-testid="stDataFrame"] {
        background: var(--ss-surface);
        border: 1px solid var(--ss-border);
        border-radius: 14px;
        overflow: hidden;
        box-shadow: 0 3px 12px rgba(20,24,35,.025);
    }

    /* form 자체를 검색 패널처럼 */
    [data-testid="stForm"] {
        background: rgba(255,255,255,.72);
        border: 1px solid var(--ss-border);
        border-radius: 15px;
        padding: .9rem .95rem .95rem;
        box-shadow: 0 3px 12px rgba(20,24,35,.025);
    }

    /* 이미지 */
    [data-testid="stImage"] img {
        border-radius: 13px !important;
        border: 1px solid rgba(20,24,35,.06);
    }

    /* divider */
    hr {
        border: 0 !important;
        border-top: 1px solid var(--ss-border) !important;
        margin: 1.5rem 0 !important;
    }

    /* 상태 카드 */
    .status-two-grid {
        gap: .8rem !important;
    }
    .status-card {
        background: var(--ss-surface);
        border: 1px solid var(--ss-border) !important;
        border-radius: 15px !important;
        box-shadow: 0 3px 12px rgba(20,24,35,.03);
    }
    .status-label {
        color: var(--ss-muted);
        font-weight: 650;
    }
    .status-value {
        color: var(--ss-text);
        font-weight: 780;
    }

    /* progress */
    [data-testid="stProgress"] > div > div > div {
        border-radius: 999px !important;
    }

    /* 작은 화면 */
    @media (max-width: 768px) {
        .block-container {
            padding-top: .9rem !important;
            padding-left: .75rem !important;
            padding-right: .75rem !important;
        }
        .ss-hero {
            border-radius: 17px;
            padding: 1rem 1rem;
        }
        .ss-logo {
            width: 36px;
            height: 36px;
            border-radius: 11px;
        }
        .ss-brand {
            font-size: 1.18rem;
        }
        .ss-hero-title {
            font-size: .94rem;
        }
        .ss-hero-sub {
            font-size: .82rem;
        }
        [data-testid="stMetric"] {
            min-height: 94px;
            padding: .75rem .8rem !important;
            border-radius: 13px;
        }
        [data-testid="stForm"] {
            padding: .7rem;
            border-radius: 13px;
        }
        div[role="radiogroup"] label {
            min-height: 34px;
        }
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# V7.4 — Product UI shell
# =========================================================
st.markdown(
    """
    <style>
    :root {
        --p-bg:#f4f6f8; --p-card:#fff; --p-text:#111827; --p-muted:#6b7280;
        --p-line:#e5e7eb; --p-accent:#ff4d5a; --p-sidebar:#111318;
    }
    .stApp { background:var(--p-bg)!important; }
    .block-container { max-width:1320px!important; padding-top:1.25rem!important; padding-bottom:5rem!important; }
    [data-testid="stHeader"] { background:transparent!important; }

    section[data-testid="stSidebar"] {
        background:var(--p-sidebar)!important;
        border-right:1px solid rgba(255,255,255,.06);
    }
    section[data-testid="stSidebar"] > div { padding-top:1rem!important; }
    section[data-testid="stSidebar"] * { color:#eef1f6; }
    section[data-testid="stSidebar"] .stCaption,
    section[data-testid="stSidebar"] [data-testid="stCaptionContainer"] p { color:#8e96a6!important; }

    section[data-testid="stSidebar"] div[role="radiogroup"] {
        display:flex!important; flex-direction:column!important; gap:.25rem!important; width:100%;
    }
    section[data-testid="stSidebar"] div[role="radiogroup"] label {
        width:100%; border:0!important; border-radius:10px!important;
        background:transparent!important; padding:.52rem .62rem!important; min-height:42px;
    }
    section[data-testid="stSidebar"] div[role="radiogroup"] label:hover {
        background:rgba(255,255,255,.07)!important;
    }
    section[data-testid="stSidebar"] div[role="radiogroup"] label:has(input:checked) {
        background:rgba(255,255,255,.11)!important;
        box-shadow:inset 3px 0 0 var(--p-accent);
    }
    section[data-testid="stSidebar"] div[role="radiogroup"] label p {
        color:#e9edf5!important; font-size:.9rem!important; font-weight:650!important;
    }
    section[data-testid="stSidebar"] input[type="radio"] { display:none!important; }
    section[data-testid="stSidebar"] div[role="radiogroup"] label > div:first-child {
        display:none!important;
    }
    section[data-testid="stSidebar"] div[role="radiogroup"] label {
        grid-template-columns:1fr!important;
    }

    .ss-side-brand { display:flex; align-items:center; gap:.65rem; padding:.15rem .25rem .95rem; }
    .ss-side-logo {
        width:34px;height:34px;border-radius:10px;display:grid;place-items:center;
        background:linear-gradient(135deg,#ff4d5a,#ff725f);color:white;font-weight:900;
        box-shadow:0 8px 20px rgba(255,77,90,.22);
    }
    .ss-side-name { font-size:1.08rem;font-weight:800;letter-spacing:-.02em; }
    .ss-channel-chip {
        border:1px solid rgba(255,255,255,.08); background:rgba(255,255,255,.045);
        border-radius:12px; padding:.72rem .75rem; margin:.15rem 0 .85rem;
    }
    .ss-channel-chip .name {
        font-weight:700;font-size:.88rem;color:#fff;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;
    }
    .ss-channel-chip .status { color:#75dfaa;font-size:.75rem;margin-top:.28rem; }

    .ss-page-head {
        display:flex;align-items:flex-start;justify-content:space-between;gap:1rem;margin:.15rem 0 1.1rem;
    }
    .ss-page-kicker {
        color:var(--p-accent);font-weight:800;font-size:.72rem;letter-spacing:.08em;margin-bottom:.28rem;
    }
    .ss-page-title {
        color:var(--p-text);font-size:1.85rem;font-weight:820;letter-spacing:-.035em;line-height:1.12;
    }
    .ss-page-desc { color:var(--p-muted);font-size:.9rem;margin-top:.34rem;max-width:760px;line-height:1.5; }
    .ss-page-badge {
        flex:0 0 auto;border:1px solid var(--p-line);background:#fff;border-radius:999px;
        padding:.42rem .68rem;color:#4b5563;font-size:.76rem;font-weight:700;
    }

    .ss-home-strip {
        display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:.8rem;margin:.35rem 0 1.1rem;
    }
    .ss-mini-card {
        background:#fff;
        border:1px solid var(--p-line);
        border-radius:14px;
        padding:.9rem .95rem;
        min-height:108px;
        display:flex;
        flex-direction:column;
        justify-content:space-between;
    }
    .ss-mini-label { color:var(--p-muted);font-size:.73rem;font-weight:700;margin-bottom:.34rem; }
    .ss-mini-title {
        color:var(--p-text);font-size:.88rem;font-weight:720;line-height:1.35;
        display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;
    }
    .ss-mini-meta { margin-top:.48rem;color:#4b5563;font-size:.76rem; }

    [data-testid="stMetric"] {
        border-radius:14px!important;border:1px solid var(--p-line)!important;box-shadow:none!important;background:#fff!important;
    }
    [data-testid="stMetric"]:hover { transform:none!important;box-shadow:none!important; }
    [data-testid="stForm"] {
        background:#fff!important;border:1px solid var(--p-line)!important;border-radius:14px!important;
        box-shadow:none!important;padding:.8rem .85rem .85rem!important;
    }
    [data-baseweb="tab-list"] {
        background:transparent!important;border-bottom:1px solid var(--p-line);border-radius:0!important;padding:0!important;
    }
    [data-baseweb="tab"] {
        border-radius:0!important;background:transparent!important;box-shadow:none!important;border-bottom:2px solid transparent;
    }
    [aria-selected="true"][data-baseweb="tab"] { border-bottom-color:var(--p-accent)!important; }
    [data-testid="stDataFrame"] { box-shadow:none!important;border-radius:12px!important; }

    @media(max-width:768px){
        .ss-page-head{display:block}.ss-page-badge{display:inline-block;margin-top:.65rem}
        .ss-home-strip{grid-template-columns:1fr}.ss-page-title{font-size:1.5rem}
        .block-container{padding-top:.75rem!important}
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# V7.6 — Product Experience Upgrade
# 판단/신뢰/우선순위를 먼저 보이게 합니다.
# =========================================================
st.markdown(
    """
    <style>
    :root {
        --v76-text:#101828; --v76-muted:#667085; --v76-line:#eaecf0;
        --v76-surface:#ffffff; --v76-bg:#f7f8fa; --v76-brand:#ff4d5a;
        --v76-positive:#067647; --v76-positive-bg:#ecfdf3; --v76-positive-line:#abefc6;
        --v76-negative:#b42318; --v76-negative-bg:#fef3f2; --v76-negative-line:#fecdca;
        --v76-warning:#b54708; --v76-warning-bg:#fffaeb; --v76-warning-line:#fedf89;
        --v76-info:#175cd3; --v76-info-bg:#eff8ff; --v76-info-line:#b2ddff;
        --v76-neutral:#475467; --v76-neutral-bg:#f2f4f7; --v76-neutral-line:#d0d5dd;
    }
    .stApp { background:var(--v76-bg)!important; }
    .block-container { max-width:1240px!important; }
    .ss-page-head { margin-bottom:1.35rem!important; }
    .ss-page-title { font-size:1.72rem!important; }
    .ss-page-desc { max-width:680px!important; }

    .v76-kpi-grid { display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:.8rem;margin:.55rem 0 1.15rem; }
    .v76-kpi { background:#fff;border:1px solid var(--v76-line);border-radius:16px;padding:1rem 1.05rem;min-height:126px;display:flex;flex-direction:column;justify-content:space-between; }
    .v76-kpi-label { color:var(--v76-muted);font-size:.78rem;font-weight:700; }
    .v76-kpi-value { color:var(--v76-text);font-size:1.62rem;line-height:1.08;letter-spacing:-.035em;font-weight:820;margin:.38rem 0 .28rem; }
    .v76-kpi-foot { display:flex;align-items:center;gap:.42rem;color:var(--v76-muted);font-size:.75rem;line-height:1.3; }

    .v76-badge { display:inline-flex;align-items:center;border-radius:999px;padding:.25rem .48rem;font-size:.7rem;line-height:1;font-weight:800;white-space:nowrap;border:1px solid var(--v76-neutral-line);background:var(--v76-neutral-bg);color:var(--v76-neutral); }
    .v76-positive { color:var(--v76-positive)!important;background:var(--v76-positive-bg)!important;border-color:var(--v76-positive-line)!important; }
    .v76-negative { color:var(--v76-negative)!important;background:var(--v76-negative-bg)!important;border-color:var(--v76-negative-line)!important; }
    .v76-warning { color:var(--v76-warning)!important;background:var(--v76-warning-bg)!important;border-color:var(--v76-warning-line)!important; }
    .v76-info { color:var(--v76-info)!important;background:var(--v76-info-bg)!important;border-color:var(--v76-info-line)!important; }
    .v76-neutral { color:var(--v76-neutral)!important;background:var(--v76-neutral-bg)!important;border-color:var(--v76-neutral-line)!important; }

    .v76-brief { background:#111318;color:#f8fafc;border-radius:17px;padding:1rem 1.1rem;margin:.15rem 0 1.25rem;display:flex;gap:.8rem;align-items:flex-start; }
    .v76-brief-dot { width:9px;height:9px;border-radius:50%;background:var(--v76-brand);margin-top:.38rem;flex:0 0 auto;box-shadow:0 0 0 5px rgba(255,77,90,.12); }
    .v76-brief-kicker { color:#98a2b3;font-size:.69rem;font-weight:800;letter-spacing:.09em;margin-bottom:.3rem; }
    .v76-brief-text { font-size:.94rem;font-weight:650;line-height:1.52;letter-spacing:-.01em; }

    .v76-insight-grid { display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:.8rem;margin:.5rem 0 1.2rem; }
    .v76-insight { background:#fff;border:1px solid var(--v76-line);border-radius:15px;padding:.9rem .95rem;min-height:118px; }
    .v76-insight-top { display:flex;justify-content:space-between;align-items:center;gap:.6rem;margin-bottom:.55rem; }
    .v76-insight-label { color:var(--v76-muted);font-size:.73rem;font-weight:750; }
    .v76-insight-value { color:var(--v76-text);font-size:1.12rem;font-weight:820;letter-spacing:-.025em;line-height:1.25; }
    .v76-insight-desc { color:var(--v76-muted);font-size:.75rem;line-height:1.45;margin-top:.38rem; }

    .v76-video-grid { display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:.85rem;margin:.5rem 0 1.25rem; }
    .v76-video-card { background:#fff;border:1px solid var(--v76-line);border-radius:16px;overflow:hidden;min-width:0; }
    .v76-video-thumb { width:100%;aspect-ratio:16/9;object-fit:cover;display:block;background:#eef1f5; }
    .v76-video-body { padding:.8rem .85rem .85rem; }
    .v76-video-badges { display:flex;flex-wrap:wrap;gap:.35rem;margin-bottom:.55rem; }
    .v76-video-title { color:var(--v76-text);font-size:.88rem;font-weight:760;line-height:1.38;min-height:2.45em;display:-webkit-box;-webkit-box-orient:vertical;-webkit-line-clamp:2;overflow:hidden; }
    .v76-video-meta { display:flex;justify-content:space-between;gap:.45rem;margin-top:.62rem;color:var(--v76-muted);font-size:.73rem; }
    .v76-video-gain { font-weight:800;color:var(--v76-text); }

    .v76-empty { border:1px dashed #cfd4dc;background:#fbfcfe;border-radius:15px;padding:1rem;margin:.45rem 0 1rem; }
    .v76-empty-title { color:var(--v76-text);font-size:.86rem;font-weight:780; }
    .v76-empty-desc { color:var(--v76-muted);font-size:.76rem;line-height:1.5;margin-top:.32rem; }

    .v76-rank-grid { display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:.75rem;margin:.45rem 0 1rem; }
    .v76-rank-card { background:#fff;border:1px solid var(--v76-line);border-radius:14px;padding:.85rem .9rem; }
    .v76-rank-no { color:var(--v76-brand);font-size:.72rem;font-weight:850; }
    .v76-rank-title { font-weight:750;font-size:.82rem;line-height:1.38;margin:.28rem 0 .5rem;min-height:2.25em; }
    .v76-rank-value { font-size:1.12rem;font-weight:820;letter-spacing:-.025em; }
    .v76-rank-meta { color:var(--v76-muted);font-size:.72rem;margin-top:.28rem; }

    [data-testid="stFormSubmitButton"] button { background:#fff!important;color:#344054!important;border:1px solid #d0d5dd!important; }
    [data-testid="stFormSubmitButton"] button:hover { background:#f9fafb!important;color:#101828!important;border-color:#98a2b3!important; }

    @media(max-width:900px){
        .v76-kpi-grid,.v76-insight-grid,.v76-video-grid,.v76-rank-grid { grid-template-columns:1fr!important; }
        .v76-kpi { min-height:104px; }
        .v76-video-title { min-height:auto; }
    }
    </style>
    """,
    unsafe_allow_html=True,
)



# =========================================================
# V7.8 — UI/UX Stabilization
# 긴 분석 화면의 밀도와 시각적 우선순위를 정리합니다.
# =========================================================
st.markdown(
    """
    <style>
    /* 일반 콘텐츠 폭과 세로 밀도 */
    .block-container {
        max-width: 1180px !important;
        padding-top: 1.05rem !important;
    }

    h2 {
        margin-top: 1.15rem !important;
        margin-bottom: .45rem !important;
    }
    h3 {
        margin-top: 1rem !important;
        margin-bottom: .4rem !important;
    }

    /* 설명 문장은 조금 더 읽히게 */
    [data-testid="stCaptionContainer"] p {
        color: #7a8495 !important;
        font-size: .82rem !important;
        line-height: 1.45 !important;
    }

    /* 폼을 덜 거대하게 */
    [data-testid="stForm"] {
        padding: .68rem .75rem .72rem !important;
        margin-bottom: .45rem !important;
        border-radius: 13px !important;
    }

    /* 입력 높이 정리 */
    [data-baseweb="select"] > div,
    [data-baseweb="input"] > div {
        min-height: 40px !important;
    }

    /* 메트릭 카드 높이 과다 방지 */
    [data-testid="stMetric"] {
        min-height: 96px !important;
        padding: .8rem .9rem !important;
    }

    /* expander를 '상세 정보' 느낌으로 */
    [data-testid="stExpander"] {
        margin-bottom: .55rem !important;
        border-color: #e4e7ec !important;
        box-shadow: none !important;
    }

    /* 구분선 간격 축소 */
    hr {
        margin: 1.15rem 0 !important;
    }

    /* 사이드바 메뉴 원형 라디오 완전히 숨김 */
    section[data-testid="stSidebar"] div[role="radiogroup"] label > div:first-child {
        display: none !important;
    }
    section[data-testid="stSidebar"] div[role="radiogroup"] label {
        display: block !important;
        padding: .5rem .68rem !important;
    }

    /* 현재 작업 흐름 안내 */
    .v78-flow {
        display:flex;
        flex-wrap:wrap;
        gap:.4rem;
        margin:.15rem 0 .9rem;
    }
    .v78-flow span {
        display:inline-flex;
        align-items:center;
        border:1px solid #e4e7ec;
        background:#fff;
        color:#667085;
        border-radius:999px;
        padding:.28rem .52rem;
        font-size:.72rem;
        font-weight:700;
    }
    .v78-flow span strong {
        color:#ff4d5a;
        margin-right:.2rem;
    }

    @media(max-width:768px){
        .block-container {
            padding-left:.7rem !important;
            padding-right:.7rem !important;
        }
        [data-testid="stMetric"] {
            min-height:88px !important;
        }
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# 2. OAuth 설정
# =========================================================

if not CLIENT_ID or not CLIENT_SECRET:

    st.error(
        "Google OAuth 설정을 찾을 수 없습니다."
    )

    st.stop()


client_config = {
    "web": {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "auth_uri":
            "https://accounts.google.com/o/oauth2/auth",
        "token_uri":
            "https://oauth2.googleapis.com/token",
        "redirect_uris": [
            REDIRECT_URI
        ],
    }
}


def create_flow():

    return Flow.from_client_config(
        client_config,
        scopes=SCOPES,
        redirect_uri=REDIRECT_URI,
        autogenerate_code_verifier=False,
    )


# =========================================================
# 3. OAuth Callback
# =========================================================

code = st.query_params.get("code")
oauth_error = st.query_params.get("error")


if oauth_error:

    st.error(
        f"Google 로그인이 실패했습니다: "
        f"{oauth_error}"
    )

    if st.button("🔄 다시 로그인"):

        st.query_params.clear()

        st.rerun()

    st.stop()


if (
    code
    and "credentials"
    not in st.session_state
):

    try:

        flow = create_flow()

        flow.fetch_token(
            code=code
        )

        credentials = flow.credentials

        st.session_state[
            "credentials"
        ] = {

            "token":
                credentials.token,

            "refresh_token":
                credentials.refresh_token,

            "token_uri":
                credentials.token_uri,

            "client_id":
                credentials.client_id,

            "client_secret":
                credentials.client_secret,

            "scopes":
                list(
                    credentials.scopes
                    or SCOPES
                ),
        }

        st.query_params.clear()

        st.rerun()

    except Exception as e:

        st.error(
            "Google 로그인 처리 중 오류가 발생했습니다."
        )
        with st.expander("기술 오류 상세보기"):
            st.code(str(e))
        st.stop()


# =========================================================
# 4. 로그인 전
# =========================================================

if (
    "credentials"
    not in st.session_state
):

    flow = create_flow()

    authorization_url, _ = (
        flow.authorization_url(
            access_type="offline",
            prompt="consent",
            include_granted_scopes="true",
        )
    )

    st.subheader(
        "🔗 YouTube 채널 연결"
    )

    st.write(
        "Google 계정을 연결하면 본인 YouTube 채널의 읽기 전용 데이터를 불러와 "
        "채널 및 Shorts 성과 분석을 시작합니다."
    )

    st.caption(
        "연결 전에도 위에서 Shorts Scope의 목적과 개인정보처리방침을 확인할 수 있습니다."
    )

    st.link_button(
        "🔴 Google로 YouTube 연결",
        authorization_url,
        type="primary",
    )

    st.stop()


# =========================================================
# 5. Credentials 복원
# =========================================================

saved = st.session_state[
    "credentials"
]

credentials = Credentials(

    token=
        saved["token"],

    refresh_token=
        saved.get(
            "refresh_token"
        ),

    token_uri=
        saved["token_uri"],

    client_id=
        saved["client_id"],

    client_secret=
        saved["client_secret"],

    scopes=
        saved.get(
            "scopes",
            SCOPES
        ),
)


# =========================================================
# 6. API 연결
# =========================================================

try:

    youtube = create_youtube_api(
        credentials
    )

    yt_analytics = (
        create_analytics_api(
            credentials
        )
    )

except Exception as e:

    st.error("YouTube API 연결 실패")
    with st.expander("기술 오류 상세보기"):
        st.code(str(e))
    st.stop()


# =========================================================
# 7. 채널 + 영상 정보
# =========================================================

try:

    channel_info = (
        get_channel_info(
            youtube
        )
    )

    if not channel_info:

        st.error(
            "YouTube 채널을 찾지 못했습니다."
        )

        st.stop()

    videos = (
        get_all_channel_videos(
            youtube,
            channel_info
        )
    )

    split = (
        split_videos_by_status(
            videos
        )
    )

    public_videos = (
        split["public"]
    )

    scheduled_videos = (
        split["scheduled"]
    )

    private_videos = (
        split["private"]
    )

    unlisted_videos = (
        split["unlisted"]
    )

except Exception as e:

    st.error("YouTube 데이터를 가져오지 못했습니다.")
    with st.expander("기술 오류 상세보기"):
        st.code(str(e))
    st.stop()


# =========================================================
# 8. 영상별 Analytics
# =========================================================

public_video_ids = [

    video["video_id"]

    for video in public_videos
]


try:

    video_analytics = (
        get_video_analytics(
            yt_analytics,
            public_video_ids
        )
    )

except Exception as e:

    video_analytics = {}

    st.warning(
        "영상별 Analytics 일부를 "
        "불러오지 못했습니다."
    )

    with st.expander(
        "기술 오류 상세보기"
    ):

        st.code(
            str(e)
        )


# =========================================================
# 9. 영상 데이터 + Analytics 합치기
# =========================================================

for video in videos:

    data = video_analytics.get(
        video["video_id"],
        {}
    )

    video["watch_minutes"] = float(
        data.get(
            "watch_minutes",
            0
        )
    )

    video["avg_duration"] = float(
        data.get(
            "average_view_duration",
            0
        )
    )

    video["avg_percentage"] = float(
        data.get(
            "average_view_percentage",
            0
        )
    )

    video["shares"] = int(
        data.get(
            "shares",
            0
        )
    )

    video["subs_gained"] = int(
        data.get(
            "subscribers_gained",
            0
        )
    )

    video["subs_lost"] = int(
        data.get(
            "subscribers_lost",
            0
        )
    )

    video["net_subs"] = int(
        data.get(
            "net_subscribers",
            0
        )
    )

    # 비율 지표
    views_for_rate = max(video["views"], 1)
    video["like_rate"] = (video["likes"] / views_for_rate) * 100
    video["comment_rate"] = (video["comments"] / views_for_rate) * 100
    video["sub_conversion_rate"] = (video["net_subs"] / views_for_rate) * 100


# =========================================================
# 9-1. 자동 성과 점수 V1
# =========================================================

# 조회수 100회 이상 + Analytics가 실제로 잡힌 공개 영상만 평가
eligible_videos = [
    video
    for video in public_videos
    if video.get("views", 0) >= 100
    and video.get("video_id") in video_analytics
]


def percentile_scores(items, key):
    if not items:
        return {}

    values = sorted(
        (item.get(key, 0), item["video_id"])
        for item in items
    )
    total = len(values)

    if total == 1:
        return {values[0][1]: 100.0}

    result = {}
    for index, (_, video_id) in enumerate(values):
        result[video_id] = (index / (total - 1)) * 100

    return result


views_pct = percentile_scores(eligible_videos, "views")
retention_pct = percentile_scores(eligible_videos, "avg_percentage")
likes_pct = percentile_scores(eligible_videos, "like_rate")
subs_pct = percentile_scores(eligible_videos, "sub_conversion_rate")


for video in public_videos:

    if video not in eligible_videos:
        video["performance_score"] = None
        video["performance_grade"] = "⏳ 데이터 부족"
        video["diagnosis_strengths"] = []
        video["diagnosis_weaknesses"] = []
        continue

    score = (
        views_pct.get(video["video_id"], 0) * 0.30
        + retention_pct.get(video["video_id"], 0) * 0.35
        + likes_pct.get(video["video_id"], 0) * 0.15
        + subs_pct.get(video["video_id"], 0) * 0.20
    )

    video["performance_score"] = round(score)

    if score >= 80:
        video["performance_grade"] = "🔥 매우 좋음"
    elif score >= 65:
        video["performance_grade"] = "🟢 좋음"
    elif score >= 45:
        video["performance_grade"] = "🟡 보통"
    else:
        video["performance_grade"] = "🔴 개선 필요"

    strengths = []
    weaknesses = []

    if retention_pct.get(video["video_id"], 0) >= 70:
        strengths.append("시청 유지")
    elif retention_pct.get(video["video_id"], 0) <= 30:
        weaknesses.append("시청 유지")

    if subs_pct.get(video["video_id"], 0) >= 70:
        strengths.append("구독 전환")
    elif subs_pct.get(video["video_id"], 0) <= 30:
        weaknesses.append("구독 전환")

    if likes_pct.get(video["video_id"], 0) >= 70:
        strengths.append("좋아요 반응")
    elif likes_pct.get(video["video_id"], 0) <= 30:
        weaknesses.append("좋아요 반응")

    if views_pct.get(video["video_id"], 0) >= 70:
        strengths.append("조회수")
    elif views_pct.get(video["video_id"], 0) <= 30:
        weaknesses.append("조회수")

    video["diagnosis_strengths"] = strengths
    video["diagnosis_weaknesses"] = weaknesses


# =========================================================
# 10. 채널 상단
# =========================================================

_page_help = {
    "🏠 홈": "채널 핵심 상태와 최근 영상",
    "📈 성장 분석": "성장상태 · 동일 나이 · 성장곡선",
    "📊 채널 패턴": "최근 영상 · 요일 · 시간대 패턴",
    "🧪 운영": "목표 · 태그 · 메모 · 실험",
    "📋 리포트": "주간 변화 · 이상 감지 · 최고기록",
    "🔎 영상 찾기": "검색 · 필터 · 전체 영상",
}

_page_meta = {
    "🏠 홈": ("OVERVIEW", "채널 대시보드", "지금 필요한 핵심 정보만 빠르게 확인합니다."),
    "📈 성장 분석": ("GROWTH", "성장 분석", "스냅샷과 Analytics로 영상의 실제 성장 흐름을 비교합니다."),
    "📊 채널 패턴": ("PATTERNS", "채널 패턴", "최근 업로드의 요일·시간대·성과 패턴을 실제 데이터로 비교합니다."),
    "🧪 운영": ("WORKSPACE", "운영 기록", "목표, 영상 메모, 태그와 실험을 한곳에서 관리합니다."),
    "📋 리포트": ("REPORTS", "리포트", "최근 변화, 이상 신호와 개인 최고기록을 정리합니다."),
    "🔎 영상 찾기": ("LIBRARY", "영상 찾기", "전체 영상을 검색하고 조건별로 빠르게 좁혀봅니다."),
}

with st.sidebar:
    st.markdown(
        """
        <div class="ss-side-brand">
            <div class="ss-side-logo">▶</div>
            <div class="ss-side-name">Shorts Scope</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.markdown(
        f"""
        <div class="ss-channel-chip">
            <div class="name">{channel_info['channel_name']}</div>
            <div class="status">● YouTube 연결됨</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.caption("ANALYTICS")
    page = st.radio(
        "화면 선택",
        ["🏠 홈", "📈 성장 분석", "📊 채널 패턴", "🧪 운영", "📋 리포트", "🔎 영상 찾기"],
        format_func=lambda x: x.split(" ", 1)[1] if " " in x else x,
        label_visibility="collapsed",
        key="main_page_v64",
    )
    st.markdown("---")
    st.caption(_page_help[page])
    st.caption("YouTube 읽기 전용 · 콘텐츠 수정/삭제 안 함")
    st.link_button(
        "개인정보처리방침",
        "https://goldmoon123.github.io/alssulcut-dashboard/privacy.html",
        use_container_width=True,
    )

_kicker, _page_title, _page_desc = _page_meta[page]
st.markdown(
    f"""
    <div class="ss-page-head">
        <div>
            <div class="ss-page-kicker">{_kicker}</div>
            <div class="ss-page-title">{_page_title}</div>
            <div class="ss-page-desc">{_page_desc}</div>
        </div>
        <div class="ss-page-badge">읽기 전용 · 실제 데이터</div>
    </div>
    """,
    unsafe_allow_html=True,
)


# 각 메뉴가 독립적으로 실행되어도 필요한 공통 값
today = datetime.now(KST).date()


def add_excel_table(ws, table_name):
    """Excel 실제 표(Table) + 필터/정렬."""
    if ws.max_row < 2 or ws.max_column < 1:
        return
    safe_name = re.sub(r"[^A-Za-z0-9_]", "_", table_name)
    if not safe_name or safe_name[0].isdigit():
        safe_name = "T_" + safe_name
    table = Table(displayName=safe_name, ref=ws.dimensions)
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    ws.add_table(table)



if page == "🏠 홈":
    # =========================================================
    # V7.5 — Home Action Center
    # "숫자 나열"보다 "지금 무엇을 봐야 하는지"를 먼저 보여줍니다.
    # 원인 단정 없이 실제 데이터/관찰 사실만 사용합니다.
    # =========================================================

    # -----------------------------
    # 기본 집계
    # -----------------------------
    _home_total_views = sum(int(v.get("views", 0) or 0) for v in public_videos)
    _home_avg_views = (
        _home_total_views / len(public_videos)
        if public_videos else 0
    )
    _home_view_values = sorted(
        int(v.get("views", 0) or 0)
        for v in public_videos
    )
    _home_median_views = (
        float(pd.Series(_home_view_values).median())
        if _home_view_values else 0
    )

    # 최근 7일 vs 이전 7일
    _home_end = today - timedelta(days=1)
    _home_start = _home_end - timedelta(days=6)
    _home_prev_end = _home_start - timedelta(days=1)
    _home_prev_start = _home_prev_end - timedelta(days=6)

    _home_week_now = None
    _home_week_prev = None
    try:
        _home_week_now = get_period_summary(
            yt_analytics,
            _home_start,
            _home_end,
        )
        _home_week_prev = get_period_summary(
            yt_analytics,
            _home_prev_start,
            _home_prev_end,
        )
    except Exception:
        pass

    def _home_pct_change(cur, prev):
        if prev in (None, 0):
            return None
        return ((cur - prev) / abs(prev)) * 100

    _home_views_change = None
    _home_watch_change = None
    _home_sub_delta = None

    if _home_week_now and _home_week_prev:
        _home_views_change = _home_pct_change(
            _home_week_now.get("views", 0),
            _home_week_prev.get("views", 0),
        )
        _home_watch_change = _home_pct_change(
            _home_week_now.get("watch_minutes", 0),
            _home_week_prev.get("watch_minutes", 0),
        )
        _home_sub_delta = (
            int(_home_week_now.get("net_subscribers", 0) or 0)
            - int(_home_week_prev.get("net_subscribers", 0) or 0)
        )

    # 최근 10개 중 중앙값 이하
    _home_recent = []
    for _hv in public_videos:
        _raw = _hv.get("published_raw")
        if not _raw:
            continue
        try:
            _hdt = datetime.fromisoformat(
                _raw.replace("Z", "+00:00")
            ).astimezone(KST)
            _home_recent.append((_hdt, _hv))
        except Exception:
            pass

    _home_recent.sort(key=lambda x: x[0], reverse=True)
    _home_recent10 = _home_recent[:10]
    _home_below_median = sum(
        1
        for _, _v in _home_recent10
        if int(_v.get("views", 0) or 0) < _home_median_views
    )

    # 최근 스냅샷 기반 "지금 볼 영상"
    _home_snapshot_by_video = {}
    _home_snapshot_fetch = {"ok": False, "rows": []}
    try:
        _home_channel_id = _connected_youtube_channel_id(youtube)
        _home_since = (
            datetime.now(timezone.utc) - timedelta(days=4)
        ).replace(minute=0, second=0, microsecond=0).isoformat()
        _home_snapshot_fetch = _fetch_channel_snapshots(
            _home_channel_id,
            _home_since,
        )
        if _home_snapshot_fetch.get("ok"):
            _home_snapshot_by_video = _group_snapshot_rows(
                _home_snapshot_fetch.get("rows", []),
                [v.get("video_id") for v in public_videos],
            )
    except Exception:
        pass

    _home_watch_candidates = []
    _home_now_utc = datetime.now(timezone.utc)

    for _hdt, _v in _home_recent[:20]:
        _rows = _home_snapshot_by_video.get(_v.get("video_id"), [])
        if _rows:
            _state = _snapshot_growth_state(
                _v,
                _rows,
                _home_now_utc,
            )
            _event = _snapshot_special_event(
                _v,
                _state,
                _rows,
                _home_now_utc,
            )
        else:
            _state = {
                "state": "⏳ 데이터 축적 중",
                "confidence": "낮음",
                "recent_gain": None,
                "previous_gain": None,
            }
            _event = {"label": None}

        _priority = 0
        if _event.get("label") == "🚀 급상승":
            _priority = 5
        elif _event.get("label") == "🔥 재상승":
            _priority = 4
        elif _state.get("state") == "↗ 상승":
            _priority = 3
        elif _state.get("state") == "↘ 하락":
            _priority = 2
        elif _state.get("state") == "⏳ 데이터 축적 중":
            _priority = 1

        _home_watch_candidates.append({
            "video": _v,
            "published": _hdt,
            "state": _state,
            "event": _event,
            "priority": _priority,
        })

    _home_watch_candidates.sort(
        key=lambda x: (
            x["priority"],
            x["published"],
        ),
        reverse=True,
    )
    _home_watch = _home_watch_candidates[:3]

    # 채널 패턴: 최근 20개 업로드 시간대 평균
    _home_band_labels = {
        "새벽 00~05시": [],
        "오전 06~11시": [],
        "오후 12~17시": [],
        "저녁 18~23시": [],
    }

    for _hdt, _v in _home_recent[:20]:
        _hour = _hdt.hour
        if 0 <= _hour <= 5:
            _band = "새벽 00~05시"
        elif 6 <= _hour <= 11:
            _band = "오전 06~11시"
        elif 12 <= _hour <= 17:
            _band = "오후 12~17시"
        else:
            _band = "저녁 18~23시"
        _home_band_labels[_band].append(
            int(_v.get("views", 0) or 0)
        )

    _home_band_summary = []
    for _band, _vals in _home_band_labels.items():
        if not _vals:
            continue
        _home_band_summary.append({
            "band": _band,
            "count": len(_vals),
            "avg": sum(_vals) / len(_vals),
        })

    _home_band_summary.sort(
        key=lambda x: x["avg"],
        reverse=True,
    )

    # -----------------------------
    # 1) 오늘의 상태 + 팩트 브리핑
    # -----------------------------
    _special_count = sum(
        1
        for x in _home_watch_candidates
        if x["event"].get("label") in ("🚀 급상승", "🔥 재상승")
    )

    _brief_parts = []
    if _home_views_change is not None:
        if _home_views_change > 0:
            _brief_parts.append(f"최근 7일 조회수는 이전 7일보다 {_home_views_change:.1f}% 높습니다.")
        elif _home_views_change < 0:
            _brief_parts.append(f"최근 7일 조회수는 이전 7일보다 {abs(_home_views_change):.1f}% 낮습니다.")
        else:
            _brief_parts.append("최근 7일 조회수는 이전 7일과 비슷한 수준입니다.")
    if _home_recent10:
        _brief_parts.append(
            f"최근 {len(_home_recent10)}개 영상 중 {_home_below_median}개가 "
            f"채널 중앙 조회수 {_home_median_views:,.0f}회 아래입니다."
        )
    if _special_count:
        _brief_parts.append(f"현재 급상승·재상승으로 감지된 영상은 {_special_count}개입니다.")
    else:
        _brief_parts.append("현재 급상승·재상승으로 감지된 영상은 없습니다.")

    _brief_text = " ".join(_brief_parts) or "현재 비교할 데이터가 충분히 쌓이는 중입니다."

    st.markdown(
        f"""
        <div class="v76-brief">
            <div class="v76-brief-dot"></div>
            <div>
                <div class="v76-brief-kicker">TODAY BRIEF</div>
                <div class="v76-brief-text">{_brief_text}</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown("### 오늘의 채널 상태")

    if _home_views_change is None:
        _views_badge_cls, _views_badge = "v76-info", "집계 중"
        _views_value = f"{int(_home_week_now.get('views', 0) or 0):,}회" if _home_week_now else "데이터 축적 중"
        _views_foot = "이전 7일 비교값 대기"
    elif _home_views_change > 0:
        _views_badge_cls, _views_badge = "v76-positive", f"↑ {_home_views_change:.1f}%"
        _views_value = f"{int(_home_week_now.get('views', 0) or 0):,}회"
        _views_foot = "이전 7일 대비"
    elif _home_views_change < 0:
        _views_badge_cls, _views_badge = "v76-negative", f"↓ {abs(_home_views_change):.1f}%"
        _views_value = f"{int(_home_week_now.get('views', 0) or 0):,}회"
        _views_foot = "이전 7일 대비"
    else:
        _views_badge_cls, _views_badge = "v76-neutral", "→ 0.0%"
        _views_value = f"{int(_home_week_now.get('views', 0) or 0):,}회"
        _views_foot = "이전 7일 대비"

    _below_ratio = (_home_below_median / len(_home_recent10)) if _home_recent10 else 0
    if not _home_recent10:
        _median_badge_cls, _median_badge = "v76-info", "축적 중"
    elif _below_ratio >= 0.7:
        _median_badge_cls, _median_badge = "v76-warning", "확인 필요"
    elif _below_ratio <= 0.3:
        _median_badge_cls, _median_badge = "v76-positive", "양호"
    else:
        _median_badge_cls, _median_badge = "v76-neutral", "중립"

    _special_badge_cls = "v76-positive" if _special_count > 0 else "v76-neutral"
    _special_badge = "감지됨" if _special_count > 0 else "없음"

    st.markdown(
        f"""
        <div class="v76-kpi-grid">
            <div class="v76-kpi">
                <div><div class="v76-kpi-label">최근 7일 조회수</div><div class="v76-kpi-value">{_views_value}</div></div>
                <div class="v76-kpi-foot"><span class="v76-badge {_views_badge_cls}">{_views_badge}</span><span>{_views_foot}</span></div>
            </div>
            <div class="v76-kpi">
                <div><div class="v76-kpi-label">최근 10개 중 중앙값 이하</div><div class="v76-kpi-value">{_home_below_median}개</div></div>
                <div class="v76-kpi-foot"><span class="v76-badge {_median_badge_cls}">{_median_badge}</span><span>채널 중앙값 {_home_median_views:,.0f}회 기준</span></div>
            </div>
            <div class="v76-kpi">
                <div><div class="v76-kpi-label">현재 특이 성장</div><div class="v76-kpi-value">{_special_count}개</div></div>
                <div class="v76-kpi-foot"><span class="v76-badge {_special_badge_cls}">{_special_badge}</span><span>급상승·재상승 감지 기준</span></div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # -----------------------------
    # 2) 팩트 기반 인사이트
    # -----------------------------
    st.markdown("### 지금 확인할 변화")

    _insights = []

    if _home_views_change is not None and _home_week_now:
        if abs(_home_views_change) >= 10:
            _insights.append({
                "title": "최근 7일 조회수 변화",
                "value": f"{_home_views_change:+.1f}%",
                "desc": (
                    f"{_home_start} ~ {_home_end} 기준 · "
                    "이전 7일과 비교한 실제 Analytics 변화"
                ),
            })

    if _home_recent10:
        _insights.append({
            "title": "최근 영상 분포",
            "value": f"{_home_below_median}/{len(_home_recent10)}",
            "desc": (
                f"최근 {len(_home_recent10)}개 중 {_home_below_median}개가 "
                f"채널 전체 중앙값 {_home_median_views:,.0f}회 미만"
            ),
        })

    if _home_band_summary:
        _best_band = _home_band_summary[0]
        _band_conf = (
            "낮음" if _best_band["count"] < 3
            else "보통" if _best_band["count"] < 6
            else "높음"
        )
        _insights.append({
            "title": "현재 시간대 상위 그룹",
            "value": _best_band["band"],
            "desc": (
                f"최근 표본 평균 {_best_band['avg']:,.0f}회 · "
                f"표본 {_best_band['count']}개 · 신뢰도 {_band_conf}"
            ),
        })

    if _insights:
        _insight_cards = []
        for _item in _insights[:3]:
            _title = _item["title"]
            _value = _item["value"]
            _desc = _item["desc"]
            if _title == "최근 7일 조회수 변화":
                if _home_views_change is not None and _home_views_change < 0:
                    _badge_cls, _badge_text = "v76-negative", "하락"
                elif _home_views_change is not None and _home_views_change > 0:
                    _badge_cls, _badge_text = "v76-positive", "상승"
                else:
                    _badge_cls, _badge_text = "v76-neutral", "유지"
            elif _title == "최근 영상 분포":
                if _home_recent10 and (_home_below_median / len(_home_recent10)) >= 0.7:
                    _badge_cls, _badge_text = "v76-warning", "확인 필요"
                else:
                    _badge_cls, _badge_text = "v76-neutral", "분포"
            else:
                _badge_cls = "v76-warning" if "_best_band" in locals() and _best_band["count"] < 3 else "v76-info"
                _badge_text = f"표본 {_best_band['count']}개" if "_best_band" in locals() else "실제 데이터"

            _insight_cards.append(
                (
                    f'<div class="v76-insight">'
                    f'<div class="v76-insight-top">'
                    f'<div class="v76-insight-label">{_title}</div>'
                    f'<span class="v76-badge {_badge_cls}">{_badge_text}</span>'
                    f'</div>'
                    f'<div class="v76-insight-value">{_value}</div>'
                    f'<div class="v76-insight-desc">{_desc}</div>'
                    f'</div>'
                )
            )
        st.markdown('<div class="v76-insight-grid">' + "".join(_insight_cards) + '</div>', unsafe_allow_html=True)

    # -----------------------------
    # 3) 지금 볼 영상
    # -----------------------------
    st.markdown("### 지금 볼 영상")
    st.caption("현재 성장 흐름에서 먼저 확인할 가치가 높은 최근 영상입니다.")

    if _home_watch:
        _video_cards = []
        for _item in _home_watch:
            _v = _item["video"]
            _state = _item["state"]
            _event = _item["event"]
            _status_label = _event.get("label") or _state.get("state") or "⏳ 데이터 축적 중"
            if _event.get("label") in ("🚀 급상승", "🔥 재상승") or _state.get("state") == "↗ 상승":
                _status_cls = "v76-positive"
            elif _state.get("state") == "↘ 하락":
                _status_cls = "v76-negative"
            elif "축적" in str(_status_label) or "대기" in str(_status_label):
                _status_cls = "v76-info"
            else:
                _status_cls = "v76-neutral"

            _confidence = _state.get("confidence", "낮음")
            _confidence_cls = "v76-positive" if _confidence == "높음" else "v76-info" if _confidence == "보통" else "v76-warning"
            _gain = _state.get("recent_gain")
            _gain_text = f"최근 +{int(_gain):,}회" if _gain is not None and int(_gain) >= 0 else "최근 증가 계산 중"
            _title = str(_v.get("title", "제목 없음")).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            _thumb = str(_v.get("thumbnail") or "")
            _thumb_html = f'<img class="v76-video-thumb" src="{_thumb}" alt="">' if _thumb else '<div class="v76-video-thumb"></div>'

            _video_cards.append(
                (
                    f'<div class="v76-video-card">'
                    f'{_thumb_html}'
                    f'<div class="v76-video-body">'
                    f'<div class="v76-video-badges">'
                    f'<span class="v76-badge {_status_cls}">{_status_label}</span>'
                    f'<span class="v76-badge {_confidence_cls}">신뢰도 {_confidence}</span>'
                    f'</div>'
                    f'<div class="v76-video-title">{_title}</div>'
                    f'<div class="v76-video-meta">'
                    f'<span>{_item["published"].strftime("%m.%d %H:%M")}</span>'
                    f'<span>{int(_v.get("views", 0) or 0):,}회</span>'
                    f'</div>'
                    f'<div class="v76-video-meta">'
                    f'<span>현재 비교구간</span>'
                    f'<span class="v76-video-gain">{_gain_text}</span>'
                    f'</div>'
                    f'</div>'
                    f'</div>'
                )
            )
        st.markdown('<div class="v76-video-grid">' + "".join(_video_cards) + '</div>', unsafe_allow_html=True)
    else:
        st.markdown(
            """
            <div class="v76-empty">
                <div class="v76-empty-title">최근 영상 데이터를 모으는 중입니다.</div>
                <div class="v76-empty-desc">스냅샷이 쌓이면 상승·하락·급상승 여부를 기준으로 먼저 볼 영상을 자동 정렬합니다.</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    # -----------------------------
    # 4) 데이터 신뢰 상태
    # -----------------------------
    st.markdown("### 데이터 신뢰 상태")

    _latest_snap_text = "확인 중"
    _snapshot_badge_cls, _snapshot_badge = "v76-warning", "확인 필요"
    _snapshot_desc = "스냅샷 상태를 확인하지 못했습니다."

    if _home_snapshot_fetch.get("ok"):
        _snap_rows = _home_snapshot_fetch.get("rows", [])
        _valid_snap_times = [
            _parse_utc(r.get("captured_at"))
            for r in _snap_rows
            if r.get("captured_at") and _parse_utc(r.get("captured_at"))
        ]
        if _valid_snap_times:
            _latest_snap = max(_valid_snap_times)
            _latest_snap_text = _latest_snap.astimezone(KST).strftime("%m.%d %H:%M")
            _snap_age_min = (datetime.now(timezone.utc) - _latest_snap).total_seconds() / 60
            if _snap_age_min <= 45:
                _snapshot_badge_cls, _snapshot_badge = "v76-positive", "정상"
            elif _snap_age_min <= 90:
                _snapshot_badge_cls, _snapshot_badge = "v76-warning", "지연"
            else:
                _snapshot_badge_cls, _snapshot_badge = "v76-negative", "확인 필요"
            _snapshot_desc = f"마지막 자체 수집 {_latest_snap_text} · 약 {_snap_age_min:.0f}분 전"
        else:
            _snapshot_badge_cls, _snapshot_badge = "v76-info", "축적 중"
            _snapshot_desc = "최근 스냅샷 행이 아직 없습니다."

    st.markdown(
        f"""
        <div class="v76-insight-grid">
            <div class="v76-insight">
                <div class="v76-insight-top"><div class="v76-insight-label">YouTube Analytics</div><span class="v76-badge v76-positive">실측</span></div>
                <div class="v76-insight-value">연결 정상</div>
                <div class="v76-insight-desc">읽기 전용 · 확정 데이터 우선</div>
            </div>
            <div class="v76-insight">
                <div class="v76-insight-top"><div class="v76-insight-label">Shorts Scope 스냅샷</div><span class="v76-badge {_snapshot_badge_cls}">{_snapshot_badge}</span></div>
                <div class="v76-insight-value">{_latest_snap_text}</div>
                <div class="v76-insight-desc">{_snapshot_desc}</div>
            </div>
            <div class="v76-insight">
                <div class="v76-insight-top"><div class="v76-insight-label">최근 Analytics</div><span class="v76-badge v76-info">집계 지연 가능</span></div>
                <div class="v76-insight-value">확정값 우선</div>
                <div class="v76-insight-desc">미확정 최근 날짜는 스냅샷 임시값으로 구분해 표시합니다.</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.divider()

    # -----------------------------
    # 5) 기존 핵심 성과 — 상세 영역으로 한 단계 아래
    # -----------------------------
    st.markdown("### 채널 기본 정보")

    c1, c2, c3 = st.columns(3)

    c1.metric(
        "구독자",
        f"{channel_info['subscribers']:,}명"
    )

    c2.metric(
        "채널 총 조회수",
        f"{channel_info['total_views']:,}회"
    )

    c3.metric(
        "공개 영상",
        f"{len(public_videos):,}개"
    )

    if public_videos:
        total_views = sum(
            video["views"]
            for video in public_videos
        )

        average_views = (
            total_views
            / len(public_videos)
        )

        best_video = max(
            public_videos,
            key=lambda x: x["views"]
        )

        total_watch_minutes = sum(
            video["watch_minutes"]
            for video in public_videos
        )

        total_net_subscribers = sum(
            video["net_subs"]
            for video in public_videos
        )

        c1, c2, c3, c4 = st.columns(4)

        c1.metric(
            "평균 조회수",
            f"{average_views:,.0f}회"
        )

        c2.metric(
            "최고 조회수",
            f"{best_video['views']:,}회"
        )

        c3.metric(
            "총 시청시간",
            format_watch_time(
                total_watch_minutes
            )
        )

        c4.metric(
            "공개 영상 순구독자",
            f"{total_net_subscribers:+,}명"
        )
    else:
        st.info("현재 공개 영상이 없습니다.")

    with st.expander(
        f"영상 상태 · 전체 {len(videos):,}개",
        expanded=False,
    ):
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("공개", f"{len(public_videos):,}개")
        c2.metric("예약", f"{len(scheduled_videos):,}개")
        c3.metric("비공개", f"{len(private_videos):,}개")
        c4.metric("일부공개", f"{len(unlisted_videos):,}개")
        st.caption("예약·비공개·일부공개 영상은 성과 평균과 순위에서 제외됩니다.")

    st.markdown("### 상세 분석")
    st.caption("기간 성과·일별 차트·월간 달력·과거 일별 분석은 필요할 때만 펼쳐서 확인합니다.")
    show_home_details = st.toggle(
        "기간·달력 상세 보기",
        value=False,
        key="show_home_period_details_v649",
    )

if page == "🏠 홈" and show_home_details:
    # =========================================================
    # 12. 날짜 / 기간 분석
    # =========================================================

    st.header(
        "📅 날짜 / 기간 분석"
    )

    st.write(
        "기간을 선택하면 해당 기간 동안 "
        "채널이 얼마나 성장했는지 확인할 수 있습니다."
    )


    # ---------------------------------------------------------
    # 기간 선택 + 분석하기 버튼
    # ---------------------------------------------------------

    # 처음 들어왔을 때는 최근 7일을 기본 적용
    if "applied_period_option" not in st.session_state:
        st.session_state.applied_period_option = "최근 7일"
        st.session_state.applied_start_date = today - timedelta(days=7)
        st.session_state.applied_end_date = today - timedelta(days=1)

    period_option = st.radio(
        "분석 기간",
        ["오늘", "최근 7일", "최근 28일", "직접 선택"],
        horizontal=True,
        index=1,
        key="period_option_input",
    )

    selected_range = None
    if period_option == "직접 선택":
        selected_range = st.date_input(
            "날짜 범위 선택",
            value=(today - timedelta(days=6), today),
            key="period_custom_range",
        )

    if period_option == "오늘":
        candidate_start = today
        candidate_end = today
    elif period_option == "최근 7일":
        candidate_end = today - timedelta(days=1)
        candidate_start = candidate_end - timedelta(days=6)
    elif period_option == "최근 28일":
        candidate_end = today - timedelta(days=1)
        candidate_start = candidate_end - timedelta(days=27)
    else:
        if isinstance(selected_range, tuple) and len(selected_range) == 2:
            candidate_start, candidate_end = selected_range
        elif selected_range:
            candidate_start = selected_range
            candidate_end = selected_range
        else:
            candidate_start = today
            candidate_end = today

    if st.button(
        "🔍 분석하기",
        type="primary",
        use_container_width=True,
        key="apply_period_button",
    ):
        st.session_state.applied_period_option = period_option
        st.session_state.applied_start_date = candidate_start
        st.session_state.applied_end_date = candidate_end

    period_option = st.session_state.applied_period_option
    start_date = st.session_state.applied_start_date
    end_date = st.session_state.applied_end_date

    st.caption(
        f"적용된 분석 기간: {start_date} ~ {end_date}"
    )

    if period_option == "오늘":
        st.info(
            "⏳ 오늘 데이터는 YouTube Analytics에서 아직 집계 중일 수 있습니다. "
            "오늘 수치가 0으로 보여도 실제 조회가 없는 뜻은 아닐 수 있으며, "
            "확정 데이터는 시간이 지나면서 반영됩니다."
        )


    # =========================================================
    # 13. 현재 기간 / 이전 기간
    # =========================================================

    try:

        try:
            current_summary = get_period_summary(
                yt_analytics,
                start_date,
                end_date,
            )
        except Exception:
            time.sleep(1)
            current_summary = get_period_summary(
                yt_analytics,
                start_date,
                end_date,
            )

        (
            previous_start,
            previous_end,
        ) = get_previous_period(
            start_date,
            end_date,
        )

        try:
            previous_summary = get_period_summary(
                yt_analytics,
                previous_start,
                previous_end,
            )
        except Exception:
            time.sleep(1)
            previous_summary = get_period_summary(
                yt_analytics,
                previous_start,
                previous_end,
            )

    except Exception as e:

        _err_text = str(e)

        if "backendError" in _err_text or "Internal error encountered" in _err_text or "HttpError 500" in _err_text:
            st.warning(
                "⚠️ YouTube Analytics 서버가 일시적으로 응답하지 않습니다. "
                "자동으로 한 번 다시 시도했지만 아직 실패했습니다. 잠시 후 다시 분석해 주세요."
            )
        else:
            st.error(
                "기간별 Analytics 데이터를 가져오지 못했습니다."
            )

        with st.expander("기술 오류 상세보기"):
            st.code(_err_text)

        st.stop()


    # =========================================================
    # 14. 변화율 함수
    # =========================================================

    def change_text(
        current,
        previous,
    ):

        change = calculate_change(
            current,
            previous
        )

        if change is None:
            return f"이전 기간 {previous:,.0f}"

        return f"{change:+.1f}%"


    # =========================================================
    # 15. 기간 성과 카드
    # =========================================================

    st.subheader(
        "📊 선택 기간 성과"
    )

    def _count_public_uploads_for_period(_start, _end):
        _count = 0
        for _video in public_videos:
            _raw = _video.get("published_raw")
            if not _raw:
                continue
            try:
                _dt = datetime.fromisoformat(
                    _raw.replace("Z", "+00:00")
                ).astimezone(KST)
                if _start <= _dt.date() <= _end:
                    _count += 1
            except Exception:
                pass
        return _count

    _previous_upload_count = _count_public_uploads_for_period(
        previous_start,
        previous_end,
    )
    _show_period_delta = _previous_upload_count > 0

    if not _show_period_delta:
        st.warning(
            "⚠️ 이전 비교기간에 공개 영상 업로드가 0개라 증감률은 숨겼습니다. "
            "현재 기간의 실제 수치만 확인하세요."
        )

    c1, c2, c3, c4 = (
        st.columns(4)
    )


    c1.metric(
        "조회수",
        f"{current_summary['views']:,}회",
        change_text(
            current_summary["views"],
            previous_summary["views"],
        ) if _show_period_delta else None,
    )


    c2.metric(
        "기간 내 순증가",
        (
            f"{current_summary['net_subscribers']:+,}명"
        ),
        change_text(
            current_summary[
                "net_subscribers"
            ],
            previous_summary[
                "net_subscribers"
            ],
        ) if _show_period_delta else None,
    )


    c3.metric(
        "시청시간",
        format_watch_time(
            current_summary[
                "watch_minutes"
            ]
        ),
        change_text(
            current_summary[
                "watch_minutes"
            ],
            previous_summary[
                "watch_minutes"
            ],
        ) if _show_period_delta else None,
    )


    c4.metric(
        "좋아요",
        f"{current_summary['likes']:,}개",
        change_text(
            current_summary["likes"],
            previous_summary["likes"],
        ) if _show_period_delta else None,
    )


    c1, c2, c3, c4 = (
        st.columns(4)
    )


    c1.metric(
        "구독자 획득",
        (
            f"+{current_summary['subscribers_gained']:,}명"
        )
    )


    c2.metric(
        "구독자 이탈",
        (
            f"-{current_summary['subscribers_lost']:,}명"
        )
    )


    c3.metric(
        "댓글",
        f"{current_summary['comments']:,}개"
    )


    c4.metric(
        "공유",
        f"{current_summary['shares']:,}회"
    )


    st.caption(
        f"비교 기간: "
        f"{previous_start} ~ {previous_end}"
    )

    st.divider()


    # =========================================================
    # 16. 일별 데이터
    # =========================================================

    try:

        daily_data = (
            get_daily_channel_data(
                yt_analytics,
                start_date,
                end_date,
            )
        )

    except Exception as e:

        daily_data = []

        st.warning("일별 데이터를 불러오지 못했습니다.")
        with st.expander("기술 오류 상세보기"):
            st.code(str(e))


    if daily_data:

        daily_df = pd.DataFrame(daily_data)
        daily_df["date"] = pd.to_datetime(daily_df["date"])

        # 오래된 날짜의 누락값만 0으로 처리하고, 최근 2일은 집계 중으로 남깁니다.
        full_dates = pd.date_range(start=start_date, end=end_date, freq="D")
        daily_df = daily_df.set_index("date").reindex(full_dates)

        metric_cols = [
            col for col in ["views", "net_subscribers"]
            if col in daily_df.columns
        ]
        recent_cutoff = pd.Timestamp(today - timedelta(days=2))

        for idx in daily_df.index:
            if idx < recent_cutoff:
                for col in metric_cols:
                    if pd.isna(daily_df.at[idx, col]):
                        daily_df.at[idx, col] = 0
            else:
                # 최근 날짜가 전부 0이면 실제 0으로 단정하지 않고 차트에서 비워 둡니다.
                if metric_cols and all(
                    pd.isna(daily_df.at[idx, col]) or float(daily_df.at[idx, col] or 0) == 0
                    for col in metric_cols
                ):
                    for col in metric_cols:
                        daily_df.at[idx, col] = float("nan")

        daily_df.index.name = "날짜"
        daily_df.index = [dt.strftime("%m/%d") for dt in daily_df.index]

        st.subheader("📈 일별 조회수")
        st.line_chart(daily_df[["views"]], use_container_width=True, height=260)

        st.subheader("👤 일별 순구독자")
        st.bar_chart(daily_df[["net_subscribers"]], use_container_width=True, height=260)
        st.caption("※ 최근 날짜의 값이 비어 있으면 0회가 아니라 YouTube Analytics 집계 중일 수 있습니다.")

    else:
        st.info("선택한 기간에 일별 Analytics 데이터가 없습니다.")

    st.divider()


    def call_with_retry(func, *args, attempts=2, delay=0.35, **kwargs):
        """YouTube API의 일시적 5xx 오류를 짧게 재시도합니다."""
        last_error = None
        for attempt in range(attempts):
            try:
                return func(*args, **kwargs)
            except Exception as exc:
                last_error = exc
                status = getattr(getattr(exc, "resp", None), "status", None)
                if status is not None and int(status) < 500:
                    raise
                if attempt < attempts - 1:
                    time.sleep(delay)
        if last_error:
            raise last_error


    # =========================================================
    # 17. 월간 성과 달력
    # =========================================================

    st.header("🗓️ 월간 성과 달력")

    # 달력은 항상 현재 날짜가 속한 달부터 시작
    if "calendar_month" not in st.session_state:
        st.session_state.calendar_month = today.replace(day=1)

    if "calendar_selected_day" not in st.session_state:
        st.session_state.calendar_selected_day = today

    calendar_month = st.session_state.calendar_month
    left, center, right = st.columns([1, 3, 1])

    with left:
        if st.button("◀ 이전 달", use_container_width=True):
            st.session_state.calendar_month = (calendar_month - timedelta(days=1)).replace(day=1)
            st.rerun()

    with center:
        st.markdown(
            f"<h3 style='text-align:center;'>{calendar_month.year}년 {calendar_month.month}월</h3>",
            unsafe_allow_html=True,
        )

    with right:
        next_month = (calendar_month.replace(day=28) + timedelta(days=4)).replace(day=1)
        if st.button(
            "다음 달 ▶",
            use_container_width=True,
            disabled=next_month > today.replace(day=1),
        ):
            st.session_state.calendar_month = next_month
            st.rerun()

    month_start = calendar_month
    last_day = calendar.monthrange(calendar_month.year, calendar_month.month)[1]
    month_end = min(date(calendar_month.year, calendar_month.month, last_day), today)

    try:
        month_daily_data = get_daily_channel_data(yt_analytics, month_start, month_end)
    except Exception:
        month_daily_data = []

    month_lookup = {item["date"]: item for item in month_daily_data}


    def make_monthly_excel(month_daily_rows, month_value, all_public_videos):
        """
        달력에서 선택한 달을 실제 분석용 엑셀로 내보냅니다.

        1시트: 월간 일별 성과
        2시트: 그달 업로드 영상
        3시트: 월간 요약
        """
        output = BytesIO()

        # -----------------------------
        # 그달 업로드 영상 정리
        # -----------------------------
        uploaded_rows = []
        upload_by_date = {}

        for video in all_public_videos:
            published_raw = video.get("published_raw")
            if not published_raw:
                continue

            try:
                published_dt = datetime.fromisoformat(
                    published_raw.replace("Z", "+00:00")
                ).astimezone(KST)
            except Exception:
                continue

            if (
                published_dt.year == month_value.year
                and published_dt.month == month_value.month
            ):
                date_key = published_dt.strftime("%Y-%m-%d")
                upload_by_date.setdefault(date_key, []).append(video.get("title", ""))

                uploaded_rows.append({
                    "업로드일(KST)": published_dt.strftime("%Y-%m-%d %H:%M"),
                    "제목": video.get("title", ""),
                    "조회수": int(video.get("views", 0)),
                    "좋아요": int(video.get("likes", 0)),
                    "댓글": int(video.get("comments", 0)),
                    "공유": int(video.get("shares", 0)),
                    "평균 시청시간(초)": round(float(video.get("avg_duration", 0)), 1),
                    "평균 시청률(%)": round(float(video.get("avg_percentage", 0)), 1),
                    "순구독자": int(video.get("net_subs", 0)),
                    "좋아요율(%)": round(float(video.get("like_rate", 0)), 2),
                    "구독전환율(%)": round(float(video.get("sub_conversion_rate", 0)), 3),
                    "성과점수": (
                        video.get("performance_score")
                        if video.get("performance_score") is not None
                        else "평가 보류"
                    ),
                    "성과등급": video.get("performance_grade", "-"),
                })

        month_upload_df = pd.DataFrame(uploaded_rows)

        # -----------------------------
        # 날짜별 Analytics 정리
        # 데이터가 없는 날짜도 0으로 포함
        # -----------------------------
        daily_lookup = {
            str(item.get("date", "")): item
            for item in month_daily_rows
        }

        last_day = calendar.monthrange(month_value.year, month_value.month)[1]
        calendar_last_date = date(month_value.year, month_value.month, last_day)

        if month_value.year == today.year and month_value.month == today.month:
            export_last_date = today
        else:
            export_last_date = calendar_last_date

        weekday_ko = ["월", "화", "수", "목", "금", "토", "일"]
        daily_export = []

        current_date = date(month_value.year, month_value.month, 1)

        while current_date <= export_last_date:
            date_key = current_date.strftime("%Y-%m-%d")
            item = daily_lookup.get(date_key, {})

            watch_minutes = float(item.get("watch_minutes", 0) or 0)
            uploaded_titles = upload_by_date.get(date_key, [])

            daily_export.append({
                "날짜": date_key,
                "요일": weekday_ko[current_date.weekday()],
                "조회수": int(item.get("views", 0) or 0),
                "시청시간(시간)": round(watch_minutes / 60, 2),
                "좋아요": int(item.get("likes", 0) or 0),
                "댓글": int(item.get("comments", 0) or 0),
                "공유": int(item.get("shares", 0) or 0),
                "구독자 획득": int(item.get("subscribers_gained", 0) or 0),
                "구독자 이탈": int(item.get("subscribers_lost", 0) or 0),
                "순구독자": int(item.get("net_subscribers", 0) or 0),
                "업로드 영상 수": len(uploaded_titles),
                "업로드 영상": " / ".join(uploaded_titles),
            })

            current_date += timedelta(days=1)

        month_daily_df = pd.DataFrame(daily_export)

        # 맨 아래 합계 행
        if not month_daily_df.empty:
            total_row = {
                "날짜": "합계",
                "요일": "",
                "조회수": int(month_daily_df["조회수"].sum()),
                "시청시간(시간)": round(float(month_daily_df["시청시간(시간)"].sum()), 2),
                "좋아요": int(month_daily_df["좋아요"].sum()),
                "댓글": int(month_daily_df["댓글"].sum()),
                "공유": int(month_daily_df["공유"].sum()),
                "구독자 획득": int(month_daily_df["구독자 획득"].sum()),
                "구독자 이탈": int(month_daily_df["구독자 이탈"].sum()),
                "순구독자": int(month_daily_df["순구독자"].sum()),
                "업로드 영상 수": int(month_daily_df["업로드 영상 수"].sum()),
                "업로드 영상": "",
            }
            month_daily_df = pd.concat(
                [month_daily_df, pd.DataFrame([total_row])],
                ignore_index=True,
            )

        # -----------------------------
        # 월간 요약
        # -----------------------------
        data_only_df = month_daily_df[month_daily_df["날짜"] != "합계"].copy()

        summary_df = pd.DataFrame([
            ["채널명", channel_info.get("channel_name", channel_info.get("title", ""))],
            ["대상 월", f"{month_value.year}-{month_value.month:02d}"],
            ["월 조회수", int(data_only_df["조회수"].sum()) if not data_only_df.empty else 0],
            ["월 순구독자", int(data_only_df["순구독자"].sum()) if not data_only_df.empty else 0],
            ["월 시청시간(시간)", round(float(data_only_df["시청시간(시간)"].sum()), 2) if not data_only_df.empty else 0],
            ["월 좋아요", int(data_only_df["좋아요"].sum()) if not data_only_df.empty else 0],
            ["월 댓글", int(data_only_df["댓글"].sum()) if not data_only_df.empty else 0],
            ["월 공유", int(data_only_df["공유"].sum()) if not data_only_df.empty else 0],
            ["그달 업로드 영상", len(month_upload_df)],
            ["생성 시각", datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S KST")],
        ], columns=["항목", "값"])

        # -----------------------------
        # Excel 출력
        # -----------------------------
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            # 열자마자 상세 데이터가 먼저 보이게 함
            month_daily_df.to_excel(writer, sheet_name="월간 일별 성과", index=False)
            month_upload_df.to_excel(writer, sheet_name="그달 업로드 영상", index=False)
            summary_df.to_excel(writer, sheet_name="월간 요약", index=False)

            for sheet_name in writer.book.sheetnames:
                ws = writer.book[sheet_name]
                ws.freeze_panes = "A2"
                ws.auto_filter.ref = ws.dimensions

                header_map = {
                    cell.value: cell.column_letter
                    for cell in ws[1]
                    if cell.value is not None
                }

                for column_cells in ws.columns:
                    max_length = 0
                    column_letter = column_cells[0].column_letter

                    for cell in column_cells:
                        try:
                            cell_length = len(str(cell.value)) if cell.value is not None else 0
                            max_length = max(max_length, cell_length)
                        except Exception:
                            pass

                    ws.column_dimensions[column_letter].width = min(
                        max(max_length + 3, 11),
                        48,
                    )

                # 날짜는 무조건 문자열로
                for header in ["날짜", "업로드일(KST)"]:
                    if header in header_map:
                        col = header_map[header]
                        for row in range(2, ws.max_row + 1):
                            ws[f"{col}{row}"].number_format = "@"

                if "날짜" in header_map:
                    ws.column_dimensions[header_map["날짜"]].width = 14

                if "업로드일(KST)" in header_map:
                    ws.column_dimensions[header_map["업로드일(KST)"]].width = 21

                if "제목" in header_map:
                    ws.column_dimensions[header_map["제목"]].width = 42

                if "업로드 영상" in header_map:
                    ws.column_dimensions[header_map["업로드 영상"]].width = 48

            for idx, sheet_name in enumerate(writer.book.sheetnames, start=1):
                add_excel_table(writer.book[sheet_name], f"MonthlyTable_{idx}")

        output.seek(0)
        return output.getvalue()


    # 공개 영상을 올린 날짜(KST)
    upload_dates = set()
    for video in public_videos:
        published_raw = video.get("published_raw")
        if not published_raw:
            continue
        try:
            published_dt = datetime.fromisoformat(
                published_raw.replace("Z", "+00:00")
            ).astimezone(KST)
            upload_dates.add(published_dt.date())
        except Exception:
            pass

    weekday_names = ["월", "화", "수", "목", "금", "토", "일"]
    for column, name in zip(st.columns(7), weekday_names):
        column.markdown(f"<div style='text-align:center;'><b>{name}</b></div>", unsafe_allow_html=True)

    for week in calendar.monthcalendar(calendar_month.year, calendar_month.month):
        columns = st.columns(7)
        for day_index, day_number in enumerate(week):
            if day_number == 0:
                columns[day_index].write("")
                continue

            current_day = date(calendar_month.year, calendar_month.month, day_number)
            day_key = current_day.strftime("%Y-%m-%d")
            has_analytics_row = day_key in month_lookup
            day_data = month_lookup.get(day_key, {})
            views = int(day_data.get("views", 0))
            net_subscribers = int(day_data.get("net_subscribers", 0))

            if current_day > today:
                columns[day_index].button(
                    f"{day_number}일\n\n-",
                    key=f"future_{current_day}",
                    disabled=True,
                    use_container_width=True,
                )
                continue

            if views >= 10000:
                view_text = f"{views / 10000:.1f}만"
            elif views >= 1000:
                view_text = f"{views / 1000:.1f}천"
            else:
                view_text = f"{views:,}"

            subscriber_text = f"\n👤 {net_subscribers:+d}" if net_subscribers != 0 else ""
            upload_mark = " 🎬" if current_day in upload_dates else ""

            # 최근 2일은 Analytics 행 자체가 없으면 0으로 단정하지 않음
            is_recent_calendar_day = 0 <= (today - current_day).days <= 2
            if is_recent_calendar_day and not has_analytics_row:
                label = f"{day_number}일{upload_mark}\n\n⏳ 집계 중"
            elif views == 0 and net_subscribers == 0:
                label = f"{day_number}일{upload_mark}"
            else:
                label = f"{day_number}일{upload_mark}\n\n👁 {view_text}{subscriber_text}"

            if columns[day_index].button(
                label,
                key=f"calendar_{current_day}",
                use_container_width=True,
            ):
                st.session_state.calendar_selected_day = current_day
                st.rerun()

    st.caption("날짜를 누르면 아래에서 그날의 상세 성과를 확인할 수 있습니다.")

    st.download_button(
        f"📥 {calendar_month.year}년 {calendar_month.month}월 성과 엑셀",
        data=make_monthly_excel(month_daily_data, calendar_month, public_videos),
        file_name=f"shorts_monthly_{calendar_month.year}_{calendar_month.month:02d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        key=f"monthly_excel_{calendar_month.year}_{calendar_month.month}",
    )

    st.divider()

    # =========================================================
    # 18. 특정 날짜 선택
    # =========================================================

    st.header(
        "🕘 과거 일별 분석"
    )

    # 최근 날짜는 Analytics가 늦더라도 스냅샷 임시값으로 먼저 확인할 수 있으므로 오늘까지 허용
    _detail_max_day = today

    # 달력/이전 세션에 오늘·어제 값이 남아 있어도
    # date_input의 최대값을 넘지 않도록 모두 안전하게 보정
    _detail_default_day = st.session_state.get(
        "calendar_selected_day",
        end_date,
    )
    if _detail_default_day > _detail_max_day:
        _detail_default_day = _detail_max_day

    if "applied_detail_day" not in st.session_state:
        st.session_state.applied_detail_day = _detail_default_day
    elif st.session_state.applied_detail_day > _detail_max_day:
        st.session_state.applied_detail_day = _detail_max_day

    if "detail_day_input" in st.session_state:
        _saved_detail_day = st.session_state["detail_day_input"]
        if _saved_detail_day is not None and _saved_detail_day > _detail_max_day:
            st.session_state["detail_day_input"] = _detail_max_day

    detail_day_input = st.date_input(
        "확인할 날짜",
        value=_detail_default_day,
        min_value=date(2005, 1, 1),
        max_value=_detail_max_day,
        key="detail_day_input",
    )
    st.session_state.calendar_selected_day = detail_day_input

    if st.button(
        "🔍 조회하기",
        type="primary",
        use_container_width=True,
        key="apply_detail_day_button",
    ):
        st.session_state.applied_detail_day = detail_day_input

    selected_day = st.session_state.applied_detail_day
    st.caption(f"현재 조회 중인 날짜: {selected_day}")

    # 최근 날짜는 YouTube Analytics의 일별 집계가 아직 완료되지 않았을 수 있음
    detail_age_days = (today - selected_day).days
    is_recent_detail = 0 <= detail_age_days <= 2
    if is_recent_detail:
        st.warning(
            "⏳ 최근 날짜의 YouTube Analytics는 아직 집계 중일 수 있습니다. "
            "확정값이 없으면 Shorts Scope 스냅샷으로 계산 가능한 조회수·좋아요를 ⚡ 임시값으로 먼저 표시합니다. "
            "순구독자와 시청시간은 Analytics 확정 전까지 집계 중으로 표시합니다."
        )


    # ---------------------------------------------------------
    # V6.6.2 — 최근 날짜: Analytics 우선 + 스냅샷 임시값 fallback
    # ---------------------------------------------------------
    try:
        day_summary = get_period_summary(
            yt_analytics,
            selected_day,
            selected_day,
        )
        day_summary_error = None
    except Exception as e:
        day_summary = None
        day_summary_error = str(e)

    _recent_all_zero = bool(
        is_recent_detail
        and day_summary
        and day_summary.get("views", 0) == 0
        and day_summary.get("watch_minutes", 0) == 0
        and day_summary.get("likes", 0) == 0
        and day_summary.get("net_subscribers", 0) == 0
    )

    _snapshot_day = {
        "ok": False,
        "reason": "not_needed",
        "views": None,
        "likes": None,
        "comments": None,
        "new_video_views": None,
        "old_video_views": None,
        "video_rows": [],
        "coverage": 0,
        "partial": False,
        "start_at": None,
        "end_at": None,
    }

    if is_recent_detail and (day_summary is None or _recent_all_zero):
        try:
            _day_channel_id = _connected_youtube_channel_id(youtube)

            _day_start_kst = datetime.combine(
                selected_day,
                datetime.min.time(),
                tzinfo=KST,
            )
            _day_next_kst = _day_start_kst + timedelta(days=1)
            _now_kst = datetime.now(KST)
            _day_end_kst = min(_day_next_kst, _now_kst)

            _fetch_since_utc = (
                _day_start_kst.astimezone(timezone.utc) - timedelta(hours=2)
            ).isoformat()

            _day_snapshot_fetch = _fetch_channel_snapshots(
                _day_channel_id,
                _fetch_since_utc,
            )

            if _day_snapshot_fetch.get("ok"):
                _all_grouped = _group_snapshot_rows(
                    _day_snapshot_fetch.get("rows", []),
                    [v.get("video_id") for v in public_videos],
                )

                _start_target = _day_start_kst.astimezone(timezone.utc)
                _end_target = _day_end_kst.astimezone(timezone.utc)
                _tolerance = timedelta(minutes=90)

                def _nearest_snapshot(_rows, _target, _prefer_before=False):
                    if not _rows:
                        return None
                    if _prefer_before:
                        _before = [r for r in _rows if r["captured_at"] <= _target]
                        if _before:
                            _candidate = _before[-1]
                            if (_target - _candidate["captured_at"]) <= _tolerance:
                                return _candidate
                    _candidate = min(
                        _rows,
                        key=lambda r: abs((r["captured_at"] - _target).total_seconds()),
                    )
                    if abs(_candidate["captured_at"] - _target) <= _tolerance:
                        return _candidate
                    return None

                _snapshot_video_rows = []
                _total_views = 0
                _total_likes = 0
                _total_comments = 0
                _new_views = 0
                _old_views = 0
                _coverage = 0

                for _video in public_videos:
                    _vid = _video.get("video_id")
                    _rows = _all_grouped.get(_vid, [])
                    if not _rows:
                        continue

                    _published_dt = _parse_utc(_video.get("published_raw"))
                    _published_kst = (
                        _published_dt.astimezone(KST)
                        if _published_dt else None
                    )
                    _uploaded_that_day = bool(
                        _published_kst
                        and _published_kst.date() == selected_day
                    )

                    _end_row = _nearest_snapshot(
                        _rows,
                        _end_target,
                        _prefer_before=True,
                    )
                    if _end_row is None:
                        continue

                    if _uploaded_that_day:
                        _start_views = 0
                        _start_likes = 0
                        _start_comments = 0
                    else:
                        _start_row = _nearest_snapshot(
                            _rows,
                            _start_target,
                            _prefer_before=False,
                        )
                        if _start_row is None:
                            continue
                        _start_views = _start_row["view_count"]
                        _start_likes = _start_row.get("like_count", 0)
                        _start_comments = _start_row.get("comment_count", 0)

                    _view_gain = max(_end_row["view_count"] - _start_views, 0)
                    _like_gain = max(_end_row.get("like_count", 0) - _start_likes, 0)
                    _comment_gain = max(_end_row.get("comment_count", 0) - _start_comments, 0)

                    _total_views += _view_gain
                    _total_likes += _like_gain
                    _total_comments += _comment_gain
                    _coverage += 1

                    if _uploaded_that_day:
                        _new_views += _view_gain
                    else:
                        _old_views += _view_gain

                    _snapshot_video_rows.append({
                        "video_id": _vid,
                        "views": _view_gain,
                        "likes": _like_gain,
                        "comments": _comment_gain,
                    })

                _snapshot_video_rows.sort(
                    key=lambda x: x["views"],
                    reverse=True,
                )

                if _coverage > 0:
                    _snapshot_day = {
                        "ok": True,
                        "reason": None,
                        "views": _total_views,
                        "likes": _total_likes,
                        "comments": _total_comments,
                        "new_video_views": _new_views,
                        "old_video_views": _old_views,
                        "video_rows": _snapshot_video_rows,
                        "coverage": _coverage,
                        "partial": selected_day == today,
                        "start_at": _start_target,
                        "end_at": _end_target,
                    }
                else:
                    _snapshot_day["reason"] = "insufficient_boundary_data"
            else:
                _snapshot_day["reason"] = _day_snapshot_fetch.get("reason")
        except Exception:
            _snapshot_day["reason"] = "snapshot_error"

    _analytics_confirmed = bool(day_summary and not _recent_all_zero)
    _use_snapshot_day = bool(
        not _analytics_confirmed
        and _snapshot_day.get("ok")
    )

    st.subheader(f"📅 {selected_day}")

    c1, c2, c3, c4 = st.columns(4)

    if _analytics_confirmed:
        c1.metric("그날 조회수", f"+{day_summary['views']:,}회")
        c2.metric("그날 순구독자", f"{day_summary['net_subscribers']:+,}명")
        c3.metric("그날 시청시간", format_watch_time(day_summary["watch_minutes"]))
        c4.metric("그날 좋아요", f"+{day_summary['likes']:,}개")
        st.caption("✅ YouTube Analytics 확정 데이터")

    elif _use_snapshot_day:
        c1.metric("그날 조회수", f"⚡ +{_snapshot_day['views']:,}회")
        c2.metric("그날 순구독자", "⏳ 집계 중")
        c3.metric("그날 시청시간", "⏳ 집계 중")
        c4.metric("그날 좋아요", f"⚡ +{_snapshot_day['likes']:,}개")

        if _snapshot_day.get("partial"):
            st.info(
                "⚡ 오늘 수치는 현재까지 쌓인 Shorts Scope 스냅샷 기준 임시값입니다. "
                "YouTube Analytics가 확정되면 자동으로 확정값을 우선 표시합니다."
            )
        else:
            st.info(
                "⚡ 아직 YouTube Analytics가 확정되지 않아 Shorts Scope 스냅샷 기준 임시값을 표시합니다. "
                "확정 Analytics가 들어오면 자동으로 교체됩니다."
            )
        st.caption(
            f"스냅샷 계산 가능 영상 {_snapshot_day['coverage']}개 · "
            "순구독자와 시청시간은 YouTube Analytics 확정 대기"
        )

    else:
        c1.metric("그날 조회수", "⏳ 집계 중")
        c2.metric("그날 순구독자", "⏳ 집계 중")
        c3.metric("그날 시청시간", "⏳ 집계 중")
        c4.metric("그날 좋아요", "⏳ 집계 중")
        st.info(
            "📌 아직 확정 Analytics가 없고, 이 날짜를 계산할 만큼 스냅샷 경계 데이터도 부족합니다. "
            "스냅샷 수집이 계속되면 최근 날짜부터 임시값을 먼저 확인할 수 있습니다."
        )

    if day_summary_error:
        with st.expander("기술 오류 상세보기"):
            st.code(day_summary_error)

    day_video_data = []
    _day_video_source = None

    if _analytics_confirmed:
        try:
            day_video_data = get_video_performance_for_day(
                yt_analytics,
                selected_day,
            )
            _day_video_source = "analytics"
        except Exception as e:
            st.warning(
                "⏳ 그날의 영상별 Analytics를 일시적으로 불러오지 못했습니다."
            )
            with st.expander("기술 오류 상세보기"):
                st.code(str(e))
    elif _use_snapshot_day:
        day_video_data = _snapshot_day.get("video_rows", [])
        _day_video_source = "snapshot"

    video_lookup = {
        video["video_id"]: video
        for video in videos
    }

    uploaded_ids_for_day = set()
    for video in public_videos:
        published_raw = video.get("published_raw")
        if not published_raw:
            continue
        try:
            published_dt = datetime.fromisoformat(
                published_raw.replace("Z", "+00:00")
            ).astimezone(KST)
            if published_dt.date() == selected_day:
                uploaded_ids_for_day.add(video["video_id"])
        except Exception:
            pass

    if _analytics_confirmed:
        new_video_views = sum(
            item["views"] for item in day_video_data
            if item["video_id"] in uploaded_ids_for_day
        )
        day_total_views = max(day_summary["views"], 0)
        old_video_views = max(day_total_views - new_video_views, 0)
    elif _use_snapshot_day:
        new_video_views = int(_snapshot_day.get("new_video_views") or 0)
        old_video_views = int(_snapshot_day.get("old_video_views") or 0)
        day_total_views = new_video_views + old_video_views
    else:
        new_video_views = 0
        old_video_views = 0
        day_total_views = 0

    st.subheader("🧩 그날 조회수 구성")

    if day_total_views > 0:
        new_video_share = (new_video_views / day_total_views) * 100
        old_video_share = (old_video_views / day_total_views) * 100
    else:
        new_video_share = 0.0
        old_video_share = 0.0

    cc1, cc2 = st.columns(2)

    if day_total_views > 0:
        _prefix = "⚡ " if _use_snapshot_day else ""
        cc1.metric(
            "당일 업로드 영상",
            f"{_prefix}{new_video_views:,}회",
            f"{new_video_share:.1f}% 기여",
        )
        cc2.metric(
            "기존 영상",
            f"{_prefix}{old_video_views:,}회",
            f"{old_video_share:.1f}% 기여",
        )
        if _use_snapshot_day:
            st.caption("⚡ 스냅샷 기준 임시 조회수 구성")
    else:
        cc1.metric("당일 업로드 영상", "⏳ 집계 중")
        cc2.metric("기존 영상", "⏳ 집계 중")
        st.info(
            "📌 조회수 구성도 아직 계산할 데이터가 부족합니다."
        )

    if day_video_data:
        st.subheader("🔥 그날 조회수를 만든 영상")

        if _day_video_source == "snapshot":
            st.caption(
                "⚡ 아래 순위는 Shorts Scope 스냅샷 기준 임시값입니다. "
                "Analytics 확정 후 공식 일별 값이 우선 표시됩니다."
            )

        for rank, performance in enumerate(
            day_video_data[:5],
            start=1,
        ):
            video = video_lookup.get(performance["video_id"])
            if not video:
                continue

            col_img, col_info = st.columns([1, 5])

            with col_img:
                if video["thumbnail"]:
                    st.image(video["thumbnail"], width=140)

            with col_info:
                st.markdown(f"### {rank}위 · {video['title']}")

                if _day_video_source == "analytics":
                    st.write(
                        f"👁️ 그날 +{performance['views']:,}회"
                        f"  |  👍 +{performance['likes']:,}"
                        f"  |  👤 {performance['net_subscribers']:+d}"
                    )
                else:
                    st.write(
                        f"⚡ 조회수 +{performance['views']:,}회"
                        f"  |  👍 +{performance.get('likes', 0):,}"
                        f"  |  💬 +{performance.get('comments', 0):,}"
                    )

        st.divider()


    # =========================================================
    # 18. 그날 업로드한 영상
    # =========================================================

    uploaded_that_day = []


    for video in public_videos:

        published_raw = (
            video.get(
                "published_raw"
            )
        )

        if not published_raw:
            continue

        try:

            published_dt = (
                datetime.fromisoformat(
                    published_raw.replace(
                        "Z",
                        "+00:00"
                    )
                )
            )

            published_dt = published_dt.astimezone(KST)

            if (
                published_dt.date()
                == selected_day
            ):

                uploaded_that_day.append(
                    video
                )

        except Exception:
            pass


    st.subheader(
        "🎬 그날 업로드한 영상"
    )


    if uploaded_that_day:

        for video in uploaded_that_day:

            col_img, col_info = (
                st.columns(
                    [1, 5]
                )
            )

            with col_img:

                if video[
                    "thumbnail"
                ]:

                    st.image(
                        video[
                            "thumbnail"
                        ],
                        width=140
                    )


            with col_info:

                st.markdown(
                    f"**{video['title']}**"
                )

                st.write(
                    f"현재 조회수 "
                    f"{video['views']:,}회"
                )

                st.write(
                    f"영상 길이 "
                    f"{video['duration']}"
                )

    else:

        st.write(
            "이날 업로드한 영상이 없습니다."
        )


    st.divider()


if page == "📊 채널 패턴":
    def _pattern_publish_dt(video):
        raw = video.get("published_raw")
        if not raw:
            return None
        try:
            return datetime.fromisoformat(raw.replace("Z", "+00:00")).astimezone(KST)
        except Exception:
            return None

    _pattern_videos = []
    for _video in public_videos:
        _dt = _pattern_publish_dt(_video)
        if _dt is None:
            continue
        _copy = dict(_video)
        _copy["_pattern_dt"] = _dt
        _copy["_analytics_ready"] = _video.get("video_id") in video_analytics
        _pattern_videos.append(_copy)

    _pattern_videos.sort(key=lambda x: x["_pattern_dt"], reverse=True)

    if not _pattern_videos:
        st.info("분석할 공개 영상이 없습니다.")
    else:
        _max_recent = len(_pattern_videos)

        if "v77_pattern_range_mode" not in st.session_state:
            st.session_state.v77_pattern_range_mode = "최근 20개"
        if "v77_pattern_custom_n" not in st.session_state:
            st.session_state.v77_pattern_custom_n = min(100, _max_recent)

        _range_options = ["최근 10개", "최근 20개", "최근 50개", "최근 100개", "전체", "직접 입력"]

        with st.form("v77_pattern_range_form"):
            _pr1, _pr2 = st.columns([1.4, 1])
            with _pr1:
                _range_mode_input = st.selectbox(
                    "분석할 최근 영상 범위",
                    options=_range_options,
                    index=_range_options.index(st.session_state.v77_pattern_range_mode)
                    if st.session_state.v77_pattern_range_mode in _range_options
                    else 1,
                    key="v77_pattern_range_mode_input",
                )
            with _pr2:
                _custom_n_input = st.number_input(
                    "직접 입력",
                    min_value=1,
                    max_value=max(_max_recent, 1),
                    value=min(
                        int(st.session_state.v77_pattern_custom_n or 1),
                        max(_max_recent, 1),
                    ),
                    step=10,
                    disabled=_range_mode_input != "직접 입력",
                    key="v77_pattern_custom_n_input",
                )

            _apply_recent_n = st.form_submit_button(
                "범위 적용",
                use_container_width=True,
            )

        if _apply_recent_n:
            st.session_state.v77_pattern_range_mode = _range_mode_input
            st.session_state.v77_pattern_custom_n = int(_custom_n_input)

        _range_mode = st.session_state.v77_pattern_range_mode

        if _range_mode == "전체":
            _recent_n = _max_recent
        elif _range_mode == "직접 입력":
            _recent_n = min(int(st.session_state.v77_pattern_custom_n), _max_recent)
        else:
            _recent_n = min(int(re.sub(r"\D", "", _range_mode)), _max_recent)

        st.caption(
            f"현재 적용: 최근 {_recent_n:,}개 / 분석 가능한 공개 영상 {_max_recent:,}개"
        )
        _recent = _pattern_videos[: int(_recent_n)]
        _recent_analytics = [v for v in _recent if v.get("_analytics_ready")]

        st.caption(
            f"최근 업로드 {len(_recent)}개 기준 · "
            f"Analytics 확인 가능 {len(_recent_analytics)}개"
        )

        _recent_views = [int(v.get("views", 0) or 0) for v in _recent]
        _views_avg = sum(_recent_views) / len(_recent_views) if _recent_views else 0
        _views_median = float(pd.Series(_recent_views).median()) if _recent_views else 0
        _best_video = max(_recent, key=lambda v: int(v.get("views", 0) or 0))

        _m1, _m2, _m3 = st.columns(3)
        _m1.metric("평균 조회수", f"{_views_avg:,.0f}회")
        _m2.metric("중앙 조회수", f"{_views_median:,.0f}회")
        _m3.metric("최고 조회수", f"{int(_best_video.get('views', 0) or 0):,}회")
        st.caption(f"최고 영상: {_best_video.get('title', '제목 없음')}")

        if _recent_analytics:
            _ret_avg = sum(float(v.get("avg_percentage", 0) or 0) for v in _recent_analytics) / len(_recent_analytics)
            _like_avg = sum(float(v.get("like_rate", 0) or 0) for v in _recent_analytics) / len(_recent_analytics)
            _sub_avg = sum(float(v.get("sub_conversion_rate", 0) or 0) for v in _recent_analytics) / len(_recent_analytics)

            _a1, _a2, _a3 = st.columns(3)
            _a1.metric("평균 시청률", f"{_ret_avg:.1f}%")
            _a2.metric("평균 좋아요율", f"{_like_avg:.2f}%")
            _a3.metric("평균 구독전환율", f"{_sub_avg:.3f}%")
        else:
            st.caption("⏳ 최근 영상의 Analytics가 아직 충분히 집계되지 않았습니다.")

        st.divider()
        st.markdown("### 요일별 성과")
        st.caption(
            "현재 앱 시간대(KST)의 업로드 요일 기준입니다. "
            "표본이 1개뿐인 요일은 패턴으로 단정하지 않습니다."
        )

        _weekday_order = ["월", "화", "수", "목", "금", "토", "일"]
        _weekday_rows = []
        for _wd in _weekday_order:
            _group = [v for v in _recent if _weekday_order[v["_pattern_dt"].weekday()] == _wd]
            if not _group:
                continue
            _views = [int(v.get("views", 0) or 0) for v in _group]
            _ready = [v for v in _group if v.get("_analytics_ready")]
            _weekday_rows.append({
                "요일": f"{_wd}요일",
                "영상 수": len(_group),
                "평균 조회수": round(sum(_views) / len(_views)),
                "중앙 조회수": round(float(pd.Series(_views).median())),
                "평균 시청률": (
                    round(sum(float(v.get("avg_percentage", 0) or 0) for v in _ready) / len(_ready), 1)
                    if _ready else None
                ),
                "판단": "비교 가능" if len(_group) >= 2 else "표본 부족",
            })

        if _weekday_rows:
            _weekday_df = pd.DataFrame(_weekday_rows)
            st.dataframe(_weekday_df, hide_index=True, use_container_width=True)
            _weekday_chart = _weekday_df[_weekday_df["영상 수"] >= 2][["요일", "평균 조회수"]]
            if not _weekday_chart.empty:
                st.markdown("#### 평균 조회수 비교")
                _weekday_max = max(int(x) for x in _weekday_chart["평균 조회수"].tolist()) or 1
                for _, _row in _weekday_chart.iterrows():
                    _label_col, _bar_col, _value_col = st.columns([1.2, 5, 1.4])
                    _value = int(_row["평균 조회수"])
                    _label_col.write(str(_row["요일"]))
                    _bar_col.progress(min(max(_value / _weekday_max, 0.0), 1.0))
                    _value_col.write(f"{_value:,}회")
            else:
                st.caption("⏳ 요일별 비교를 하기에는 아직 표본이 부족합니다.")

        st.divider()
        st.markdown("### 업로드 시간대별 성과")
        st.caption(
            "현재 앱 시간대(KST) 기준입니다. "
            "업로드 수가 적은 시간대는 '표본 부족'으로 구분합니다."
        )

        def _time_band(hour):
            if 0 <= hour < 6:
                return "새벽 00~05시"
            if 6 <= hour < 12:
                return "오전 06~11시"
            if 12 <= hour < 18:
                return "오후 12~17시"
            return "저녁 18~23시"

        _band_order = ["새벽 00~05시", "오전 06~11시", "오후 12~17시", "저녁 18~23시"]
        _band_rows = []
        for _band in _band_order:
            _group = [v for v in _recent if _time_band(v["_pattern_dt"].hour) == _band]
            if not _group:
                continue
            _views = [int(v.get("views", 0) or 0) for v in _group]
            _ready = [v for v in _group if v.get("_analytics_ready")]
            _band_rows.append({
                "시간대": _band,
                "영상 수": len(_group),
                "평균 조회수": round(sum(_views) / len(_views)),
                "중앙 조회수": round(float(pd.Series(_views).median())),
                "평균 시청률": (
                    round(sum(float(v.get("avg_percentage", 0) or 0) for v in _ready) / len(_ready), 1)
                    if _ready else None
                ),
                "판단": "비교 가능" if len(_group) >= 2 else "표본 부족",
            })

        if _band_rows:
            _band_df = pd.DataFrame(_band_rows)
            _pattern_cards = []

            _weekday_compare = [row for row in _weekday_rows if row["영상 수"] >= 2]
            if _weekday_compare:
                _best_weekday = max(_weekday_compare, key=lambda row: row["평균 조회수"])
                _weekday_conf = "낮음" if _best_weekday["영상 수"] < 3 else "보통" if _best_weekday["영상 수"] < 6 else "높음"
                _weekday_cls = "v76-warning" if _weekday_conf == "낮음" else "v76-info" if _weekday_conf == "보통" else "v76-positive"
                _pattern_cards.append(
                    (
                        f'<div class="v76-insight">'
                        f'<div class="v76-insight-top">'
                        f'<div class="v76-insight-label">요일 관찰</div>'
                        f'<span class="v76-badge {_weekday_cls}">신뢰도 {_weekday_conf}</span>'
                        f'</div>'
                        f'<div class="v76-insight-value">{_best_weekday["요일"]}</div>'
                        f'<div class="v76-insight-desc">'
                        f'현재 표본 중 평균 조회수 {_best_weekday["평균 조회수"]:,}회로 가장 높음 · '
                        f'표본 {_best_weekday["영상 수"]}개'
                        f'</div>'
                        f'</div>'
                    )
                )

            _band_compare = [row for row in _band_rows if row["영상 수"] >= 2]
            if _band_compare:
                _best_band_row = max(_band_compare, key=lambda row: row["평균 조회수"])
                _band_conf = "낮음" if _best_band_row["영상 수"] < 3 else "보통" if _best_band_row["영상 수"] < 6 else "높음"
                _band_cls = "v76-warning" if _band_conf == "낮음" else "v76-info" if _band_conf == "보통" else "v76-positive"
                _pattern_cards.append(
                    (
                        f'<div class="v76-insight">'
                        f'<div class="v76-insight-top">'
                        f'<div class="v76-insight-label">시간대 관찰</div>'
                        f'<span class="v76-badge {_band_cls}">신뢰도 {_band_conf}</span>'
                        f'</div>'
                        f'<div class="v76-insight-value">{_best_band_row["시간대"]}</div>'
                        f'<div class="v76-insight-desc">'
                        f'현재 표본 중 평균 조회수 {_best_band_row["평균 조회수"]:,}회로 가장 높음 · '
                        f'표본 {_best_band_row["영상 수"]}개'
                        f'</div>'
                        f'</div>'
                    )
                )

            if _pattern_cards:
                st.markdown("#### 현재 데이터에서 보이는 패턴")
                st.markdown('<div class="v76-insight-grid">' + "".join(_pattern_cards) + '</div>', unsafe_allow_html=True)
                st.caption("※ 현재 표본에서 관찰된 차이이며 업로드 시간이나 요일이 성과의 원인이라고 단정하지 않습니다.")

            st.dataframe(_band_df, hide_index=True, use_container_width=True)
            _band_chart = _band_df[_band_df["영상 수"] >= 2][["시간대", "평균 조회수"]]
            if not _band_chart.empty:
                st.markdown("#### 평균 조회수 비교")
                _band_max = max(int(x) for x in _band_chart["평균 조회수"].tolist()) or 1
                for _, _row in _band_chart.iterrows():
                    _label_col, _bar_col, _value_col = st.columns([1.8, 5, 1.4])
                    _value = int(_row["평균 조회수"])
                    _label_col.write(str(_row["시간대"]))
                    _bar_col.progress(min(max(_value / _band_max, 0.0), 1.0))
                    _value_col.write(f"{_value:,}회")
            else:
                st.caption("⏳ 시간대별 비교를 하기에는 아직 표본이 부족합니다.")

        st.divider()
        st.markdown("### 소재 · 주제별 패턴")
        st.info(
            "제목만 보고 소재를 AI처럼 추정하지 않습니다. "
            "V6.7 운영에서 사용자가 영상 태그/주제를 기록할 수 있게 만든 뒤, "
            "그 실제 태그를 이 화면과 자동 연결합니다."
        )
        st.caption(
            "즉 주제 패턴 화면의 자리는 준비하되, 근거 없는 자동 분류는 하지 않습니다."
        )

    st.divider()



if page == "🧪 운영":
    try:
        _ops_channel_id = _connected_youtube_channel_id(youtube)
    except Exception:
        _ops_channel_id = None

    if not _ops_channel_id:
        st.warning("현재 연결된 YouTube 채널을 확인하지 못했습니다.")
    elif not SUPABASE_URL or not SUPABASE_SECRET_KEY:
        st.warning("Supabase 운영 데이터 저장 설정이 연결되지 않았습니다.")
    else:
        _tab_goal, _tab_video, _tab_experiment = st.tabs(
            ["목표", "영상 기록", "실험 기록"]
        )

        # =====================================================
        # 목표
        # =====================================================
        with _tab_goal:
            st.subheader("채널 목표")
            st.caption(
                "목표는 성과 판정 기준이 아니라 운영 계획을 기록하는 용도입니다."
            )

            _goal_fetch = _supabase_table_get(
                "channel_goals",
                {
                    "select": "*",
                    "channel_id": f"eq.{_ops_channel_id}",
                    "limit": 1,
                },
            )
            _goal_row = (
                _goal_fetch.get("rows", [])[0]
                if _goal_fetch.get("ok") and _goal_fetch.get("rows")
                else {}
            )

            with st.form("v67_channel_goal_form"):
                _goal_weekly_uploads = st.number_input(
                    "주간 업로드 목표",
                    min_value=0,
                    max_value=100,
                    value=int(_goal_row.get("weekly_upload_goal") or 0),
                    step=1,
                )
                _goal_target_views = st.number_input(
                    "영상 1개 목표 조회수",
                    min_value=0,
                    value=int(_goal_row.get("target_views") or 0),
                    step=1000,
                )
                _goal_note = st.text_area(
                    "목표 메모",
                    value=str(_goal_row.get("goal_note") or ""),
                    placeholder="예: 이번 달은 업로드 빈도보다 유지율 높은 포맷 찾기에 집중",
                )

                _save_goal = st.form_submit_button(
                    "💾 목표 저장",
                    use_container_width=True,
                )

            if _save_goal:
                _result = _supabase_table_upsert(
                    "channel_goals",
                    {
                        "channel_id": _ops_channel_id,
                        "user_id": None,
                        "weekly_upload_goal": int(_goal_weekly_uploads),
                        "target_views": int(_goal_target_views),
                        "goal_note": _goal_note.strip(),
                        "updated_at": datetime.now(timezone.utc).isoformat(),
                    },
                    "channel_id",
                )
                if _result.get("ok"):
                    st.success("목표를 저장했습니다.")
                else:
                    st.error(
                        "목표 저장에 실패했습니다. 운영 테이블이 아직 만들어지지 않았을 수 있습니다."
                    )
                    if _result.get("message"):
                        with st.expander("기술 오류 상세보기"):
                            st.code(_result["message"])

        # =====================================================
        # 영상 기록
        # =====================================================
        with _tab_video:
            st.subheader("영상별 기록")
            st.caption(
                "여기 내용은 사용자가 직접 기록합니다. "
                "Shorts Scope가 근거 없이 원인이나 개선점을 자동 생성하지 않습니다."
            )

            _ops_public = sorted(
                public_videos,
                key=lambda v: v.get("published_raw") or "",
                reverse=True,
            )

            if "v72_ops_applied_search" not in st.session_state:
                st.session_state.v72_ops_applied_search = ""

            with st.form("v72_ops_video_search_form"):
                _video_search_input = st.text_input(
                    "영상 검색",
                    value=st.session_state.v72_ops_applied_search,
                    placeholder="제목 일부 입력",
                    key="v72_ops_video_search_input",
                )
                _ops_search_submit = st.form_submit_button(
                    "🔍 검색",
                    type="primary",
                    use_container_width=True,
                )

            if _ops_search_submit:
                st.session_state.v72_ops_applied_search = _video_search_input.strip().lower()

            _video_search = st.session_state.v72_ops_applied_search
            _ops_filtered = [
                v for v in _ops_public
                if not _video_search
                or _video_search in str(v.get("title", "")).lower()
            ]

            if not _ops_filtered:
                st.info("검색 조건에 맞는 영상이 없습니다.")
            else:
                _video_option_map = {}
                _video_options = []
                for _v in _ops_filtered:
                    _raw = _v.get("published_raw")
                    _date_text = ""
                    if _raw:
                        try:
                            _date_text = datetime.fromisoformat(
                                _raw.replace("Z", "+00:00")
                            ).astimezone(KST).strftime("%Y.%m.%d")
                        except Exception:
                            pass
                    _label = (
                        f"{_date_text} | {_v.get('title', '제목 없음')} "
                        f"| {int(_v.get('views', 0) or 0):,}회"
                    )
                    _key = f"{_label} [{_v.get('video_id')}]"
                    _video_options.append(_key)
                    _video_option_map[_key] = _v

                _available_ids = [v.get("video_id") for v in _ops_filtered]
                if (
                    "v72_ops_applied_video_id" not in st.session_state
                    or st.session_state.v72_ops_applied_video_id not in _available_ids
                ):
                    st.session_state.v72_ops_applied_video_id = _available_ids[0]

                _default_key = next(
                    k for k, v in _video_option_map.items()
                    if v.get("video_id") == st.session_state.v72_ops_applied_video_id
                )

                with st.form("v72_ops_video_open_form"):
                    _selected_video_key_input = st.selectbox(
                        "기록할 영상",
                        _video_options,
                        index=_video_options.index(_default_key),
                        format_func=lambda x: x.rsplit(" [", 1)[0],
                        key="v72_ops_video_select_input",
                    )
                    _open_video_submit = st.form_submit_button(
                        "📂 선택한 영상 열기",
                        type="primary",
                        use_container_width=True,
                    )

                if _open_video_submit:
                    st.session_state.v72_ops_applied_video_id = (
                        _video_option_map[_selected_video_key_input].get("video_id")
                    )

                _selected_video = next(
                    v for v in _ops_filtered
                    if v.get("video_id") == st.session_state.v72_ops_applied_video_id
                )
                _selected_video_id = _selected_video.get("video_id")

                _note_fetch = _supabase_table_get(
                    "video_notes",
                    {
                        "select": "*",
                        "channel_id": f"eq.{_ops_channel_id}",
                        "video_id": f"eq.{_selected_video_id}",
                        "limit": 1,
                    },
                )
                _note = (
                    _note_fetch.get("rows", [])[0]
                    if _note_fetch.get("ok") and _note_fetch.get("rows")
                    else {}
                )

                _tag_options = [
                    "정보형",
                    "실험형",
                    "비교형",
                    "과정형",
                    "반전형",
                    "문제해결형",
                    "스토리형",
                    "기타",
                ]
                _saved_tags = _note.get("tags") or []
                if not isinstance(_saved_tags, list):
                    _saved_tags = []

                with st.form(f"v67_video_note_form_{_selected_video_id}"):
                    _topic = st.text_input(
                        "주제",
                        value=str(_note.get("topic") or ""),
                        placeholder="예: 타이어 제작 / 과학 원리 / 생활 기술",
                    )
                    _tags = st.multiselect(
                        "영상 태그",
                        options=_tag_options,
                        default=[x for x in _saved_tags if x in _tag_options],
                    )

                    _n1, _n2 = st.columns(2)
                    with _n1:
                        _first_line = st.text_input(
                            "첫 문장",
                            value=str(_note.get("first_line") or ""),
                        )
                        _first_scene = st.text_input(
                            "첫 장면",
                            value=str(_note.get("first_scene") or ""),
                        )
                    with _n2:
                        _hook_type = st.text_input(
                            "훅 유형",
                            value=str(_note.get("hook_type") or ""),
                            placeholder="예: 질문 / 충격 장면 / 결과 먼저",
                        )
                        _script_structure = st.text_input(
                            "구조",
                            value=str(_note.get("script_structure") or ""),
                            placeholder="예: 훅 → 원리 → 결과",
                        )

                    _user_comment = st.text_area(
                        "내 코멘트",
                        value=str(_note.get("user_comment") or ""),
                        placeholder="이 영상에 대해 기억해둘 자유 메모",
                    )
                    _what_worked = st.text_area(
                        "잘됐다고 생각한 점",
                        value=str(_note.get("what_worked_user") or ""),
                        placeholder="사용자가 직접 판단해서 기록",
                    )
                    _what_failed = st.text_area(
                        "아쉬운 점 / 개선점",
                        value=str(_note.get("what_failed_user") or ""),
                        placeholder="예: 초반 설명이 길었음 / 저장 유도가 약했음",
                    )
                    _next_use = st.text_area(
                        "다음에 반복하거나 바꿀 것",
                        value=str(_note.get("next_use_user") or ""),
                    )

                    _save_note = st.form_submit_button(
                        "💾 영상 기록 저장",
                        use_container_width=True,
                    )

                if _save_note:
                    _result = _supabase_table_upsert(
                        "video_notes",
                        {
                            "channel_id": _ops_channel_id,
                            "user_id": None,
                            "video_id": _selected_video_id,
                            "title": _selected_video.get("title"),
                            "topic": _topic.strip(),
                            "tags": _tags,
                            "first_line": _first_line.strip(),
                            "first_scene": _first_scene.strip(),
                            "hook_type": _hook_type.strip(),
                            "script_structure": _script_structure.strip(),
                            "user_comment": _user_comment.strip(),
                            "what_worked_user": _what_worked.strip(),
                            "what_failed_user": _what_failed.strip(),
                            "next_use_user": _next_use.strip(),
                            "source": "user",
                            "updated_at": datetime.now(timezone.utc).isoformat(),
                        },
                        "channel_id,video_id",
                    )
                    if _result.get("ok"):
                        st.success("영상 기록을 저장했습니다.")
                    else:
                        st.error(
                            "영상 기록 저장에 실패했습니다. 운영 테이블이 아직 만들어지지 않았을 수 있습니다."
                        )
                        if _result.get("message"):
                            with st.expander("기술 오류 상세보기"):
                                st.code(_result["message"])

                st.divider()
                st.markdown("#### 📦 쇼마스터 전달용 요약")
                _va = video_analytics.get(_selected_video_id, {})
                _export_text = (
                    f"[Shorts Scope 영상 기록]\\n"
                    f"제목: {_selected_video.get('title', '')}\\n"
                    f"조회수: {int(_selected_video.get('views', 0) or 0):,}\\n"
                    f"평균 시청률: {float(_va.get('average_view_percentage', 0) or 0):.1f}%\\n"
                    f"주제: {_topic}\\n"
                    f"태그: {', '.join(_tags)}\\n"
                    f"첫 문장: {_first_line}\\n"
                    f"첫 장면: {_first_scene}\\n"
                    f"훅 유형: {_hook_type}\\n"
                    f"구조: {_script_structure}\\n"
                    f"내 코멘트: {_user_comment}\\n"
                    f"잘된 점(사용자 기록): {_what_worked}\\n"
                    f"아쉬운 점(사용자 기록): {_what_failed}\\n"
                    f"다음에 반복/변경: {_next_use}\\n"
                )
                st.code(_export_text, language=None)
                st.caption(
                    "현재는 수동 복사용입니다. GPT API 연결 단계에서 쇼마스터로 자동 전달할 수 있게 확장합니다."
                )

        # =====================================================
        # 실험 기록
        # =====================================================
        with _tab_experiment:
            st.subheader("실험 기록")
            st.caption(
                "무엇을 바꿨는지 먼저 기록하고, 결과는 나중에 실제 데이터가 나온 뒤 작성합니다."
            )

            with st.form("v67_experiment_add"):
                _exp_title = st.text_input(
                    "실험 이름",
                    placeholder="예: 첫 1초에 결과 장면 먼저 보여주기",
                )
                _exp_hypothesis = st.text_area(
                    "가설",
                    placeholder="예: 첫 장면에서 결과를 먼저 보여주면 초반 이탈이 줄어들 것이다.",
                )
                _exp_change = st.text_area(
                    "실제로 바꿀 것",
                    placeholder="예: 첫 1초 완성 장면 → 2초부터 제작 과정",
                )
                _exp_start = st.date_input(
                    "시작일",
                    value=today,
                    key="v67_exp_start",
                )
                _add_exp = st.form_submit_button(
                    "➕ 실험 추가",
                    use_container_width=True,
                )

            if _add_exp:
                if not _exp_title.strip():
                    st.warning("실험 이름을 입력해주세요.")
                else:
                    _result = _supabase_table_insert(
                        "experiments",
                        {
                            "channel_id": _ops_channel_id,
                            "user_id": None,
                            "title": _exp_title.strip(),
                            "hypothesis": _exp_hypothesis.strip(),
                            "change_made": _exp_change.strip(),
                            "status": "진행 중",
                            "start_date": _exp_start.isoformat(),
                            "result_note": "",
                            "created_at": datetime.now(timezone.utc).isoformat(),
                            "updated_at": datetime.now(timezone.utc).isoformat(),
                        },
                    )
                    if _result.get("ok"):
                        st.success("실험을 추가했습니다.")
                    else:
                        st.error(
                            "실험 저장에 실패했습니다. 운영 테이블이 아직 만들어지지 않았을 수 있습니다."
                        )

            _exp_fetch = _supabase_table_get(
                "experiments",
                {
                    "select": "*",
                    "channel_id": f"eq.{_ops_channel_id}",
                    "order": "created_at.desc",
                    "limit": 100,
                },
            )
            _experiments = (
                _exp_fetch.get("rows", [])
                if _exp_fetch.get("ok")
                else []
            )

            if _experiments:
                st.markdown("#### 실험 목록")
                for _exp in _experiments:
                    _exp_id = _exp.get("id")
                    _status = _exp.get("status") or "진행 중"
                    with st.expander(
                        f"{_status} · {_exp.get('title', '이름 없음')}"
                    ):
                        st.write(f"**가설:** {_exp.get('hypothesis') or '-'}")
                        st.write(f"**변경 내용:** {_exp.get('change_made') or '-'}")
                        st.write(f"**시작일:** {_exp.get('start_date') or '-'}")

                        _new_status = st.selectbox(
                            "상태",
                            ["진행 중", "완료", "보류"],
                            index=(
                                ["진행 중", "완료", "보류"].index(_status)
                                if _status in ["진행 중", "완료", "보류"]
                                else 0
                            ),
                            key=f"v67_exp_status_{_exp_id}",
                        )
                        _result_note = st.text_area(
                            "결과 메모",
                            value=str(_exp.get("result_note") or ""),
                            key=f"v67_exp_result_{_exp_id}",
                        )

                        _ec1, _ec2 = st.columns(2)
                        if _ec1.button(
                            "💾 수정 저장",
                            key=f"v67_exp_save_{_exp_id}",
                            use_container_width=True,
                        ):
                            _result = _supabase_table_update(
                                "experiments",
                                {
                                    "id": f"eq.{_exp_id}",
                                    "channel_id": f"eq.{_ops_channel_id}",
                                },
                                {
                                    "status": _new_status,
                                    "result_note": _result_note.strip(),
                                    "updated_at": datetime.now(timezone.utc).isoformat(),
                                },
                            )
                            if _result.get("ok"):
                                st.success("실험 기록을 수정했습니다.")
                                st.rerun()
                            else:
                                st.error("실험 기록 수정에 실패했습니다.")

                        if _ec2.button(
                            "🗑️ 삭제",
                            key=f"v67_exp_delete_{_exp_id}",
                            use_container_width=True,
                        ):
                            _result = _supabase_table_delete(
                                "experiments",
                                {
                                    "id": f"eq.{_exp_id}",
                                    "channel_id": f"eq.{_ops_channel_id}",
                                },
                            )
                            if _result.get("ok"):
                                st.success("실험 기록을 삭제했습니다.")
                                st.rerun()
                            else:
                                st.error("실험 기록 삭제에 실패했습니다.")
            else:
                st.caption("아직 저장된 실험이 없습니다.")

    st.divider()

    # =====================================================
    # V7.7 운영 기록 전체 Excel Export
    # =====================================================
    st.markdown("### 운영 기록 내보내기")
    st.caption("목표 · 영상 메모/태그 · 실험 기록을 한 파일에 시트별로 저장합니다.")

    _export_goal = _supabase_table_get(
        "channel_goals",
        {
            "select": "*",
            "channel_id": f"eq.{_ops_channel_id}",
            "order": "updated_at.desc",
            "limit": 1000,
        },
    )
    _export_notes = _supabase_table_get(
        "video_notes",
        {
            "select": "*",
            "channel_id": f"eq.{_ops_channel_id}",
            "order": "updated_at.desc",
            "limit": 10000,
        },
    )
    _export_experiments = _supabase_table_get(
        "experiments",
        {
            "select": "*",
            "channel_id": f"eq.{_ops_channel_id}",
            "order": "created_at.desc",
            "limit": 10000,
        },
    )

    def _ops_excel_bytes():
        _buf = BytesIO()

        _goal_rows = (
            _export_goal.get("rows", [])
            if _export_goal.get("ok") else []
        )
        _note_rows = (
            _export_notes.get("rows", [])
            if _export_notes.get("ok") else []
        )
        _exp_rows = (
            _export_experiments.get("rows", [])
            if _export_experiments.get("ok") else []
        )

        def _excel_kst_datetime(value, date_only=False):
            """Supabase/YouTube UTC ISO 값을 사용자용 KST 문자열로 변환."""
            if not value:
                return ""
            try:
                _dt = datetime.fromisoformat(
                    str(value).replace("Z", "+00:00")
                )
                if _dt.tzinfo is None:
                    _dt = _dt.replace(tzinfo=timezone.utc)
                _dt = _dt.astimezone(KST)
                return (
                    _dt.strftime("%Y-%m-%d")
                    if date_only
                    else _dt.strftime("%Y-%m-%d %H:%M")
                )
            except Exception:
                return str(value)

        def _excel_date(value):
            if not value:
                return ""
            try:
                return datetime.fromisoformat(str(value)).strftime("%Y-%m-%d")
            except Exception:
                return str(value)

        # -------------------------------------------------
        # 1) 목표 — 실제 Shorts Scope 입력 항목과 1:1 대응
        # -------------------------------------------------
        _goal_columns = [
            "주간 업로드 목표",
            "영상 1개 목표 조회수",
            "목표 메모",
            "마지막 수정",
        ]

        _goal_export_rows = []
        for _row in _goal_rows:
            _goal_export_rows.append({
                "주간 업로드 목표": int(_row.get("weekly_upload_goal") or 0),
                "영상 1개 목표 조회수": int(_row.get("target_views") or 0),
                "목표 메모": _row.get("goal_note") or "",
                "마지막 수정": _excel_kst_datetime(_row.get("updated_at")),
            })

        _goal_df = pd.DataFrame(
            _goal_export_rows,
            columns=_goal_columns,
        )

        # -------------------------------------------------
        # 2) 영상 기록 — 사람이 읽는 필드 우선, 내부 ID는 맨 뒤
        # -------------------------------------------------
        _note_columns = [
            "제목",
            "업로드일",
            "현재 조회수",
            "주제",
            "태그",
            "첫 문장",
            "첫 장면",
            "훅 유형",
            "구조",
            "내 코멘트",
            "잘된 점",
            "아쉬운 점 / 개선점",
            "다음에 반복하거나 바꿀 것",
            "마지막 수정",
            "영상 ID",
        ]

        _video_lookup_export = {
            v.get("video_id"): v
            for v in public_videos
        }

        _note_export_rows = []
        for _row in _note_rows:
            _vid = _row.get("video_id")
            _video = _video_lookup_export.get(_vid, {})

            _tags = _row.get("tags") or []
            if isinstance(_tags, list):
                _tags = ", ".join(str(x) for x in _tags)
            else:
                _tags = str(_tags or "")

            _note_export_rows.append({
                "제목": _row.get("title") or _video.get("title") or "",
                "업로드일": _excel_kst_datetime(
                    _video.get("published_raw"),
                    date_only=True,
                ),
                "현재 조회수": int(_video.get("views", 0) or 0),
                "주제": _row.get("topic") or "",
                "태그": _tags,
                "첫 문장": _row.get("first_line") or "",
                "첫 장면": _row.get("first_scene") or "",
                "훅 유형": _row.get("hook_type") or "",
                "구조": _row.get("script_structure") or "",
                "내 코멘트": _row.get("user_comment") or "",
                "잘된 점": _row.get("what_worked_user") or "",
                "아쉬운 점 / 개선점": _row.get("what_failed_user") or "",
                "다음에 반복하거나 바꿀 것": _row.get("next_use_user") or "",
                "마지막 수정": _excel_kst_datetime(_row.get("updated_at")),
                "영상 ID": _vid or "",
            })

        _note_df = pd.DataFrame(
            _note_export_rows,
            columns=_note_columns,
        )

        # -------------------------------------------------
        # 3) 실험 기록 — 운영 판단에 필요한 순서로 정리
        # -------------------------------------------------
        _experiment_columns = [
            "실험명",
            "가설",
            "변경 내용",
            "시작일",
            "상태",
            "결과 메모",
            "생성일",
            "마지막 수정",
            "실험 ID",
        ]

        _exp_export_rows = []
        for _row in _exp_rows:
            _exp_export_rows.append({
                "실험명": _row.get("title") or "",
                "가설": _row.get("hypothesis") or "",
                "변경 내용": _row.get("change_made") or "",
                "시작일": _excel_date(_row.get("start_date")),
                "상태": _row.get("status") or "",
                "결과 메모": _row.get("result_note") or "",
                "생성일": _excel_kst_datetime(_row.get("created_at")),
                "마지막 수정": _excel_kst_datetime(_row.get("updated_at")),
                "실험 ID": _row.get("id") or "",
            })

        _exp_df = pd.DataFrame(
            _exp_export_rows,
            columns=_experiment_columns,
        )

        # -------------------------------------------------
        # 4) 요약
        # -------------------------------------------------
        _summary = pd.DataFrame([
            {"항목": "채널", "값": channel_info.get("channel_name", "")},
            {"항목": "주간 업로드 목표", "값": (
                int(_goal_rows[0].get("weekly_upload_goal") or 0)
                if _goal_rows else ""
            )},
            {"항목": "영상 1개 목표 조회수", "값": (
                int(_goal_rows[0].get("target_views") or 0)
                if _goal_rows else ""
            )},
            {"항목": "영상 기록 수", "값": len(_note_df)},
            {"항목": "실험 기록 수", "값": len(_exp_df)},
            {"항목": "파일 생성 시각", "값": datetime.now(KST).strftime("%Y-%m-%d %H:%M KST")},
        ])

        with pd.ExcelWriter(_buf, engine="openpyxl") as _writer:
            _summary.to_excel(_writer, sheet_name="요약", index=False)
            _goal_df.to_excel(_writer, sheet_name="목표", index=False)
            _note_df.to_excel(_writer, sheet_name="영상 기록", index=False)
            _exp_df.to_excel(_writer, sheet_name="실험 기록", index=False)

            _column_widths = {
                "요약": {
                    "A": 24,
                    "B": 34,
                },
                "목표": {
                    "A": 20,
                    "B": 24,
                    "C": 54,
                    "D": 20,
                },
                "영상 기록": {
                    "A": 44,
                    "B": 14,
                    "C": 14,
                    "D": 22,
                    "E": 24,
                    "F": 38,
                    "G": 34,
                    "H": 18,
                    "I": 28,
                    "J": 42,
                    "K": 42,
                    "L": 42,
                    "M": 42,
                    "N": 20,
                    "O": 18,
                },
                "실험 기록": {
                    "A": 32,
                    "B": 44,
                    "C": 44,
                    "D": 14,
                    "E": 14,
                    "F": 42,
                    "G": 20,
                    "H": 20,
                    "I": 12,
                },
            }

            for _ws in _writer.book.worksheets:
                _ws.freeze_panes = "A2"

                # 헤더
                for _cell in _ws[1]:
                    _cell.font = _cell.font.copy(bold=True)
                    _cell.alignment = _cell.alignment.copy(
                        vertical="center",
                    )
                _ws.row_dimensions[1].height = 24

                # 열 너비
                for _col, _width in _column_widths.get(_ws.title, {}).items():
                    _ws.column_dimensions[_col].width = _width

                # 데이터 셀
                for _row in _ws.iter_rows(min_row=2):
                    for _cell in _row:
                        _cell.alignment = _cell.alignment.copy(
                            vertical="top",
                            wrap_text=True,
                        )

                # 데이터가 있을 때만 필터
                if _ws.max_row >= 2:
                    _ws.auto_filter.ref = _ws.dimensions

        _buf.seek(0)
        return _buf.getvalue()

    st.download_button(
        "운영 기록 전체 Excel",
        data=_ops_excel_bytes(),
        file_name=f"shorts_scope_operations_{today.strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )


if page == "📋 리포트":
    _report_today = today

    if "v77_report_mode" not in st.session_state:
        st.session_state.v77_report_mode = "최근 7일"
    if "v77_report_custom_start" not in st.session_state:
        st.session_state.v77_report_custom_start = _report_today - timedelta(days=7)
    if "v77_report_custom_end" not in st.session_state:
        st.session_state.v77_report_custom_end = _report_today - timedelta(days=1)

    with st.form("v77_report_period_form"):
        _rp0, _rp1, _rp2 = st.columns([1.2, 1, 1])
        with _rp0:
            _report_mode_input = st.selectbox(
                "리포트 기간",
                ["최근 7일", "최근 14일", "최근 30일", "최근 90일", "직접 선택"],
                index=["최근 7일", "최근 14일", "최근 30일", "최근 90일", "직접 선택"].index(
                    st.session_state.v77_report_mode
                ),
            )
        with _rp1:
            _report_custom_start_input = st.date_input(
                "시작일",
                value=st.session_state.v77_report_custom_start,
                disabled=_report_mode_input != "직접 선택",
            )
        with _rp2:
            _report_custom_end_input = st.date_input(
                "종료일",
                value=st.session_state.v77_report_custom_end,
                max_value=_report_today - timedelta(days=1),
                disabled=_report_mode_input != "직접 선택",
            )

        _report_apply = st.form_submit_button(
            "기간 적용",
            use_container_width=True,
        )

    if _report_apply:
        if (
            _report_mode_input == "직접 선택"
            and _report_custom_start_input > _report_custom_end_input
        ):
            st.error("시작일은 종료일보다 늦을 수 없습니다.")
        else:
            st.session_state.v77_report_mode = _report_mode_input
            st.session_state.v77_report_custom_start = _report_custom_start_input
            st.session_state.v77_report_custom_end = _report_custom_end_input

    _report_mode = st.session_state.v77_report_mode

    if _report_mode == "직접 선택":
        _report_start = st.session_state.v77_report_custom_start
        _report_end = min(
            st.session_state.v77_report_custom_end,
            _report_today - timedelta(days=1),
        )
    else:
        _report_days = int(re.sub(r"\D", "", _report_mode))
        _report_end = _report_today - timedelta(days=1)
        _report_start = _report_end - timedelta(days=_report_days - 1)

    _report_span_days = (_report_end - _report_start).days + 1
    _prev_end = _report_start - timedelta(days=1)
    _prev_start = _prev_end - timedelta(days=_report_span_days - 1)

    def _report_upload_count(start_d, end_d):
        _count = 0
        for _v in public_videos:
            _raw = _v.get("published_raw")
            if not _raw:
                continue
            try:
                _dt = datetime.fromisoformat(
                    _raw.replace("Z", "+00:00")
                ).astimezone(KST)
                if start_d <= _dt.date() <= end_d:
                    _count += 1
            except Exception:
                pass
        return _count

    def _safe_pct_change(cur, prev):
        if prev == 0:
            return None
        return ((cur - prev) / abs(prev)) * 100

    # =====================================================
    # 주간 리포트
    # =====================================================
    st.markdown(f"### 선택 기간 · {_report_span_days}일")
    st.caption(
        f"{_report_start} ~ {_report_end} ↔ 이전 동일 길이 {_prev_start} ~ {_prev_end} · 오늘 제외"
    )

    try:
        _week_now = get_period_summary(
            yt_analytics,
            _report_start,
            _report_end,
        )
        _week_prev = get_period_summary(
            yt_analytics,
            _prev_start,
            _prev_end,
        )
        _week_error = None
    except Exception as _exc:
        _week_now = None
        _week_prev = None
        _week_error = str(_exc)

    if _week_now and _week_prev:
        _now_uploads = _report_upload_count(_report_start, _report_end)
        _prev_uploads = _report_upload_count(_prev_start, _prev_end)

        _r1, _r2, _r3, _r4 = st.columns(4)
        _r1.metric(
            "조회수",
            f"{_week_now['views']:,}회",
            (
                f"{_safe_pct_change(_week_now['views'], _week_prev['views']):+.1f}%"
                if _safe_pct_change(_week_now['views'], _week_prev['views']) is not None
                else "비교 불가"
            ),
        )
        _r2.metric(
            "시청시간",
            format_watch_time(_week_now["watch_minutes"]),
            (
                f"{_safe_pct_change(_week_now['watch_minutes'], _week_prev['watch_minutes']):+.1f}%"
                if _safe_pct_change(_week_now['watch_minutes'], _week_prev['watch_minutes']) is not None
                else "비교 불가"
            ),
        )
        _r3.metric(
            "순구독자",
            f"{_week_now['net_subscribers']:+,}명",
            f"{_week_now['net_subscribers'] - _week_prev['net_subscribers']:+,}명",
        )
        _r4.metric(
            "업로드",
            f"{_now_uploads}개",
            f"{_now_uploads - _prev_uploads:+d}개",
        )

        _week_rows = pd.DataFrame([
            {
                "항목": "조회수",
                f"현재 {_report_span_days}일": _week_now["views"],
                f"이전 {_report_span_days}일": _week_prev["views"],
            },
            {
                "항목": "시청시간(분)",
                f"현재 {_report_span_days}일": round(_week_now["watch_minutes"], 1),
                f"이전 {_report_span_days}일": round(_week_prev["watch_minutes"], 1),
            },
            {
                "항목": "순구독자",
                f"현재 {_report_span_days}일": _week_now["net_subscribers"],
                f"이전 {_report_span_days}일": _week_prev["net_subscribers"],
            },
        ])
        with st.expander("현재 기간 ↔ 이전 기간 상세 비교", expanded=False):
            st.dataframe(_week_rows, hide_index=True, use_container_width=True)

        if _prev_uploads == 0:
            st.caption(
                "※ 이전 비교 기간의 업로드가 0개라 조회수 변화만으로 영상 성과 개선을 단정하지 않습니다."
            )
    else:
        st.info("⏳ 선택한 기간의 리포트 데이터를 아직 불러오지 못했습니다.")
        if _week_error:
            with st.expander("기술 오류 상세보기"):
                st.code(_week_error)

    # =====================================================
    # 최근 7일 업로드 영상
    # =====================================================
    _recent_uploaded = []
    for _v in public_videos:
        _raw = _v.get("published_raw")
        if not _raw:
            continue
        try:
            _dt = datetime.fromisoformat(
                _raw.replace("Z", "+00:00")
            ).astimezone(KST)
            if _report_start <= _dt.date() <= _report_end:
                _recent_uploaded.append((_dt, _v))
        except Exception:
            pass

    _recent_uploaded.sort(
        key=lambda x: int(x[1].get("views", 0) or 0),
        reverse=True,
    )

    st.markdown(f"### 선택 기간 업로드 성과 · {_report_span_days}일")
    if _recent_uploaded:
        _recent_views = [
            int(v.get("views", 0) or 0)
            for _, v in _recent_uploaded
        ]
        _recent_median = float(pd.Series(_recent_views).median())

        _rv1, _rv2, _rv3 = st.columns(3)
        _rv1.metric("업로드 수", f"{len(_recent_uploaded)}개")
        _rv2.metric(
            "영상 중앙 조회수",
            f"{_recent_median:,.0f}회",
        )
        _rv3.metric(
            "최고 조회수",
            f"{max(_recent_views):,}회",
        )

        _top3_cards = []
        for _idx, (_dt, _v) in enumerate(_recent_uploaded[:3], start=1):
            _vv = int(_v.get("views", 0) or 0)
            _vs_median = (_vv / _recent_median) if _recent_median > 0 else None
            _ratio_text = f"중앙값 대비 ×{_vs_median:.1f}" if _vs_median is not None else "중앙값 비교 불가"
            _title = str(_v.get("title", "제목 없음")).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            _top3_cards.append(
                (
                    f'<div class="v76-rank-card">'
                    f'<div class="v76-rank-no">TOP {_idx}</div>'
                    f'<div class="v76-rank-title">{_title}</div>'
                    f'<div class="v76-rank-value">{_vv:,}회</div>'
                    f'<div class="v76-rank-meta">{_ratio_text} · {_dt.strftime("%m.%d")}</div>'
                    f'</div>'
                )
            )

        st.markdown("#### 선택 기간 상위 영상")
        st.markdown('<div class="v76-rank-grid">' + "".join(_top3_cards) + '</div>', unsafe_allow_html=True)

        with st.expander("전체 업로드 성과 표 보기", expanded=False):
            _recent_table = []
            for _idx, (_dt, _v) in enumerate(_recent_uploaded, start=1):
                _recent_table.append({
                    "순위": _idx,
                    "업로드": _dt.strftime("%m.%d"),
                    "제목": _v.get("title", ""),
                    "조회수": int(_v.get("views", 0) or 0),
                    "좋아요": int(_v.get("likes", 0) or 0),
                    "댓글": int(_v.get("comments", 0) or 0),
                })
            st.dataframe(pd.DataFrame(_recent_table), hide_index=True, use_container_width=True)
    else:
        st.caption("최근 7일 업로드 영상이 없습니다.")

    st.divider()

    # =====================================================
    # 이상 변화 감지
    # =====================================================
    st.markdown("### 이상 변화 감지")
    st.caption(
        "스냅샷에서 확인되는 급상승·재상승·강한 하락만 표시합니다. "
        "원인 설명은 하지 않습니다."
    )

    try:
        _report_channel_id = _connected_youtube_channel_id(youtube)
        _report_since = (
            datetime.now(timezone.utc) - timedelta(days=8)
        ).isoformat()
        _report_snap_fetch = _fetch_channel_snapshots(
            _report_channel_id,
            _report_since,
        )
    except Exception:
        _report_snap_fetch = {
            "ok": False,
            "reason": "snapshot_error",
            "rows": [],
        }

    _anomalies = []
    if _report_snap_fetch.get("ok"):
        _report_by_video = _group_snapshot_rows(
            _report_snap_fetch.get("rows", []),
            [v.get("video_id") for v in public_videos],
        )
        _now_utc = datetime.now(timezone.utc)

        for _v in public_videos:
            _rows = _report_by_video.get(_v.get("video_id"), [])
            if not _rows:
                continue

            _state = _snapshot_growth_state(
                _v,
                _rows,
                _now_utc,
            )
            _event = _snapshot_special_event(
                _v,
                _state,
                _rows,
                _now_utc,
            )

            if _event.get("label"):
                _anomalies.append({
                    "priority": 3 if _event.get("event") == "surge" else 2,
                    "상태": _event.get("label"),
                    "제목": _v.get("title", ""),
                    "최근 증가": int(_state.get("recent_gain") or 0),
                    "직전 증가": int(_state.get("previous_gain") or 0),
                    "신뢰도": _event.get("confidence"),
                    "근거": _event.get("reason") or "",
                })
            elif (
                _state
                and _state.get("state") == "↘ 하락"
                and _state.get("confidence") in ("보통", "높음")
                and int(_state.get("recent_gain") or 0) > 0
            ):
                _anomalies.append({
                    "priority": 1,
                    "상태": "↘ 강한 하락",
                    "제목": _v.get("title", ""),
                    "최근 증가": int(_state.get("recent_gain") or 0),
                    "직전 증가": int(_state.get("previous_gain") or 0),
                    "신뢰도": _state.get("confidence"),
                    "근거": (
                        f"최근 {_state.get('window_hours')}시간 "
                        f"+{int(_state.get('recent_gain') or 0):,}회 · "
                        f"직전 +{int(_state.get('previous_gain') or 0):,}회"
                    ),
                })

    if _anomalies:
        _anomalies.sort(
            key=lambda x: (x["priority"], x["최근 증가"]),
            reverse=True,
        )
        _anomaly_df = pd.DataFrame(_anomalies[:20])[
            ["상태", "제목", "최근 증가", "직전 증가", "신뢰도", "근거"]
        ]
        st.dataframe(
            _anomaly_df,
            hide_index=True,
            use_container_width=True,
        )
    else:
        st.caption(
            "현재 표시할 급상승·재상승·강한 하락이 없습니다. "
            "스냅샷 데이터가 더 쌓이면 자동으로 감지합니다."
        )

    st.divider()


    # =====================================================
    # V7 스냅샷 데이터 상태
    # =====================================================
    st.markdown("### 스냅샷 데이터 상태")
    st.caption(
        "자체 데이터 품질을 확인합니다. 최근 수집이 끊기면 성장 분석의 신뢰도가 떨어질 수 있습니다."
    )

    try:
        _health_channel_id = _connected_youtube_channel_id(youtube)
        _health_since = (
            datetime.now(timezone.utc) - timedelta(hours=12)
        ).isoformat()
        _health_fetch = _fetch_channel_snapshots(
            _health_channel_id,
            _health_since,
        )
    except Exception:
        _health_fetch = {
            "ok": False,
            "reason": "snapshot_error",
            "rows": [],
        }

    if _health_fetch.get("ok"):
        _health_rows = _health_fetch.get("rows", [])
        _health_times = []
        _health_video_ids = set()

        for _row in _health_rows:
            try:
                _dt = _parse_utc(_row.get("captured_at"))
                if _dt:
                    _health_times.append(_dt)
                if _row.get("video_id"):
                    _health_video_ids.add(_row.get("video_id"))
            except Exception:
                pass

        if _health_times:
            _latest_snapshot = max(_health_times)
            _age_minutes = (
                datetime.now(timezone.utc) - _latest_snapshot
            ).total_seconds() / 60

            _hc1, _hc2, _hc3 = st.columns(3)
            _hc1.metric(
                "마지막 스냅샷",
                _latest_snapshot.astimezone(KST).strftime("%m.%d %H:%M"),
            )
            _hc2.metric(
                "최근 12시간 스냅샷 행",
                f"{len(_health_rows):,}개",
            )
            _hc3.metric(
                "최근 수집 영상",
                f"{len(_health_video_ids)}개",
            )

            if _age_minutes <= 45:
                st.success("✅ 스냅샷 자동수집 정상")
            elif _age_minutes <= 90:
                st.warning("⚠️ 최근 스냅샷이 평소보다 늦습니다.")
            else:
                st.error(
                    "🚨 스냅샷 수집이 지연되고 있습니다. Edge Function / Cron 상태 확인이 필요합니다."
                )
        else:
            st.warning("최근 12시간 스냅샷이 없습니다.")
    else:
        st.warning("스냅샷 상태를 확인하지 못했습니다.")

    st.divider()

    # =====================================================
    # 개인 최고기록
    # =====================================================
    st.markdown("### 개인 최고기록")
    st.caption(
        "현재 분석 가능한 공개 영상 기준입니다. "
        "YouTube 공식 기록이 아니라 Shorts Scope 내부 비교입니다."
    )

    if public_videos:
        _best_views = max(
            public_videos,
            key=lambda v: int(v.get("views", 0) or 0),
        )
        _best_likes = max(
            public_videos,
            key=lambda v: int(v.get("likes", 0) or 0),
        )
        _best_comments = max(
            public_videos,
            key=lambda v: int(v.get("comments", 0) or 0),
        )

        _analytics_candidates = [
            v for v in public_videos
            if v.get("video_id") in video_analytics
        ]

        _record_rows = [
            {
                "기록": "최고 조회수",
                "영상": _best_views.get("title", ""),
                "값": f"{int(_best_views.get('views', 0) or 0):,}회",
            },
            {
                "기록": "최다 좋아요",
                "영상": _best_likes.get("title", ""),
                "값": f"{int(_best_likes.get('likes', 0) or 0):,}개",
            },
            {
                "기록": "최다 댓글",
                "영상": _best_comments.get("title", ""),
                "값": f"{int(_best_comments.get('comments', 0) or 0):,}개",
            },
        ]

        if _analytics_candidates:
            _best_ret = max(
                _analytics_candidates,
                key=lambda v: float(v.get("avg_percentage", 0) or 0),
            )
            _best_sub = max(
                _analytics_candidates,
                key=lambda v: float(v.get("sub_conversion_rate", 0) or 0),
            )
            _record_rows.extend([
                {
                    "기록": "최고 평균 시청률",
                    "영상": _best_ret.get("title", ""),
                    "값": f"{float(_best_ret.get('avg_percentage', 0) or 0):.1f}%",
                },
                {
                    "기록": "최고 구독전환율",
                    "영상": _best_sub.get("title", ""),
                    "값": f"{float(_best_sub.get('sub_conversion_rate', 0) or 0):.3f}%",
                },
            ])

        st.dataframe(
            pd.DataFrame(_record_rows),
            hide_index=True,
            use_container_width=True,
        )

    st.divider()


if page == "📈 성장 분석":
    st.markdown(
        """
        <div class="v78-flow">
            <span><strong>1</strong> 기준 시점 확인</span>
            <span><strong>2</strong> 채널 추세 확인</span>
            <span><strong>3</strong> 영상 검색</span>
            <span><strong>4</strong> 상세 분석</span>
        </div>
        """,
        unsafe_allow_html=True,
    )
    # =========================================================
    # 20. 자동 성과 리포트 V6
    # =========================================================

    st.caption(
        f"데이터 조회 {datetime.now(KST).strftime('%m.%d %H:%M')} · 최근 Analytics는 집계가 지연될 수 있습니다."
    )

    report_videos = [
        v for v in public_videos
        if v.get("views", 0) >= 100 and v.get("video_id") in video_analytics
    ]
    st.caption(
        f"분석 대상: 공개 영상 {len(public_videos)}개 중 {len(report_videos)}개 · "
        f"마지막 조회: {datetime.now(KST).strftime('%Y-%m-%d %H:%M KST')}"
    )

    # V6.5-1: 현재 성장상태용 Supabase 스냅샷
    _snapshot_fetch = {"ok": False, "rows": [], "reason": "not_started", "message": None}
    _snapshot_by_video = {}
    _snapshot_channel_id = None
    try:
        _snapshot_channel_id = _connected_youtube_channel_id(youtube)
        _snapshot_since = (
            datetime.now(timezone.utc) - timedelta(days=4)
        ).replace(minute=0, second=0, microsecond=0).isoformat()
        _snapshot_fetch = _fetch_channel_snapshots(_snapshot_channel_id, _snapshot_since)
        if _snapshot_fetch.get("ok"):
            _snapshot_by_video = _group_snapshot_rows(
                _snapshot_fetch.get("rows", []),
                [v.get("video_id") for v in report_videos],
            )
    except Exception:
        _snapshot_fetch = {
            "ok": False,
            "rows": [],
            "reason": "unexpected_error",
            "message": "스냅샷 성장 데이터를 준비하지 못했습니다.",
        }

    if _snapshot_fetch.get("ok"):
        _snapshot_video_count = len(_snapshot_by_video)
        _snapshot_row_count = len(_snapshot_fetch.get("rows", []))

        if _snapshot_video_count > 0:
            st.caption(
                f"⚡ 스냅샷 성장 데이터 연결 · 현재 분석 가능한 영상 {_snapshot_video_count}개"
            )
        elif _snapshot_row_count == 0:
            st.caption(
                "⏳ 최근 4일 스냅샷 데이터가 없어 성장상태 분석을 기다리는 중입니다. "
                "수집이 다시 시작되면 자동으로 표시됩니다."
            )
        else:
            st.caption(
                "⏳ 최근 스냅샷은 확인됐지만 현재 성장 분석 대상 영상과 매칭되는 데이터가 없습니다. "
                "데이터가 더 쌓이면 자동으로 표시됩니다."
            )

        if _snapshot_fetch.get("truncated"):
            st.warning("스냅샷 조회량이 많아 일부 최신 데이터만 사용 중입니다.")
    elif _snapshot_fetch.get("reason") == "not_configured":
        st.caption("⏳ 스냅샷 성장상태 연결 준비 중 · 기존 Analytics 분석은 그대로 사용할 수 있습니다.")
    else:
        st.caption("⚠️ 스냅샷 성장상태를 불러오지 못했습니다. 기존 Analytics 분석은 그대로 사용할 수 있습니다.")


    # ---------------------------------------------------------
    # V7 — 자체 스냅샷 기준 시점 데이터
    # ---------------------------------------------------------
    st.markdown("### 스냅샷 기준 시점")
    st.caption(
        "업로드 후 1h / 6h / 12h / 24h / 48h / 72h 시점의 실제 스냅샷을 사용합니다. "
        "기준 시점 근처 데이터가 없으면 0으로 채우지 않고 '데이터 부족'으로 표시합니다."
    )

    try:
        _v7_channel_id = _connected_youtube_channel_id(youtube)
        _v7_since = (
            datetime.now(timezone.utc) - timedelta(days=10)
        ).isoformat()
        _v7_fetch = _fetch_channel_snapshots(
            _v7_channel_id,
            _v7_since,
        )
    except Exception:
        _v7_fetch = {
            "ok": False,
            "reason": "snapshot_error",
            "rows": [],
        }

    if _v7_fetch.get("ok"):
        _v7_grouped = _group_snapshot_rows(
            _v7_fetch.get("rows", []),
            [v.get("video_id") for v in public_videos],
        )

        _v7_recent_videos = []
        for _v in public_videos:
            _raw = _v.get("published_raw")
            if not _raw:
                continue
            try:
                _pdt = datetime.fromisoformat(
                    _raw.replace("Z", "+00:00")
                ).astimezone(KST)
                _v7_recent_videos.append((_pdt, _v))
            except Exception:
                pass

        _v7_recent_videos.sort(key=lambda x: x[0], reverse=True)
        _v7_recent_videos = _v7_recent_videos[:20]

        if _v7_recent_videos:
            _v7_lookup = {}
            _v7_options = []
            for _pdt, _v in _v7_recent_videos:
                _label = (
                    f"{_pdt.strftime('%Y.%m.%d')} | "
                    f"{_v.get('title', '제목 없음')} | "
                    f"{int(_v.get('views', 0) or 0):,}회"
                )
                _key = f"{_label} [{_v.get('video_id')}]"
                _v7_lookup[_key] = _v
                _v7_options.append(_key)

            _v7_available_ids = [v.get("video_id") for _, v in _v7_recent_videos]
            if (
                "v72_v7_applied_video_id" not in st.session_state
                or st.session_state.v72_v7_applied_video_id not in _v7_available_ids
            ):
                st.session_state.v72_v7_applied_video_id = _v7_available_ids[0]

            _v7_default_key = next(
                k for k, v in _v7_lookup.items()
                if v.get("video_id") == st.session_state.v72_v7_applied_video_id
            )

            if "v77_v7_search_text" not in st.session_state:
                st.session_state.v77_v7_search_text = ""
            if "v77_v7_period" not in st.session_state:
                st.session_state.v77_v7_period = "최근 90일"
            if "v77_v7_sort" not in st.session_state:
                st.session_state.v77_v7_sort = "최신순"

            with st.form("v77_v7_video_search_form"):
                _v7_search_input = st.text_input(
                    "영상 제목 검색",
                    value=st.session_state.v77_v7_search_text,
                    placeholder="제목 일부를 입력하세요",
                )
                _vf1, _vf2 = st.columns(2)
                with _vf1:
                    _v7_period_input = st.selectbox(
                        "기간",
                        ["최근 30일", "최근 90일", "최근 1년", "전체"],
                        index=["최근 30일", "최근 90일", "최근 1년", "전체"].index(
                            st.session_state.v77_v7_period
                        ),
                    )
                with _vf2:
                    _v7_sort_input = st.selectbox(
                        "정렬",
                        ["최신순", "조회수순", "오래된순"],
                        index=["최신순", "조회수순", "오래된순"].index(
                            st.session_state.v77_v7_sort
                        ),
                    )
                _v7_filter_submit = st.form_submit_button(
                    "영상 찾기",
                    use_container_width=True,
                )

            if _v7_filter_submit:
                st.session_state.v77_v7_search_text = _v7_search_input.strip()
                st.session_state.v77_v7_period = _v7_period_input
                st.session_state.v77_v7_sort = _v7_sort_input

            _v7_filtered = list(_v7_recent_videos)
            _needle = st.session_state.v77_v7_search_text.lower().strip()

            if _needle:
                _v7_filtered = [
                    pair for pair in _v7_filtered
                    if _needle in str(pair[1].get("title", "")).lower()
                ]

            _days_map = {"최근 30일": 30, "최근 90일": 90, "최근 1년": 365}
            _selected_period = st.session_state.v77_v7_period
            if _selected_period in _days_map:
                _cutoff = datetime.now(KST) - timedelta(days=_days_map[_selected_period])
                _v7_filtered = [
                    pair for pair in _v7_filtered
                    if pair[0] >= _cutoff
                ]

            if st.session_state.v77_v7_sort == "조회수순":
                _v7_filtered.sort(
                    key=lambda pair: int(pair[1].get("views", 0) or 0),
                    reverse=True,
                )
            elif st.session_state.v77_v7_sort == "오래된순":
                _v7_filtered.sort(key=lambda pair: pair[0])
            else:
                _v7_filtered.sort(key=lambda pair: pair[0], reverse=True)

            _v7_filtered = _v7_filtered[:100]

            if _v7_filtered:
                _v7_result_lookup = {}
                _v7_result_options = []
                for _pdt, _v in _v7_filtered:
                    _label = (
                        f"{_pdt.strftime('%Y.%m.%d')} | "
                        f"{_v.get('title', '제목 없음')} | "
                        f"{int(_v.get('views', 0) or 0):,}회"
                    )
                    _key = f"{_label} [{_v.get('video_id')}]"
                    _v7_result_lookup[_key] = _v
                    _v7_result_options.append(_key)

                _current_key = next(
                    (
                        k for k, v in _v7_result_lookup.items()
                        if v.get("video_id") == st.session_state.v72_v7_applied_video_id
                    ),
                    _v7_result_options[0],
                )

                with st.form("v77_v7_video_pick_form"):
                    _v7_selected_input = st.selectbox(
                        f"검색 결과에서 선택 · 최대 100개 표시",
                        _v7_result_options,
                        index=_v7_result_options.index(_current_key),
                        format_func=lambda x: x.rsplit(" [", 1)[0],
                    )
                    _v7_lookup_submit = st.form_submit_button(
                        "기준 시점 데이터 보기",
                        use_container_width=True,
                    )

                if _v7_lookup_submit:
                    st.session_state.v72_v7_applied_video_id = (
                        _v7_result_lookup[_v7_selected_input].get("video_id")
                    )
            else:
                st.info("검색 조건에 맞는 영상이 없습니다.")

            _v7_video = next(
                v for _, v in _v7_recent_videos
                if v.get("video_id") == st.session_state.v72_v7_applied_video_id
            )
            _v7_rows = _v7_grouped.get(_v7_video.get("video_id"), [])

            if _v7_rows:
                _v7_table = pd.DataFrame(
                    _snapshot_milestone_table(
                        _v7_video,
                        _v7_rows,
                        datetime.now(timezone.utc),
                    )
                )
                st.dataframe(
                    _v7_table,
                    hide_index=True,
                    use_container_width=True,
                )

                _g_0_6 = _snapshot_interval_gain(
                    _v7_video, _v7_rows, 0, 6, datetime.now(timezone.utc)
                )
                _g_6_24 = _snapshot_interval_gain(
                    _v7_video, _v7_rows, 6, 24, datetime.now(timezone.utc)
                )
                _g_24_72 = _snapshot_interval_gain(
                    _v7_video, _v7_rows, 24, 72, datetime.now(timezone.utc)
                )

                _g1, _g2, _g3 = st.columns(3)
                _g1.metric(
                    "0~6시간 증가",
                    (
                        f"+{_g_0_6['gain']:,}회"
                        if _g_0_6 and _g_0_6.get("gain") is not None
                        else "데이터 부족"
                    ),
                )
                _g2.metric(
                    "6~24시간 증가",
                    (
                        f"+{_g_6_24['gain']:,}회"
                        if _g_6_24 and _g_6_24.get("gain") is not None
                        else "데이터 부족"
                    ),
                )
                _g3.metric(
                    "24~72시간 증가",
                    (
                        f"+{_g_24_72['gain']:,}회"
                        if _g_24_72 and _g_24_72.get("gain") is not None
                        else "데이터 부족"
                    ),
                )

                st.caption(
                    "※ 조회수 조정으로 누적값이 줄어든 구간은 증가량을 억지로 계산하지 않습니다."
                )
            else:
                st.caption("⏳ 선택한 영상의 스냅샷이 아직 충분하지 않습니다.")
        else:
            st.caption("최근 공개 영상이 없습니다.")
    else:
        st.caption("⚠️ V7 스냅샷 기준 시점 데이터를 불러오지 못했습니다.")

    st.divider()

    # -----------------------------
    # 채널 추세
    # -----------------------------
    st.markdown("### 채널 추세")

    trend_last_day = today - timedelta(days=1)

    if "v72_trend_applied_option" not in st.session_state:
        st.session_state.v72_trend_applied_option = "최근 7일"
        st.session_state.v72_trend_applied_start = trend_last_day - timedelta(days=6)
        st.session_state.v72_trend_applied_end = trend_last_day

    _trend_options = ["최근 7일", "최근 14일", "최근 28일", "직접 선택"]

    with st.form("v72_trend_form"):
        trend_option_input = st.selectbox(
            "추세 분석 기간",
            _trend_options,
            index=_trend_options.index(st.session_state.v72_trend_applied_option),
            key="v72_trend_period_input",
        )
        _tc1, _tc2 = st.columns(2)
        with _tc1:
            trend_start_input = st.date_input(
                "직접 선택 시작일",
                value=st.session_state.v72_trend_applied_start,
                max_value=trend_last_day,
                disabled=trend_option_input != "직접 선택",
                key="v72_trend_start_input",
            )
        with _tc2:
            trend_end_input = st.date_input(
                "직접 선택 종료일",
                value=st.session_state.v72_trend_applied_end,
                max_value=trend_last_day,
                disabled=trend_option_input != "직접 선택",
                key="v72_trend_end_input",
            )
        _trend_submit = st.form_submit_button(
            "🔍 추세 조회",
            type="primary",
            use_container_width=True,
        )

    if _trend_submit:
        if trend_option_input == "최근 7일":
            _new_end = trend_last_day
            _new_start = _new_end - timedelta(days=6)
        elif trend_option_input == "최근 14일":
            _new_end = trend_last_day
            _new_start = _new_end - timedelta(days=13)
        elif trend_option_input == "최근 28일":
            _new_end = trend_last_day
            _new_start = _new_end - timedelta(days=27)
        else:
            _new_start = trend_start_input
            _new_end = trend_end_input

        if _new_start > _new_end:
            _new_start = _new_end

        st.session_state.v72_trend_applied_option = trend_option_input
        st.session_state.v72_trend_applied_start = _new_start
        st.session_state.v72_trend_applied_end = _new_end

    trend_option = st.session_state.v72_trend_applied_option
    trend_start = st.session_state.v72_trend_applied_start
    trend_end = st.session_state.v72_trend_applied_end
    st.caption(f"현재 적용: {trend_start} ~ {trend_end}")

    trend_days = (trend_end - trend_start).days + 1
    trend_prev_end = trend_start - timedelta(days=1)
    trend_prev_start = trend_prev_end - timedelta(days=trend_days - 1)

    def _count_uploads(start_d, end_d):
        count = 0
        for _v in public_videos:
            _raw = _v.get("published_raw")
            if not _raw:
                continue
            try:
                _dt = datetime.fromisoformat(_raw.replace("Z", "+00:00")).astimezone(KST)
                if start_d <= _dt.date() <= end_d:
                    count += 1
            except Exception:
                pass
        return count

    def _trend_change(cur, prev):
        if prev == 0:
            return "계산 불가" if cur != 0 else "0%"
        return f"{((cur-prev)/abs(prev))*100:+.1f}%"

    # 그래프 토글 등 Streamlit 재실행 때 같은 기간의 Analytics를 반복 호출하지 않도록
    # session_state에 결과를 보관합니다. YouTube Analytics 500/backendError는
    # 일시적일 수 있어 각 호출을 최대 2회 시도합니다.
    def _trend_api_call_with_retry(callable_fn, *args):
        last_exc = None
        for attempt in range(2):
            try:
                return callable_fn(*args)
            except Exception as exc:
                last_exc = exc
                err_text = str(exc)
                is_backend = (
                    "backendError" in err_text
                    or "Internal error encountered" in err_text
                    or "HttpError 500" in err_text
                )
                if not is_backend or attempt == 1:
                    raise
                time.sleep(1)
        raise last_exc

    _trend_cache_key = (
        str(_snapshot_channel_id or "unknown"),
        str(trend_start), str(trend_end),
        str(trend_prev_start), str(trend_prev_end),
    )
    _trend_cache = st.session_state.get("growth_trend_cache_v652")

    if _trend_cache and _trend_cache.get("key") == _trend_cache_key:
        trend_now = _trend_cache.get("trend_now")
        trend_prev = _trend_cache.get("trend_prev")
        trend_daily_rows = _trend_cache.get("trend_daily_rows", [])
        trend_error = None
    else:
        try:
            trend_now = _trend_api_call_with_retry(
                get_period_summary, yt_analytics, trend_start, trend_end
            )
            trend_prev = _trend_api_call_with_retry(
                get_period_summary, yt_analytics, trend_prev_start, trend_prev_end
            )
            trend_daily_rows = _trend_api_call_with_retry(
                get_daily_channel_data, yt_analytics, trend_start, trend_end
            )
            trend_error = None
            st.session_state["growth_trend_cache_v652"] = {
                "key": _trend_cache_key,
                "trend_now": trend_now,
                "trend_prev": trend_prev,
                "trend_daily_rows": trend_daily_rows,
            }
        except Exception as exc:
            trend_now = None
            trend_prev = None
            trend_daily_rows = []
            trend_error = str(exc)

    if trend_error:
        if (
            "backendError" in trend_error
            or "Internal error encountered" in trend_error
            or "HttpError 500" in trend_error
        ):
            st.warning(
                "⚠️ YouTube Analytics 서버가 일시적으로 응답하지 않습니다. "
                "자동으로 다시 시도했지만 아직 실패했습니다. 잠시 후 다시 확인해 주세요."
            )
        else:
            st.warning("채널 추세 데이터를 일부 불러오지 못했습니다.")
        with st.expander("기술 오류 상세보기"):
            st.code(trend_error)

    if trend_now and trend_prev:
        now_uploads = _count_uploads(trend_start, trend_end)
        prev_uploads = _count_uploads(trend_prev_start, trend_prev_end)

        st.caption(
            f"현재 {trend_start} ~ {trend_end} ↔ 이전 {trend_prev_start} ~ {trend_prev_end} · 오늘 제외"
        )

        _views_change_text = (
            "비교 주의"
            if prev_uploads == 0
            else _trend_change(trend_now["views"], trend_prev["views"])
        )

        trend_table = pd.DataFrame([
            ["조회수", f"{trend_now['views']:,}회", f"{trend_prev['views']:,}회",
             _views_change_text],
            ["시청시간", f"{trend_now['watch_minutes']/60:,.1f}시간",
             f"{trend_prev['watch_minutes']/60:,.1f}시간",
             _trend_change(trend_now["watch_minutes"], trend_prev["watch_minutes"])],
            ["순구독자", f"{trend_now['net_subscribers']:+,}명",
             f"{trend_prev['net_subscribers']:+,}명",
             f"{trend_now['net_subscribers']-trend_prev['net_subscribers']:+,}명"],
            ["업로드", f"{now_uploads:,}개", f"{prev_uploads:,}개",
             f"{now_uploads-prev_uploads:+,}개"],
        ], columns=["지표", "현재 기간", "이전 기간", "변화"])

        st.dataframe(trend_table, hide_index=True, use_container_width=True)

        if prev_uploads == 0:
            st.warning(
                "⚠️ 이전 기간 업로드가 0개입니다. 조회수 증가율이 커 보여도 "
                "콘텐츠 자체의 성과가 같은 비율로 개선됐다고 볼 수는 없습니다."
            )
        elif trend_prev["views"] == 0:
            st.warning(
                "⚠️ 이전 기간 조회수가 0회라 변화율 비교가 의미 없습니다."
            )

        if now_uploads > 0 and prev_uploads > 0:
            now_per_upload = trend_now["views"] / now_uploads
            prev_per_upload = trend_prev["views"] / prev_uploads
            tc1, tc2 = st.columns(2)
            tc1.metric("조회수 ÷ 업로드 수 (참고)", f"{now_per_upload:,.0f}회")
            tc2.metric("이전 기간", f"{prev_per_upload:,.0f}회")
            st.caption(
                "※ 이 값에는 기존 영상 조회수도 포함됩니다. 신규 영상 1편의 실제 평균 조회수는 아닙니다."
            )
        else:
            st.caption("※ 두 기간 모두 업로드가 있을 때만 '조회수 ÷ 업로드 수'를 표시합니다.")

        if trend_daily_rows:
            trend_df = pd.DataFrame(trend_daily_rows)
            if not trend_df.empty and "date" in trend_df.columns and "views" in trend_df.columns:
                trend_df["date"] = pd.to_datetime(trend_df["date"])
                trend_df = trend_df.set_index("date")
                st.markdown("#### 일별 조회수 흐름")
                st.line_chart(trend_df[["views"]], use_container_width=True, height=240)

    st.divider()

    if report_videos:
        base_views = sum(v["views"] for v in report_videos) / len(report_videos)
        median_views = float(pd.Series([v["views"] for v in report_videos]).median())
        base_ret = sum(v["avg_percentage"] for v in report_videos) / len(report_videos)
        base_like = sum(v["like_rate"] for v in report_videos) / len(report_videos)
        base_sub = sum(v["sub_conversion_rate"] for v in report_videos) / len(report_videos)

        with st.expander("내 채널 기준선 보기", expanded=False):
            b1,b2,b3=st.columns(3)
            b1.metric("분석 대상 평균 조회수", f"{base_views:,.0f}회")
            b2.metric("분석 대상 중앙 조회수", f"{median_views:,.0f}회")
            b3.metric("평균 시청률", f"{base_ret:.1f}%")

            b4,b5=st.columns(2)
            b4.metric("평균 좋아요율", f"{base_like:.2f}%")
            b5.metric("평균 구독전환율", f"{base_sub:.3f}%")
            st.caption("※ 현재 분석 가능한 내 영상들의 비교값이며 YouTube 공식 기준이 아닙니다.")

        st.markdown("### 영상별 성과 비교")
        st.caption(
            "제목 검색 · 기간 · 성장상태 · 정렬로 영상을 찾고, 선택한 영상 1개만 상세 분석합니다."
        )
        st.caption(
            "※ 이 화면은 실제 수치 비교만 표시합니다. 원인 판단과 개선 메모는 운영(V6.7)에서 사용자가 직접 기록합니다."
        )

        # 실제 영상별 일별 Analytics는 한 번만 가져와 아래 모든 성장 비교에서 재사용합니다.
        try:
            _report_ids = [v["video_id"] for v in report_videos]
            _published_dates_pt = []
            _PT = ZoneInfo("America/Los_Angeles")
            for _v in report_videos:
                _raw = _v.get("published_raw")
                if not _raw:
                    continue
                try:
                    _pdt = datetime.fromisoformat(_raw.replace("Z", "+00:00")).astimezone(_PT)
                    _published_dates_pt.append(_pdt.date())
                except Exception:
                    pass

            _growth_start = min(_published_dates_pt) if _published_dates_pt else today - timedelta(days=90)
            _growth_end = today - timedelta(days=1)

            daily_video_growth = (
                get_daily_video_analytics(
                    yt_analytics,
                    _report_ids,
                    _growth_start,
                    _growth_end,
                )
                if _growth_start <= _growth_end
                else {}
            )
            growth_error = None
        except Exception as _growth_exc:
            daily_video_growth = {}
            growth_error = str(_growth_exc)

        def _published_pt_date(video):
            raw = video.get("published_raw")
            if not raw:
                return None
            try:
                return datetime.fromisoformat(raw.replace("Z", "+00:00")).astimezone(
                    ZoneInfo("America/Los_Angeles")
                ).date()
            except Exception:
                return None

        def _growth_series(video):
            pub_date = _published_pt_date(video)
            if not pub_date:
                return []

            rows = daily_video_growth.get(video.get("video_id"), [])
            by_date = {str(r.get("date")): int(r.get("views", 0)) for r in rows}

            # Analytics 날짜 기준 D0부터 어제까지. 없는 날짜는 0회로 채움.
            last_day = today - timedelta(days=1)
            if pub_date > last_day:
                return []

            out = []
            cumulative = 0
            d = pub_date
            age = 0
            while d <= last_day:
                views = by_date.get(d.isoformat(), 0)
                cumulative += views
                out.append({
                    "D": age,
                    "date": d,
                    "daily_views": views,
                    "cumulative_views": cumulative,
                })
                d += timedelta(days=1)
                age += 1
            return out

        growth_cache = {
            v["video_id"]: _growth_series(v)
            for v in report_videos
        }

        def _cum_at(video, d_index):
            series = growth_cache.get(video.get("video_id"), [])
            if len(series) <= d_index:
                return None
            return series[d_index]["cumulative_views"]

        def _comparison_for(video):
            series = growth_cache.get(video.get("video_id"), [])
            if not series:
                return None

            max_d = series[-1]["D"]
            milestone = 3 if max_d >= 3 else (1 if max_d >= 1 else 0)
            current_value = _cum_at(video, milestone)
            if current_value is None:
                return None

            peers = []
            for other in report_videos:
                val = _cum_at(other, milestone)
                if val is not None:
                    peers.append((other.get("video_id"), val))

            if not peers:
                return None

            values = [x[1] for x in peers]
            median = float(pd.Series(values).median())
            rank = 1 + sum(1 for _, val in peers if val > current_value)

            sample = len(peers)
            top_percent = (rank / sample) * 100 if sample else None

            return {
                "milestone": milestone,
                "value": current_value,
                "median": median,
                "rank": rank,
                "sample": sample,
                "top_percent": top_percent,
            }

        def _daily_analytics_trend(video):
            series = growth_cache.get(video.get("video_id"), [])
            if len(series) < 4:
                return "⏳ 데이터 축적 중"

            daily = [x["daily_views"] for x in series]
            recent = daily[-2:]
            previous = daily[-4:-2]

            recent_avg = sum(recent) / len(recent)
            previous_avg = sum(previous) / len(previous)

            # 오래된 영상이 다시 살아나는 경우
            if previous_avg > 0 and recent_avg >= previous_avg * 2 and recent_avg >= 100:
                return "🔥 재상승"
            if previous_avg == 0:
                return "↗ 성장 중" if recent_avg > 0 else "→ 정체"
            ratio = recent_avg / previous_avg
            if ratio >= 1.25:
                return "↗ 성장 중"
            if ratio <= 0.60:
                return "↘ 둔화"
            return "→ 유지"

        ordered=sorted(report_videos,key=lambda v:v.get("views",0),reverse=True)

        def pct_diff(value, base):
            return "비교 불가" if base == 0 else f"{((value-base)/abs(base))*100:+.1f}%"

        def pp_diff(value, base, digits):
            diff = value - base
            threshold = {
                1: 1.0,      # 시청률: ±1.0%p 이내는 비슷
                2: 0.05,     # 좋아요율: ±0.05%p 이내는 비슷
                3: 0.005,    # 구독전환율: ±0.005%p 이내는 비슷
            }.get(digits, 0)
            if abs(diff) <= threshold:
                return "≈ 비슷"
            return f"{diff:+.{digits}f}%p"

        def compare_level(value, base, metric_name):
            if metric_name == "조회수":
                if base == 0:
                    return "similar"
                diff_ratio = abs(value - base) / abs(base)
                if diff_ratio <= 0.05:
                    return "similar"
            elif metric_name == "시청률":
                if abs(value - base) <= 1.0:
                    return "similar"
            elif metric_name == "좋아요율":
                if abs(value - base) <= 0.05:
                    return "similar"
            elif metric_name == "구독전환율":
                if abs(value - base) <= 0.005:
                    return "similar"

            return "high" if value > base else "low"

        # ---------------------------------------------------------
        # V6.5 통합 — 다중 성장곡선 비교
        # ---------------------------------------------------------
        with st.expander("📉 여러 영상 성장곡선 비교", expanded=False):
            st.caption(
                "YouTube Analytics의 D+N 누적 조회수를 같은 축에서 비교합니다. "
                "최대 5개 영상까지 선택할 수 있습니다."
            )

            _curve_options = []
            _curve_lookup = {}
            for _idx, _video in enumerate(ordered, start=1):
                _label = f"{_idx}. {_video['title'][:55]}"
                _curve_options.append(_label)
                _curve_lookup[_label] = _video

            _default_curves = _curve_options[: min(3, len(_curve_options))]
            if "v72_curve_applied_videos" not in st.session_state:
                st.session_state.v72_curve_applied_videos = _default_curves
                st.session_state.v72_curve_applied_days = 28

            _valid_saved_curves = [
                x for x in st.session_state.v72_curve_applied_videos
                if x in _curve_options
            ]
            if not _valid_saved_curves:
                _valid_saved_curves = _default_curves
                st.session_state.v72_curve_applied_videos = _valid_saved_curves

            if "v77_curve_search_text" not in st.session_state:
                st.session_state.v77_curve_search_text = ""

            with st.form("v77_curve_search_form"):
                _curve_search_input = st.text_input(
                    "비교 영상 검색",
                    value=st.session_state.v77_curve_search_text,
                    placeholder="제목 일부 입력 → 후보를 먼저 줄입니다",
                )
                _curve_search_submit = st.form_submit_button(
                    "후보 검색",
                    use_container_width=True,
                )

            if _curve_search_submit:
                st.session_state.v77_curve_search_text = _curve_search_input.strip()

            _curve_needle = st.session_state.v77_curve_search_text.lower().strip()
            _curve_filtered_options = [
                label for label in _curve_options
                if not _curve_needle or _curve_needle in label.lower()
            ][:100]

            # 이미 선택된 영상은 검색 결과에서 빠져도 유지
            for _saved in _valid_saved_curves:
                if _saved not in _curve_filtered_options:
                    _curve_filtered_options.insert(0, _saved)

            with st.form("v72_multi_curve_form"):
                _selected_curves_input = st.multiselect(
                    "비교할 영상 · 검색 결과 최대 100개",
                    options=_curve_filtered_options,
                    default=[
                        x for x in _valid_saved_curves
                        if x in _curve_filtered_options
                    ],
                    max_selections=5,
                    key="v72_multi_growth_curve_input",
                )
                _curve_days_input = st.selectbox(
                    "비교 구간",
                    [7, 14, 28],
                    index=[7, 14, 28].index(st.session_state.v72_curve_applied_days),
                    format_func=lambda x: f"D+0 ~ D+{x}",
                    key="v72_multi_growth_days_input",
                )
                _curve_submit = st.form_submit_button(
                    "📈 선택한 영상 비교",
                    type="primary",
                    use_container_width=True,
                )

            if _curve_submit:
                st.session_state.v72_curve_applied_videos = _selected_curves_input
                st.session_state.v72_curve_applied_days = _curve_days_input

            _selected_curves = st.session_state.v72_curve_applied_videos
            _curve_days = st.session_state.v72_curve_applied_days

            _multi_rows = {}
            for _label in _selected_curves:
                _video = _curve_lookup[_label]
                _series = growth_cache.get(_video.get("video_id"), [])
                if not _series:
                    continue

                _short_label = _label if len(_label) <= 34 else _label[:31] + "..."
                for _point in _series:
                    _d = int(_point["D"])
                    if _d > _curve_days:
                        break
                    _multi_rows.setdefault(_d, {})[_short_label] = _point["cumulative_views"]

            if _multi_rows:
                _multi_df = pd.DataFrame.from_dict(_multi_rows, orient="index").sort_index()
                _multi_df.index = [f"D+{int(x)}" for x in _multi_df.index]
                _multi_df.index.name = "업로드 후"
                st.line_chart(_multi_df, use_container_width=True, height=320)
                st.caption(
                    "※ 영상마다 집계 가능한 마지막 D+N 시점이 다를 수 있어 "
                    "뒤쪽 구간은 일부 선이 먼저 끝날 수 있습니다."
                )
            else:
                st.caption("⏳ 선택한 영상의 일별 성장 데이터가 아직 충분하지 않습니다.")

        st.divider()

        # ---------------------------------------------------------
        # V6.6.1 — 영상별 성과 비교 검색 / 필터 / 정렬 / 단일 선택
        # 영상이 수백 개가 되어도 전체 expander를 만들지 않고,
        # 원하는 영상 1개만 찾아 상세를 표시합니다.
        # ---------------------------------------------------------
        st.markdown("#### 🔎 비교할 영상 찾기")

        if "v72_compare_applied" not in st.session_state:
            st.session_state.v72_compare_applied = {
                "search": "",
                "period": "전체",
                "state": "전체",
                "sort": "최신순",
            }

        _compare_saved = st.session_state.v72_compare_applied
        _period_options = ["전체", "최근 7일", "최근 30일", "최근 90일"]
        _state_options = [
            "전체",
            "🚀 급상승",
            "🔥 재상승",
            "↗ 상승",
            "→ 유지",
            "↘ 하락",
            "💤 정체",
            "⏳ 데이터 축적 중",
        ]
        _sort_options = ["최신순", "조회수순", "성장속도순", "동일 나이 순위순"]

        with st.form("v72_compare_filter_form"):
            _filter_search_input = st.text_input(
                "영상 제목 검색",
                value=_compare_saved["search"],
                placeholder="제목 일부를 입력하세요",
                key="v72_compare_search_input",
            )
            _fc1, _fc2, _fc3 = st.columns(3)
            with _fc1:
                _filter_period_input = st.selectbox(
                    "기간",
                    _period_options,
                    index=_period_options.index(_compare_saved["period"]),
                    key="v72_compare_period_input",
                )
            with _fc2:
                _filter_state_input = st.selectbox(
                    "성장 상태",
                    _state_options,
                    index=_state_options.index(_compare_saved["state"]),
                    key="v72_compare_state_input",
                )
            with _fc3:
                _filter_sort_input = st.selectbox(
                    "정렬",
                    _sort_options,
                    index=_sort_options.index(_compare_saved["sort"]),
                    key="v72_compare_sort_input",
                )
            _compare_search_submit = st.form_submit_button(
                "🔍 검색",
                type="primary",
                use_container_width=True,
            )

        if _compare_search_submit:
            st.session_state.v72_compare_applied = {
                "search": _filter_search_input.strip().lower(),
                "period": _filter_period_input,
                "state": _filter_state_input,
                "sort": _filter_sort_input,
            }

        _compare_saved = st.session_state.v72_compare_applied
        _filter_search = _compare_saved["search"]
        _filter_period = _compare_saved["period"]
        _filter_state = _compare_saved["state"]
        _filter_sort = _compare_saved["sort"]

        def _compare_filter_meta(_video):
            _raw = _video.get("published_raw")
            _published_dt = None
            if _raw:
                try:
                    _published_dt = datetime.fromisoformat(
                        _raw.replace("Z", "+00:00")
                    ).astimezone(KST)
                except Exception:
                    pass

            _comp_meta = _comparison_for(_video)
            _snap_rows = _snapshot_by_video.get(_video.get("video_id"), [])
            _now = datetime.now(timezone.utc)
            _snap_state = (
                _snapshot_growth_state(_video, _snap_rows, _now)
                if _snapshot_fetch.get("ok")
                else None
            )
            _snap_event = (
                _snapshot_special_event(_video, _snap_state, _snap_rows, _now)
                if _snap_state
                else {"event": None, "label": None}
            )

            _state_label = (
                _snap_state.get("state")
                if _snap_state
                else "⏳ 데이터 축적 중"
            )
            _event_label = _snap_event.get("label")
            _velocity = (
                float(_snap_state.get("recent_velocity") or 0)
                if _snap_state
                else 0.0
            )

            return {
                "video": _video,
                "published_dt": _published_dt,
                "comp": _comp_meta,
                "state": _state_label,
                "event": _event_label,
                "velocity": _velocity,
            }

        _compare_items = [_compare_filter_meta(_v) for _v in ordered]

        _days_limit = {
            "최근 7일": 7,
            "최근 30일": 30,
            "최근 90일": 90,
        }.get(_filter_period)

        _filtered_items = []
        for _item in _compare_items:
            _video = _item["video"]

            if _filter_search and _filter_search not in str(
                _video.get("title", "")
            ).lower():
                continue

            if _days_limit is not None:
                _pdt = _item.get("published_dt")
                if _pdt is None:
                    continue
                _age_days = (datetime.now(KST).date() - _pdt.date()).days
                if _age_days < 0 or _age_days > _days_limit:
                    continue

            if _filter_state != "전체":
                if _filter_state in ("🚀 급상승", "🔥 재상승"):
                    if _item.get("event") != _filter_state:
                        continue
                elif _item.get("state") != _filter_state:
                    continue

            _filtered_items.append(_item)

        if _filter_sort == "최신순":
            _filtered_items.sort(
                key=lambda x: x.get("published_dt") or datetime.min.replace(tzinfo=KST),
                reverse=True,
            )
        elif _filter_sort == "조회수순":
            _filtered_items.sort(
                key=lambda x: int(x["video"].get("views", 0) or 0),
                reverse=True,
            )
        elif _filter_sort == "성장속도순":
            _filtered_items.sort(
                key=lambda x: x.get("velocity", 0),
                reverse=True,
            )
        else:
            _filtered_items.sort(
                key=lambda x: (
                    (x.get("comp") or {}).get("rank")
                    if (x.get("comp") or {}).get("rank") is not None
                    else 10**9
                )
            )

        st.caption(
            f"조건에 맞는 영상 {len(_filtered_items)}개 / 전체 분석 대상 {len(ordered)}개"
        )

        if not _filtered_items:
            st.info("조건에 맞는 영상이 없습니다. 검색어나 필터를 바꿔주세요.")
            _selected_pairs = []
        else:
            _option_lookup = {}
            _option_labels = []
            for _item in _filtered_items:
                _video = _item["video"]
                _pdt = _item.get("published_dt")
                _date_text = _pdt.strftime("%Y.%m.%d") if _pdt else "날짜 없음"
                _status_bits = []
                if _item.get("event"):
                    _status_bits.append(_item["event"])
                if _item.get("state"):
                    _status_bits.append(_item["state"])
                _status_text = " · ".join(_status_bits)

                _label = (
                    f"{_date_text} | {_video.get('title', '제목 없음')} "
                    f"| {int(_video.get('views', 0) or 0):,}회"
                )
                if _status_text:
                    _label += f" | {_status_text}"

                # 제목이 같아도 video_id로 내부 식별
                _key = f"{_label} [{_video.get('video_id')}]"
                _option_lookup[_key] = _item
                _option_labels.append(_key)

            _filtered_ids = [
                _item["video"].get("video_id")
                for _item in _filtered_items
            ]
            if (
                "v72_compare_applied_video_id" not in st.session_state
                or st.session_state.v72_compare_applied_video_id not in _filtered_ids
            ):
                st.session_state.v72_compare_applied_video_id = _filtered_ids[0]

            _detail_default_key = next(
                k for k, item in _option_lookup.items()
                if item["video"].get("video_id") == st.session_state.v72_compare_applied_video_id
            )

            if "v77_detail_search_text" not in st.session_state:
                st.session_state.v77_detail_search_text = ""

            with st.form("v77_detail_search_form"):
                _detail_search_input = st.text_input(
                    "상세 분석 영상 제목 검색",
                    value=st.session_state.v77_detail_search_text,
                    placeholder="제목 일부를 입력하세요",
                )
                _detail_search_submit = st.form_submit_button(
                    "후보 검색",
                    use_container_width=True,
                )

            if _detail_search_submit:
                st.session_state.v77_detail_search_text = _detail_search_input.strip()

            _detail_needle = st.session_state.v77_detail_search_text.lower().strip()
            _detail_options = [
                key for key in _option_labels
                if (
                    not _detail_needle
                    or _detail_needle in key.rsplit(" [", 1)[0].lower()
                )
            ][:100]

            if _detail_default_key not in _detail_options:
                _detail_options.insert(0, _detail_default_key)

            with st.form("v72_compare_detail_form"):
                _selected_key_input = st.selectbox(
                    "상세 분석할 영상 · 검색 결과 최대 100개",
                    options=_detail_options,
                    index=_detail_options.index(_detail_default_key),
                    format_func=lambda x: x.rsplit(" [", 1)[0],
                    key="v72_compare_selected_video_input",
                )
                _detail_submit = st.form_submit_button(
                    "선택한 영상 분석 보기",
                    use_container_width=True,
                )

            if _detail_submit:
                st.session_state.v72_compare_applied_video_id = (
                    _option_lookup[_selected_key_input]["video"].get("video_id")
                )

            _selected_item = next(
                item for item in _filtered_items
                if item["video"].get("video_id") == st.session_state.v72_compare_applied_video_id
            )
            _selected_video = _selected_item["video"]
            _original_rank = next(
                (
                    _idx
                    for _idx, _ov in enumerate(ordered, start=1)
                    if _ov.get("video_id") == _selected_video.get("video_id")
                ),
                1,
            )
            _selected_pairs = [(_original_rank, _selected_video)]

        for rank,v in _selected_pairs:
            published_text = "업로드일 확인 불가"
            age_text = ""
            raw = v.get("published_raw")
            if raw:
                try:
                    published_dt = datetime.fromisoformat(raw.replace("Z", "+00:00")).astimezone(KST)
                    published_date = published_dt.date()
                    age_days = max((today - published_date).days, 0)
                    published_text = published_date.strftime("%Y.%m.%d")
                    age_text = f" · 업로드 후 {age_days}일"
                except Exception:
                    pass

            _comp = _comparison_for(v)
            _snapshot_rows_for_video = _snapshot_by_video.get(v.get("video_id"), [])
            _snapshot_now = datetime.now(timezone.utc)
            _snapshot_state = _snapshot_growth_state(
                v,
                _snapshot_rows_for_video,
                _snapshot_now,
            ) if _snapshot_fetch.get("ok") else None
            _snapshot_event = _snapshot_special_event(
                v,
                _snapshot_state,
                _snapshot_rows_for_video,
                _snapshot_now,
            ) if _snapshot_state else {"event": None, "label": None, "confidence": "낮음", "reason": None}

            _daily_trend = _daily_analytics_trend(v)
            _state = _snapshot_state["state"] if _snapshot_state else f"📅 {_daily_trend}"
            _event_text = f" | {_snapshot_event['label']}" if _snapshot_event.get("label") else ""

            if _comp and _comp["sample"] >= 5:
                _top_pct = _comp.get("top_percent")
                _top_text = f" · 상위 {_top_pct:.1f}%" if _top_pct is not None else ""
                _same_age_text = (
                    f"D+{_comp['milestone']} {_comp['rank']}위/{_comp['sample']}개{_top_text}"
                )
            elif _comp:
                _same_age_text = f"⏳ 비교 데이터 부족 · {_comp['sample']}개"
            else:
                _same_age_text = "⏳ 성장 데이터 집계 중"

            summary = (
                f"📅 {published_text}{age_text} | "
                f"조회수 {v['views']:,} | {_same_age_text} | {_state}{_event_text}"
            )

            with st.expander(f"{rank}. {v['title']}  |  {summary}"):
                rows=[
                    ["조회수",f"{v['views']:,}회",f"{base_views:,.0f}회",pct_diff(v["views"],base_views)],
                    ["평균 시청률",f"{v['avg_percentage']:.1f}%",f"{base_ret:.1f}%",pp_diff(v["avg_percentage"],base_ret,1)],
                    ["좋아요율",f"{v['like_rate']:.2f}%",f"{base_like:.2f}%",pp_diff(v["like_rate"],base_like,2)],
                    ["구독전환율",f"{v['sub_conversion_rate']:.3f}%",f"{base_sub:.3f}%",pp_diff(v["sub_conversion_rate"],base_sub,3)],
                ]
                st.dataframe(pd.DataFrame(rows,columns=["지표","이 영상","채널 기준선","차이"]),
                             hide_index=True,use_container_width=True)

                high=[]; low=[]; similar=[]
                for name,val,base in [
                    ("조회수",v["views"],base_views),("시청률",v["avg_percentage"],base_ret),
                    ("좋아요율",v["like_rate"],base_like),("구독전환율",v["sub_conversion_rate"],base_sub)]:
                    level = compare_level(val, base, name)
                    if level == "high":
                        high.append(name)
                    elif level == "low":
                        low.append(name)
                    else:
                        similar.append(name)

                parts=[]
                if high: parts.append("높음: "+", ".join(high))
                if similar: parts.append("비슷: "+", ".join(similar))
                if low: parts.append("낮음: "+", ".join(low))
                st.caption(
                    "📌 기준선 비교 · " + (" · ".join(parts) if parts else "채널 기준선과 비슷한 수준")
                )

                st.markdown("**⚡ 현재 성장 상태 (스냅샷)**")
                if _snapshot_state:
                    st.write(f"**{_snapshot_state['state']}**")
                    if _snapshot_state.get("recent_gain") is not None:
                        _w = _snapshot_state["window_hours"]
                        _ratio = _snapshot_state.get("ratio")
                        _ratio_text = (
                            "직전 구간 0회/h"
                            if _ratio is None and (_snapshot_state.get("previous_velocity") or 0) <= 0
                            else (f"속도 {_ratio:.2f}배" if _ratio is not None else "속도 비교 보류")
                        )
                        st.caption(
                            f"최근 {_w}시간 +{_snapshot_state['recent_gain']:,}회 "
                            f"({_snapshot_state['recent_velocity']:.1f}회/h) · "
                            f"직전 {_w}시간 +{_snapshot_state['previous_gain']:,}회 "
                            f"({_snapshot_state['previous_velocity']:.1f}회/h) · "
                            f"{_ratio_text} · 신뢰도 {_snapshot_state['confidence']}"
                        )
                    else:
                        st.caption("비교 가능한 시간대의 스냅샷이 더 쌓이면 자동으로 판정합니다.")
                    if _snapshot_state.get("note"):
                        st.caption(f"※ {_snapshot_state['note']}")

                    if _snapshot_event.get("label"):
                        st.markdown(f"**{_snapshot_event['label']}**")
                        st.caption(
                            f"{_snapshot_event['reason']} · "
                            f"이벤트 신뢰도 {_snapshot_event['confidence']}"
                        )
                    else:
                        st.caption(
                            "특별 이벤트(🚀 급상승 / 🔥 재상승)는 "
                            "충분한 속도 변화와 활동량이 확인될 때만 표시합니다."
                        )
                elif _snapshot_fetch.get("ok"):
                    if len(_snapshot_fetch.get("rows", [])) == 0:
                        st.caption(
                            "⏳ 최근 4일 스냅샷 데이터가 없습니다. "
                            "새 스냅샷이 쌓이면 최근 성장속도를 자동으로 계산합니다."
                        )
                    else:
                        st.caption(
                            "⏳ 이 영상의 비교 가능한 스냅샷 데이터가 아직 부족합니다. "
                            "데이터가 더 쌓이면 자동으로 판정합니다."
                        )
                    st.caption(f"현재 일별 Analytics 참고: {_daily_trend}")
                elif _snapshot_fetch.get("reason") == "not_configured":
                    st.caption("⏳ Supabase 스냅샷 읽기 설정 후 최근 성장속도가 표시됩니다.")
                    st.caption(f"현재 일별 Analytics 참고: {_daily_trend}")
                else:
                    st.caption("⚠️ 스냅샷 성장상태를 불러오지 못했습니다.")
                    st.caption(f"현재 일별 Analytics 참고: {_daily_trend}")

                st.markdown("**D+N 실제 성장 데이터**")
                _series = growth_cache.get(v.get("video_id"), [])

                if growth_error:
                    st.caption("영상별 일별 Analytics를 불러오지 못해 성장 비교는 표시하지 않습니다.")
                elif not _series:
                    st.caption("⏳ 아직 일별 Analytics가 충분히 집계되지 않았습니다.")
                else:
                    _milestone_rows = []
                    for _d in [0, 1, 3, 7, 14, 28]:
                        _value = _cum_at(v, _d)
                        if _value is not None:
                            _milestone_rows.append({
                                "시점": f"D+{_d}",
                                "누적 조회수": f"{_value:,}회",
                            })

                    if _milestone_rows:
                        st.dataframe(
                            pd.DataFrame(_milestone_rows),
                            hide_index=True,
                            use_container_width=True,
                        )

                    if _comp:
                        if _comp["sample"] >= 5:
                            _median_diff = (
                                None if _comp["median"] == 0
                                else ((_comp["value"] - _comp["median"]) / _comp["median"]) * 100
                            )
                            _diff_text = (
                                "비교 불가"
                                if _median_diff is None
                                else f"{_median_diff:+.1f}%"
                            )
                            _top_pct = _comp.get("top_percent")
                            _top_text = (
                                f" · 상위 {_top_pct:.1f}%"
                                if _top_pct is not None
                                else ""
                            )
                            st.write(
                                f"**D+{_comp['milestone']} 동일 시점:** "
                                f"{_comp['rank']}위 / {_comp['sample']}개{_top_text} · "
                                f"채널 중앙값 {_comp['median']:,.0f}회 · "
                                f"중앙값 대비 {_diff_text}"
                            )
                        else:
                            st.caption(
                                f"⏳ D+{_comp['milestone']} 비교 가능 영상이 {_comp['sample']}개라 "
                                "순위 평가는 보류합니다."
                            )

                    _chart_df = pd.DataFrame([
                        {
                            "업로드 후": f"D+{x['D']}",
                            "이 영상 누적 조회수": x["cumulative_views"],
                        }
                        for x in _series[:29]
                    ])
                    if not _chart_df.empty:
                        _chart_df = _chart_df.set_index("업로드 후")
                        _show_growth_chart = st.toggle(
                            "성장 그래프 보기",
                            value=False,
                            key=f"growth_chart_{v['video_id']}",
                        )
                        if _show_growth_chart:
                            st.line_chart(
                                _chart_df,
                                use_container_width=True,
                                height=220,
                            )
                            st.caption(
                                "※ D+N은 YouTube Analytics의 날짜 단위 데이터 기준입니다. "
                                "정확한 업로드 후 N×24시간 값과는 다를 수 있습니다."
                            )
    else:
        st.info("아직 비교 가능한 영상이 없습니다.")

if page == "🔎 영상 찾기":
    # =========================================================
    # 21. 전체 영상 분석
    # =========================================================

    def format_video_upload_kst(video):
        """
        Excel/표에서 날짜가 #### 또는 이상한 숫자로 보이지 않도록
        YouTube 원본 업로드 시각을 한국시간 문자열로 고정합니다.
        """
        published_raw = video.get("published_raw")

        if published_raw:
            try:
                published_dt = datetime.fromisoformat(
                    published_raw.replace("Z", "+00:00")
                ).astimezone(KST)
                return published_dt.strftime("%Y-%m-%d %H:%M")
            except Exception:
                pass

        published_value = video.get("published", "")
        if published_value is None:
            return ""

        return str(published_value)


    table_data = []


    for video in videos:

        has_analytics = (
            video["video_id"]
            in video_analytics
        )

        table_data.append({

            "상태":
                video["status"],

            "제목":
                video["title"],

            "조회수":
                video["views"],

            "좋아요":
                video["likes"],

            "댓글":
                video["comments"],

            "공유":
                video["shares"],

            "좋아요율":
                f"{video['like_rate']:.2f}%",

            "댓글률":
                f"{video['comment_rate']:.2f}%",

            "구독전환율":
                f"{video['sub_conversion_rate']:.3f}%",

            "성과점수":
                (
                    (
                        video.get("performance_score")
                        if video.get("performance_score") is not None
                        else "평가 보류"
                    )
                    if video["status"] == "🟢 공개"
                    else "-"
                ),

            "성과등급":
                (
                    video.get("performance_grade", "-")
                    if video["status"] == "🟢 공개"
                    else "-"
                ),

            "평균 시청":
                (
                    f"{video['avg_duration']:.1f}초"
                    if has_analytics
                    else "-"
                ),

            "평균 시청률":
                (
                    f"{video['avg_percentage']:.1f}%"
                    if has_analytics
                    else "-"
                ),

            "구독자":
                (
                    video["net_subs"]
                    if has_analytics
                    else 0
                ),

            "길이":
                video["duration"],

            "업로드":
                format_video_upload_kst(video),

            "예약 공개":
                video["scheduled"],
        })


    df = pd.DataFrame(
        table_data
    )


    # =========================================================
    # 22-1. 영상 검색 / 조회수 필터
    # =========================================================

    st.caption(
        "조건을 정한 뒤 검색 버튼을 눌러 결과를 확인하세요. "
        "입력값을 바꾸기만 해서는 결과가 변경되지 않습니다."
    )

    if "applied_video_filter" not in st.session_state:
        st.session_state.applied_video_filter = {
            "mode": "전체",
            "views": 5000,
            "min_views": 1000,
            "max_views": 10000,
            "status": "전체",
            "title": "",
            "sort": "조회수 높은 순",
        }

    _saved_filter = st.session_state.applied_video_filter
    _filter_modes = ["전체", "이상", "이하", "범위"]
    _sort_options_all = [
        "조회수 높은 순",
        "조회수 낮은 순",
        "최신 업로드 순",
        "오래된 업로드 순",
        "평균 시청률 높은 순",
        "구독 증가 높은 순",
    ]
    status_options = ["전체"] + sorted(df["상태"].dropna().astype(str).unique().tolist())
    _saved_status = (
        _saved_filter["status"]
        if _saved_filter["status"] in status_options
        else "전체"
    )

    with st.form("v72_video_search_form"):
        filter_mode_input = st.radio(
            "조회수 조건",
            _filter_modes,
            index=_filter_modes.index(_saved_filter["mode"]),
            horizontal=True,
            key="v72_views_filter_mode_input",
        )

        _vc1, _vc2, _vc3 = st.columns(3)
        with _vc1:
            filter_views_input = st.number_input(
                "기준 조회수 (이상/이하)",
                min_value=0,
                value=int(_saved_filter["views"]),
                step=500,
                key="v72_views_filter_value_input",
            )
        with _vc2:
            filter_min_input = st.number_input(
                "최소 조회수 (범위)",
                min_value=0,
                value=int(_saved_filter["min_views"]),
                step=500,
                key="v72_views_filter_min_input",
            )
        with _vc3:
            filter_max_input = st.number_input(
                "최대 조회수 (범위)",
                min_value=0,
                value=int(_saved_filter["max_views"]),
                step=500,
                key="v72_views_filter_max_input",
            )

        _fc1, _fc2 = st.columns(2)
        with _fc1:
            filter_status_input = st.selectbox(
                "공개 상태",
                status_options,
                index=status_options.index(_saved_status),
                key="v72_video_status_filter_input",
            )
        with _fc2:
            title_query_input = st.text_input(
                "제목 검색",
                value=_saved_filter["title"],
                placeholder="예: 비버, 화산, 교통사고",
                key="v72_video_title_filter_input",
            )

        sort_option_input = st.selectbox(
            "정렬",
            _sort_options_all,
            index=_sort_options_all.index(_saved_filter["sort"]),
            key="v72_video_sort_input",
        )

        _video_filter_submit = st.form_submit_button(
            "🔍 검색",
            type="primary",
            use_container_width=True,
        )

    if _video_filter_submit:
        st.session_state.applied_video_filter = {
            "mode": filter_mode_input,
            "views": int(filter_views_input),
            "min_views": int(filter_min_input),
            "max_views": int(filter_max_input),
            "status": filter_status_input,
            "title": title_query_input.strip(),
            "sort": sort_option_input,
        }

    applied_filter = st.session_state.applied_video_filter
    filtered_df = df.copy()

    if applied_filter["mode"] == "이상":
        filtered_df = filtered_df[filtered_df["조회수"] >= applied_filter["views"]]
    elif applied_filter["mode"] == "이하":
        filtered_df = filtered_df[filtered_df["조회수"] <= applied_filter["views"]]
    elif applied_filter["mode"] == "범위":
        low = min(applied_filter["min_views"], applied_filter["max_views"])
        high = max(applied_filter["min_views"], applied_filter["max_views"])
        filtered_df = filtered_df[
            (filtered_df["조회수"] >= low) & (filtered_df["조회수"] <= high)
        ]

    if applied_filter["status"] != "전체":
        filtered_df = filtered_df[filtered_df["상태"] == applied_filter["status"]]

    if applied_filter["title"]:
        filtered_df = filtered_df[
            filtered_df["제목"].astype(str).str.contains(
                applied_filter["title"], case=False, na=False
            )
        ]

    # 검색 결과 정렬
    sort_name = applied_filter["sort"]
    if sort_name == "조회수 높은 순":
        filtered_df = filtered_df.sort_values("조회수", ascending=False)
    elif sort_name == "조회수 낮은 순":
        filtered_df = filtered_df.sort_values("조회수", ascending=True)
    elif sort_name == "평균 시청률 높은 순":
        filtered_df = filtered_df.assign(
            _sort_pct=pd.to_numeric(
                filtered_df["평균 시청률"].astype(str).str.replace("%", "", regex=False),
                errors="coerce",
            )
        ).sort_values("_sort_pct", ascending=False).drop(columns=["_sort_pct"])
    elif sort_name == "구독 증가 높은 순":
        filtered_df = filtered_df.sort_values("구독자", ascending=False)
    elif sort_name in ["최신 업로드 순", "오래된 업로드 순"]:
        filtered_df = filtered_df.assign(
            _sort_date=pd.to_datetime(filtered_df["업로드"], errors="coerce")
        ).sort_values(
            "_sort_date",
            ascending=(sort_name == "오래된 업로드 순"),
        ).drop(columns=["_sort_date"])

    result_count = len(filtered_df)
    result_avg_views = int(filtered_df["조회수"].mean()) if result_count else 0
    result_max_views = int(filtered_df["조회수"].max()) if result_count else 0

    fc1, fc2, fc3 = st.columns(3)
    fc1.metric("검색 영상", f"{result_count:,}개")
    fc2.metric("평균 조회수", f"{result_avg_views:,}회")
    fc3.metric("최고 조회수", f"{result_max_views:,}회")

    st.caption(
        f"현재 적용: 조회수 {applied_filter['mode']} · "
        f"상태 {applied_filter['status']} · 정렬 {applied_filter['sort']}"
    )

    with st.expander(
        f"📋 검색 결과 보기 ({result_count:,}개)",
        expanded=(0 < result_count <= 10),
    ):
        if filtered_df.empty:
            st.info("조건에 맞는 영상이 없습니다.")
        else:
            st.dataframe(
                filtered_df,
                use_container_width=True,
                hide_index=True,
                height=min(420, 90 + (len(filtered_df) * 34)),
            )


    # =========================================================
    # 22-2. 엑셀 다운로드
    # =========================================================

    def make_excel_file(primary_df, primary_sheet_name):
        """
        전체 영상/검색 결과 엑셀.
        날짜·길이를 엑셀이 임의의 날짜/시간 형식으로 바꾸지 않도록 문자열로 내보냅니다.
        """
        output = BytesIO()

        export_df = primary_df.copy()

        # Excel이 업로드 날짜/영상 길이를 자동 날짜·시간으로 오인하지 않게 문자열 고정
        for column_name in ["업로드", "예약 공개", "길이"]:
            if column_name in export_df.columns:
                export_df[column_name] = export_df[column_name].fillna("").astype(str)

        summary_rows = [
            ["채널명", channel_info.get("channel_name", channel_info.get("title", ""))],
            ["구독자", channel_info.get("subscribers", 0)],
            ["채널 총 조회수", channel_info.get("total_views", channel_info.get("views", 0))],
            ["전체 감지 영상", len(videos)],
            ["공개 영상", len(public_videos)],
            ["예약 영상", len(scheduled_videos)],
            ["내보낸 영상", len(export_df)],
            ["생성 시각", datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S KST")],
        ]
        summary_df = pd.DataFrame(summary_rows, columns=["항목", "값"])

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            export_df.to_excel(writer, sheet_name=primary_sheet_name, index=False)
            summary_df.to_excel(writer, sheet_name="채널 요약", index=False)

            for sheet_name in writer.book.sheetnames:
                ws = writer.book[sheet_name]
                ws.freeze_panes = "A2"
                ws.auto_filter.ref = ws.dimensions

                for column_cells in ws.columns:
                    max_length = 0
                    column_letter = column_cells[0].column_letter

                    for cell in column_cells:
                        try:
                            cell_length = len(str(cell.value)) if cell.value is not None else 0
                            max_length = max(max_length, cell_length)
                        except Exception:
                            pass

                    ws.column_dimensions[column_letter].width = min(
                        max(max_length + 3, 11),
                        48,
                    )

                # 영상 데이터 시트에서 자주 잘리던 열은 최소 폭 보장
                if sheet_name == primary_sheet_name:
                    header_map = {
                        cell.value: cell.column_letter
                        for cell in ws[1]
                        if cell.value is not None
                    }

                    if "업로드" in header_map:
                        ws.column_dimensions[header_map["업로드"]].width = 21

                    if "예약 공개" in header_map:
                        ws.column_dimensions[header_map["예약 공개"]].width = 21

                    if "길이" in header_map:
                        ws.column_dimensions[header_map["길이"]].width = 12

                    if "제목" in header_map:
                        ws.column_dimensions[header_map["제목"]].width = 42

                    # 날짜/시간 열은 '텍스트'로 고정
                    for header in ["업로드", "예약 공개", "길이"]:
                        if header in header_map:
                            col = header_map[header]
                            for row in range(2, ws.max_row + 1):
                                ws[f"{col}{row}"].number_format = "@"

            for idx, sheet_name in enumerate(writer.book.sheetnames, start=1):
                add_excel_table(writer.book[sheet_name], f"VideoTable_{idx}")

        output.seek(0)
        return output.getvalue()

    st.markdown("### 📥 엑셀 다운로드")
    download_c1, download_c2 = st.columns(2)

    with download_c1:
        st.download_button(
            "📥 전체 영상 엑셀",
            data=make_excel_file(df, "전체 영상"),
            file_name=f"shorts_all_videos_{today.strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    with download_c2:
        st.download_button(
            f"📥 검색 결과 엑셀 ({result_count}개)",
            data=make_excel_file(filtered_df, "검색 결과") if result_count else b"",
            file_name=f"shorts_filtered_videos_{today.strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            disabled=result_count == 0,
        )

    with st.expander("📋 전체 영상 데이터 보기", expanded=False):
        st.dataframe(
            df,
            use_container_width=True,
            hide_index=True,
            height=340,
        )




    # =========================================================
    # 22. 영상 TOP 랭킹
    # =========================================================

    with st.expander("🏆 공개 영상 TOP 랭킹 보기", expanded=False):
        if "v72_applied_rank_limit" not in st.session_state:
            st.session_state.v72_applied_rank_limit = 5

        with st.form("v72_rank_limit_form"):
            _rank_limit_input = st.radio(
                "표시 개수",
                [5, 10, 20],
                index=[5, 10, 20].index(st.session_state.v72_applied_rank_limit),
                horizontal=True,
                format_func=lambda n: f"TOP {n}",
                key="v72_ranking_display_count_input",
            )
            _rank_limit_submit = st.form_submit_button(
                "🔍 적용",
                use_container_width=True,
            )

        if _rank_limit_submit:
            st.session_state.v72_applied_rank_limit = _rank_limit_input

        rank_limit = st.session_state.v72_applied_rank_limit

        rank_tab1, rank_tab2, rank_tab3, rank_tab4 = st.tabs(
            ["👁️ 조회수", "📊 시청률", "👤 구독전환", "👍 좋아요율"]
        )

        def render_ranked_videos(ranked, metric_name, metric_formatter):
            if not ranked:
                st.info("표시할 공개 영상이 없습니다.")
                return

            for rank, video in enumerate(ranked[:rank_limit], start=1):
                col_img, col_info = st.columns([1, 5])
                with col_img:
                    if video["thumbnail"]:
                        st.image(video["thumbnail"], width=125)
                with col_info:
                    st.markdown(f"**{rank}위 · {video['title']}**")
                    score_value = video.get("performance_score")

                    if score_value is None:
                        score_text = "⏳ 데이터 부족"
                    else:
                        score_text = (
                            f"🚦 {score_value}점 · "
                            f"{video.get('performance_grade', '-')}"
                        )

                    st.write(
                        f"**{metric_name}: {metric_formatter(video)}**  |  "
                        f"{score_text}"
                    )
                    st.write(
                        f"👁️ {video['views']:,}회  |  "
                        f"👍 {video['likes']:,} ({video['like_rate']:.2f}%)  |  "
                        f"💬 {video['comments']:,} ({video['comment_rate']:.2f}%)  |  "
                        f"👤 {video['net_subs']:+d} ({video['sub_conversion_rate']:.3f}%)"
                    )
                    if video["video_id"] in video_analytics:
                        st.write(
                            f"⏱️ 평균 시청 {video['avg_duration']:.1f}초  |  "
                            f"📊 평균 시청률 {video['avg_percentage']:.1f}%  |  "
                            f"🎞️ {video['duration']}"
                        )
                    st.link_button(
                        "▶️ YouTube에서 보기",
                        "https://www.youtube.com/watch?v=" + video["video_id"],
                        key=f"rank_{metric_name}_{video['video_id']}"
                    )
                st.divider()

        with rank_tab1:
            render_ranked_videos(
                sorted(public_videos, key=lambda x: x["views"], reverse=True),
                "조회수", lambda v: f"{v['views']:,}회"
            )

        with rank_tab2:
            render_ranked_videos(
                sorted(public_videos, key=lambda x: x["avg_percentage"], reverse=True),
                "평균 시청률", lambda v: f"{v['avg_percentage']:.1f}%"
            )

        with rank_tab3:
            render_ranked_videos(
                sorted(public_videos, key=lambda x: x["sub_conversion_rate"], reverse=True),
                "구독전환율", lambda v: f"{v['sub_conversion_rate']:.3f}%"
            )

        with rank_tab4:
            render_ranked_videos(
                sorted(public_videos, key=lambda x: x["like_rate"], reverse=True),
                "좋아요율", lambda v: f"{v['like_rate']:.2f}%"
            )



    # =========================================================
    # 23. 예약 영상
    # =========================================================

    if scheduled_videos:

        st.divider()

        st.subheader(
            "🟡 예약 영상"
        )


        for video in scheduled_videos:

            col_img, col_info = (
                st.columns(
                    [1, 5]
                )
            )


            with col_img:

                if video[
                    "thumbnail"
                ]:

                    st.image(
                        video[
                            "thumbnail"
                        ],
                        width=140
                    )


            with col_info:

                st.markdown(
                    f"**{video['title']}**"
                )

                st.write(
                    f"공개 예정: "
                    f"{video['scheduled']}"
                )

                st.write(
                    f"영상 길이: "
                    f"{video['duration']}"
                )


    # =========================================================
    # 24. 연결 해제
    # =========================================================

    st.divider()


    if st.button(
        "🔓 YouTube 연결 해제"
    ):

        st.session_state.clear()

        st.query_params.clear()

        st.rerun()

import os
from datetime import date, datetime, timedelta, timezone
from zoneinfo import ZoneInfo
import calendar

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

# localhost 개발용
os.environ["OAUTHLIB_INSECURE_TRANSPORT"] = "1"

CLIENT_ID = os.getenv("GOOGLE_CLIENT_ID")
CLIENT_SECRET = os.getenv("GOOGLE_CLIENT_SECRET")

REDIRECT_URI = "https://alssulcut-dashboard-ukejv5o8kkinevripbvu97.streamlit.app"

SCOPES = [
    "https://www.googleapis.com/auth/youtube.readonly",
    "https://www.googleapis.com/auth/yt-analytics.readonly",
]

st.set_page_config(
    page_title="알쓸컷 쇼츠 분석기",
    page_icon="📊",
    layout="wide",
)

st.title("📊 알쓸컷 쇼츠 분석기")

st.write(
    "YouTube 채널을 연결하고 "
    "쇼츠 성과를 한눈에 확인하세요."
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
            "Google 로그인 처리 중 "
            "오류가 발생했습니다."
        )

        st.code(
            str(e)
        )

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
        "Google 계정으로 로그인한 뒤 "
        "분석할 YouTube 채널을 선택하세요."
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

    st.error(
        "YouTube API 연결 실패"
    )

    st.code(
        str(e)
    )

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

    st.error(
        "YouTube 데이터를 가져오지 못했습니다."
    )

    st.code(
        str(e)
    )

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
        "오류 내용"
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

st.success(
    "✅ YouTube 채널 연결 완료!"
)

st.subheader(
    f"📺 {channel_info['channel_name']}"
)


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
    "YouTube 공개 영상",
    f"{channel_info['public_video_count']:,}개"
)

st.divider()


# =========================================================
# 11. 영상 현황
# =========================================================

st.subheader(
    "🎬 영상 현황"
)

c1, c2, c3, c4, c5 = (
    st.columns(5)
)

c1.metric(
    "전체 감지",
    f"{len(videos)}개"
)

c2.metric(
    "🟢 공개",
    f"{len(public_videos)}개"
)

c3.metric(
    "🟡 예약",
    f"{len(scheduled_videos)}개"
)

c4.metric(
    "🔒 비공개",
    f"{len(private_videos)}개"
)

c5.metric(
    "🔵 일부공개",
    f"{len(unlisted_videos)}개"
)

st.caption(
    "※ 예약·비공개·일부공개 영상은 "
    "성과 평균과 순위에서 제외됩니다."
)

st.divider()


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
# 기간 버튼
# ---------------------------------------------------------

period_option = st.radio(

    "분석 기간",

    [
        "오늘",
        "최근 7일",
        "최근 28일",
        "직접 선택",
    ],

    horizontal=True,
    index=1,
)


today = datetime.now(KST).date()


if period_option == "오늘":

    start_date = today
    end_date = today


elif period_option == "최근 7일":

    # 오늘 Analytics는 집계 지연이 있을 수 있으므로
    # 완료 데이터 기준으로 어제까지 7일을 사용
    end_date = today - timedelta(days=1)
    start_date = end_date - timedelta(days=6)


elif period_option == "최근 28일":

    # 오늘을 제외한 최근 완료 데이터 28일
    end_date = today - timedelta(days=1)
    start_date = end_date - timedelta(days=27)


else:

    selected_range = (
        st.date_input(
            "날짜 범위 선택",
            value=(
                today
                - timedelta(days=6),
                today
            ),
        )
    )

    if isinstance(
        selected_range,
        tuple
    ):

        if len(
            selected_range
        ) == 2:

            start_date = (
                selected_range[0]
            )

            end_date = (
                selected_range[1]
            )

        else:

            start_date = today
            end_date = today

    else:

        start_date = (
            selected_range
        )

        end_date = (
            selected_range
        )


st.caption(
    f"분석 기간: "
    f"{start_date} ~ {end_date}"
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

    current_summary = (
        get_period_summary(
            yt_analytics,
            start_date,
            end_date,
        )
    )

    (
        previous_start,
        previous_end,
    ) = get_previous_period(
        start_date,
        end_date,
    )

    previous_summary = (
        get_period_summary(
            yt_analytics,
            previous_start,
            previous_end,
        )
    )

except Exception as e:

    st.error(
        "기간별 Analytics 데이터를 "
        "가져오지 못했습니다."
    )

    st.code(
        str(e)
    )

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

c1, c2, c3, c4 = (
    st.columns(4)
)


c1.metric(
    "조회수",
    f"{current_summary['views']:,}회",
    change_text(
        current_summary["views"],
        previous_summary["views"],
    ),
)


c2.metric(
    "순구독자",
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
    ),
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
    ),
)


c4.metric(
    "좋아요",
    f"{current_summary['likes']:,}개",
    change_text(
        current_summary["likes"],
        previous_summary["likes"],
    ),
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

    st.warning(
        "일별 데이터를 불러오지 못했습니다."
    )

    st.code(
        str(e)
    )


if daily_data:

    daily_df = pd.DataFrame(daily_data)
    daily_df["date"] = pd.to_datetime(daily_df["date"])

    # 데이터가 없는 날짜도 0으로 채움
    full_dates = pd.date_range(start=start_date, end=end_date, freq="D")
    daily_df = daily_df.set_index("date").reindex(full_dates).fillna(0)
    daily_df.index.name = "날짜"
    daily_df.index = [dt.strftime("%m/%d") for dt in daily_df.index]

    st.subheader("📈 일별 조회수")
    st.line_chart(daily_df[["views"]], use_container_width=True)

    st.subheader("👤 일별 순구독자")
    st.bar_chart(daily_df[["net_subscribers"]], use_container_width=True)

else:
    st.info("선택한 기간에 일별 Analytics 데이터가 없습니다.")

st.divider()


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
        day_data = month_lookup.get(current_day.strftime("%Y-%m-%d"), {})
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

        # 성과가 없는 날은 최대한 단순하게 표시
        if views == 0 and net_subscribers == 0:
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
st.divider()

# =========================================================
# 18. 특정 날짜 선택
# =========================================================

st.header(
    "🔎 하루 자세히 보기"
)

selected_day = st.date_input(
    "확인할 날짜",
    value=st.session_state.get("calendar_selected_day", end_date),
    min_value=date(2005, 1, 1),
    max_value=today,
)
st.session_state.calendar_selected_day = selected_day


try:

    day_summary = (
        get_period_summary(
            yt_analytics,
            selected_day,
            selected_day,
        )
    )

except Exception as e:

    st.error(
        "선택 날짜 데이터를 "
        "가져오지 못했습니다."
    )

    st.code(
        str(e)
    )

    day_summary = None


if day_summary:

    st.subheader(
        f"📅 {selected_day}"
    )

    c1, c2, c3, c4 = (
        st.columns(4)
    )

    c1.metric(
        "그날 조회수",
        f"+{day_summary['views']:,}회"
    )

    c2.metric(
        "그날 순구독자",
        (
            f"{day_summary['net_subscribers']:+,}명"
        )
    )

    c3.metric(
        "그날 시청시간",
        format_watch_time(
            day_summary[
                "watch_minutes"
            ]
        )
    )

    c4.metric(
        "그날 좋아요",
        f"+{day_summary['likes']:,}개"
    )


    # =====================================================
    # 그날 영상별 성과
    # =====================================================

    try:

        day_video_data = (
            get_video_performance_for_day(
                yt_analytics,
                selected_day,
            )
        )

    except Exception as e:

        day_video_data = []

        st.warning(
            "그날의 영상별 데이터를 "
            "불러오지 못했습니다."
        )

        st.code(
            str(e)
        )


    video_lookup = {

        video["video_id"]:
            video

        for video in videos
    }

    # 당일 업로드 영상 vs 기존 영상 조회수 기여도
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

    new_video_views = sum(
        item["views"] for item in day_video_data
        if item["video_id"] in uploaded_ids_for_day
    )
    old_video_views = max(day_summary["views"] - new_video_views, 0)

    st.subheader("🧩 그날 조회수 구성")

    day_total_views = max(day_summary["views"], 0)

    if day_total_views > 0:
        new_video_share = (new_video_views / day_total_views) * 100
        old_video_share = (old_video_views / day_total_views) * 100
    else:
        new_video_share = 0.0
        old_video_share = 0.0

    cc1, cc2 = st.columns(2)

    cc1.metric(
        "당일 업로드 영상",
        f"{new_video_views:,}회",
        f"{new_video_share:.1f}% 기여",
    )

    cc2.metric(
        "기존 영상",
        f"{old_video_views:,}회",
        f"{old_video_share:.1f}% 기여",
    )

    if selected_day == today and day_total_views == 0:
        st.caption(
            "※ 오늘 Analytics가 아직 반영되지 않았다면 위 구성도 0으로 표시될 수 있습니다."
        )


    if day_video_data:

        st.subheader(
            "🔥 그날 조회수를 만든 영상"
        )

        for rank, performance in enumerate(
            day_video_data[:5],
            start=1,
        ):

            video = video_lookup.get(
                performance["video_id"]
            )

            if not video:
                continue

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
                    f"### {rank}위 · "
                    f"{video['title']}"
                )

                st.write(
                    f"👁️ 그날 "
                    f"+{performance['views']:,}회"
                    f"  |  "
                    f"👍 "
                    f"+{performance['likes']:,}"
                    f"  |  "
                    f"👤 "
                    f"{performance['net_subscribers']:+d}"
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


# =========================================================
# 19. 공개 영상 전체 성과
# =========================================================

st.header(
    "🎯 전체 공개 영상 분석"
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


    c1, c2, c3, c4 = (
        st.columns(4)
    )

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
        "영상 순구독자",
        f"{total_net_subscribers:+,}명"
    )


st.divider()


# =========================================================
# 20. 자동 성과진단 V1
# =========================================================

st.header("🚦 자동 성과진단 V1")
st.caption(
    "※ 성과점수는 YouTube 공식 점수가 아니라, 현재 채널의 평가 가능한 공개 영상끼리 "
    "조회수·평균 시청률·좋아요율·구독전환율을 비교한 '채널 내 상대점수'입니다. "
    "조회수 100회 미만 또는 Analytics 데이터가 부족한 영상은 평가를 보류합니다."
)

try:
    recent7_end = today - timedelta(days=1)
    recent7_start = recent7_end - timedelta(days=6)

    recent7 = get_period_summary(
        yt_analytics,
        recent7_start,
        recent7_end,
    )

    prev7_start, prev7_end = get_previous_period(
        recent7_start,
        recent7_end,
    )

    prev7 = get_period_summary(
        yt_analytics,
        prev7_start,
        prev7_end,
    )

    view_change = calculate_change(
        recent7["views"],
        prev7["views"],
    )

    if recent7["views"] == 0:
        channel_state = "⚪ 집계 대기"
        channel_message = "최근 데이터가 아직 충분히 집계되지 않았습니다."
    elif view_change is None:
        channel_state = "🟢 성장 시작"
        channel_message = "이전 비교기간의 조회수가 거의 없어 최근 성장이 새로 잡히고 있습니다."
    elif view_change >= 25:
        channel_state = "🔥 강한 상승세"
        channel_message = f"최근 7일 조회수가 이전 7일보다 {view_change:+.1f}% 증가했습니다."
    elif view_change >= 5:
        channel_state = "🟢 상승세"
        channel_message = f"최근 7일 조회수가 이전 7일보다 {view_change:+.1f}% 증가했습니다."
    elif view_change <= -25:
        channel_state = "🔴 하락세"
        channel_message = f"최근 7일 조회수가 이전 7일보다 {view_change:+.1f}% 감소했습니다."
    elif view_change <= -5:
        channel_state = "🟡 약한 하락"
        channel_message = f"최근 7일 조회수가 이전 7일보다 {view_change:+.1f}% 감소했습니다."
    else:
        channel_state = "🟡 보합"
        channel_message = "최근 7일 조회수는 이전 7일과 비슷한 수준입니다."

    st.subheader(f"채널 상태: {channel_state}")
    st.write(channel_message)

    dc1, dc2, dc3 = st.columns(3)
    dc1.metric(
        "최근 7일 조회수",
        f"{recent7['views']:,}회",
    )
    dc2.metric(
        "최근 7일 순구독자",
        f"{recent7['net_subscribers']:+,}명",
    )
    dc3.metric(
        "최근 7일 시청시간",
        format_watch_time(recent7["watch_minutes"]),
    )

except Exception as e:
    st.warning("채널 상태 진단을 불러오지 못했습니다.")
    with st.expander("진단 오류 보기"):
        st.code(str(e))


if eligible_videos:
    diagnosed = sorted(
        eligible_videos,
        key=lambda x: x.get("performance_score", 0),
        reverse=True,
    )

    best_diagnosis = diagnosed[0]
    weakest_diagnosis = diagnosed[-1]

    best_col, weak_col = st.columns(2)

    with best_col:
        st.subheader("🏅 종합 성과 1위")
        st.markdown(
            f"**{best_diagnosis['title']}**  \n"
            f"{best_diagnosis['performance_grade']} · "
            f"**{best_diagnosis['performance_score']}점**"
        )

        strengths = best_diagnosis.get("diagnosis_strengths", [])
        if strengths:
            st.write("강점: " + ", ".join(strengths))

        st.write(
            f"조회수 {best_diagnosis['views']:,}회 · "
            f"시청률 {best_diagnosis['avg_percentage']:.1f}% · "
            f"좋아요율 {best_diagnosis['like_rate']:.2f}% · "
            f"구독전환율 {best_diagnosis['sub_conversion_rate']:.3f}%"
        )

    with weak_col:
        st.subheader("🛠️ 개선 우선 영상")
        st.markdown(
            f"**{weakest_diagnosis['title']}**  \n"
            f"{weakest_diagnosis['performance_grade']} · "
            f"**{weakest_diagnosis['performance_score']}점**"
        )

        weaknesses = weakest_diagnosis.get("diagnosis_weaknesses", [])
        if weaknesses:
            st.write("우선 점검: " + ", ".join(weaknesses))

        st.write(
            f"조회수 {weakest_diagnosis['views']:,}회 · "
            f"시청률 {weakest_diagnosis['avg_percentage']:.1f}% · "
            f"좋아요율 {weakest_diagnosis['like_rate']:.2f}% · "
            f"구독전환율 {weakest_diagnosis['sub_conversion_rate']:.3f}%"
        )

    with st.expander(
        "📋 모든 공개 영상 성과점수 보기",
        expanded=False,
    ):
        diagnosis_table = pd.DataFrame([
            {
                "영상": video["title"],
                "점수": video.get("performance_score", 0),
                "등급": video.get("performance_grade", "-"),
                "조회수": video["views"],
                "평균 시청률": f"{video['avg_percentage']:.1f}%",
                "좋아요율": f"{video['like_rate']:.2f}%",
                "구독전환율": f"{video['sub_conversion_rate']:.3f}%",
            }
            for video in diagnosed
        ])

        st.dataframe(
            diagnosis_table,
            use_container_width=True,
            hide_index=True,
        )

st.divider()


# =========================================================
# 21. 영상 TOP 랭킹
# =========================================================

st.subheader("🏆 공개 영상 TOP 랭킹")

rank_tab1, rank_tab2, rank_tab3, rank_tab4 = st.tabs(
    ["👁️ 조회수", "📊 시청률", "👤 구독전환", "👍 좋아요율"]
)

def render_ranked_videos(ranked, metric_name, metric_formatter):
    if not ranked:
        st.info("표시할 공개 영상이 없습니다.")
        return

    for rank, video in enumerate(ranked[:5], start=1):
        col_img, col_info = st.columns([1, 5])
        with col_img:
            if video["thumbnail"]:
                st.image(video["thumbnail"], width=160)
        with col_info:
            st.markdown(f"### {rank}위 · {video['title']}")
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
# 22. 전체 영상 표
# =========================================================

st.subheader(
    "📋 전체 영상"
)


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
            video["published"],

        "예약 공개":
            video["scheduled"],
    })


df = pd.DataFrame(
    table_data
)


with st.expander("전체 영상 표 펼쳐보기", expanded=False):
    st.dataframe(
        df,
        use_container_width=True,
        hide_index=True,
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

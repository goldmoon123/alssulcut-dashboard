import os
from datetime import date, datetime, timedelta, timezone
from zoneinfo import ZoneInfo
from io import BytesIO
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
    page_title="알쓸컷 쇼츠 분석기",
    page_icon="📊",
    layout="wide",
)

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
            padding: 0.45rem 0.55rem !important;
            border: 1px solid rgba(128,128,128,.18);
            border-radius: 12px;
            min-height: 92px;
        }
        [data-testid="stMetricLabel"] { font-size: 0.82rem !important; }
        [data-testid="stMetricValue"] { font-size: 1.65rem !important; }
        [data-testid="stMetricDelta"] { font-size: 0.75rem !important; }

        /* 4칸 지표는 모바일에서 2 x 2 */
        [data-testid="stHorizontalBlock"]:has(> div:nth-child(4)):not(:has(> div:nth-child(5))) {
            display: grid !important;
            grid-template-columns: repeat(2, minmax(0, 1fr)) !important;
            gap: 0.55rem !important;
        }

        /* 5칸 현황도 너무 길어지지 않게 2열 */
        [data-testid="stHorizontalBlock"]:has(> div:nth-child(5)):not(:has(> div:nth-child(6))) {
            display: grid !important;
            grid-template-columns: repeat(2, minmax(0, 1fr)) !important;
            gap: 0.55rem !important;
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
    }
    </style>
    """,
    unsafe_allow_html=True,
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
    "총 조회수",
    f"{channel_info['total_views']:,}회"
)

c3.metric(
    "공개 영상",
    f"{channel_info['public_video_count']:,}개"
)

st.divider()


# =========================================================
# 11. 영상 현황
# =========================================================

st.subheader(
    "🎬 영상 현황"
)

c1, c2, c3 = st.columns(3)

c1.metric(
    "전체",
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

c4, c5 = st.columns(2)

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
# 기간 선택 + 분석하기 버튼
# ---------------------------------------------------------

today = datetime.now(KST).date()

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
    st.line_chart(daily_df[["views"]], use_container_width=True, height=260)

    st.subheader("👤 일별 순구독자")
    st.bar_chart(daily_df[["net_subscribers"]], use_container_width=True, height=260)

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

    dc1, dc2, dc3, _diag_blank = st.columns(4)
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
            height=300,
        )

st.divider()


# =========================================================
# 21. 영상 TOP 랭킹
# =========================================================

with st.expander("🏆 공개 영상 TOP 랭킹 보기", expanded=False):
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
# 22. 전체 영상 표
# =========================================================

st.subheader(
    "🔎 전체 영상 분석"
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


# =========================================================
# 22-1. 영상 검색 / 조회수 필터
# =========================================================

st.caption(
    "조건을 정한 뒤 검색 버튼을 눌러 결과를 확인하세요. "
    "검색 전에는 기존 결과가 그대로 유지됩니다."
)

# 조회수 조건은 바꾸면 입력칸만 바뀌고, 실제 결과는 검색 버튼을 눌러야 적용됩니다.
filter_mode_input = st.radio(
    "조회수 조건",
    ["전체", "이상", "이하", "범위"],
    horizontal=True,
    key="views_filter_mode_input",
)

filter_min_input = 0
filter_max_input = 0
filter_views_input = 5000

if filter_mode_input in ["이상", "이하"]:
    filter_views_input = st.number_input(
        "기준 조회수",
        min_value=0,
        value=5000,
        step=500,
        key="views_filter_value_input",
    )
elif filter_mode_input == "범위":
    range_c1, range_c2 = st.columns(2)
    with range_c1:
        filter_min_input = st.number_input(
            "최소 조회수",
            min_value=0,
            value=1000,
            step=500,
            key="views_filter_min_input",
        )
    with range_c2:
        filter_max_input = st.number_input(
            "최대 조회수",
            min_value=0,
            value=10000,
            step=500,
            key="views_filter_max_input",
        )

filter_c1, filter_c2 = st.columns(2)
with filter_c1:
    status_options = ["전체"] + sorted(df["상태"].dropna().astype(str).unique().tolist())
    filter_status_input = st.selectbox(
        "공개 상태",
        status_options,
        key="video_status_filter_input",
    )
with filter_c2:
    title_query_input = st.text_input(
        "제목 검색",
        placeholder="예: 비버, 화산, 교통사고",
        key="video_title_filter_input",
    )

sort_option_input = st.selectbox(
    "정렬",
    [
        "조회수 높은 순",
        "조회수 낮은 순",
        "최신 업로드 순",
        "오래된 업로드 순",
        "평균 시청률 높은 순",
        "구독 증가 높은 순",
    ],
    key="video_sort_input",
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

if st.button(
    "🔍 검색",
    type="primary",
    use_container_width=True,
    key="apply_video_filter_button",
):
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
    """사용자가 열자마자 영상 데이터가 먼저 보이도록 첫 시트에 실제 데이터를 둡니다."""
    output = BytesIO()

    summary_rows = [
        ["채널명", channel_info.get("channel_name", channel_info.get("title", ""))],
        ["구독자", channel_info.get("subscribers", 0)],
        ["채널 총 조회수", channel_info.get("total_views", channel_info.get("views", 0))],
        ["전체 감지 영상", len(videos)],
        ["공개 영상", len(public_videos)],
        ["예약 영상", len(scheduled_videos)],
        ["내보낸 영상", len(primary_df)],
        ["생성 시각", datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S KST")],
    ]
    summary_df = pd.DataFrame(summary_rows, columns=["항목", "값"])

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # 핵심 수정: 영상 데이터 시트를 첫 번째로 저장
        primary_df.to_excel(writer, sheet_name=primary_sheet_name, index=False)
        summary_df.to_excel(writer, sheet_name="채널 요약", index=False)

        for sheet_name in writer.book.sheetnames:
            ws = writer.book[sheet_name]
            for column_cells in ws.columns:
                max_length = 0
                column_letter = column_cells[0].column_letter
                for cell in column_cells:
                    try:
                        cell_length = len(str(cell.value)) if cell.value is not None else 0
                        max_length = max(max_length, cell_length)
                    except Exception:
                        pass
                ws.column_dimensions[column_letter].width = min(max(max_length + 2, 10), 45)

            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions

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

with st.expander("📋 전체 영상 표 펼쳐보기", expanded=False):
    st.dataframe(
        df,
        use_container_width=True,
        hide_index=True,
        height=340,
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

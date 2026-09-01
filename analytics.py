from datetime import date, timedelta

from googleapiclient.discovery import build


# =========================================================
# YouTube Analytics API 생성
# =========================================================

def create_analytics_api(credentials):

    return build(
        "youtubeAnalytics",
        "v2",
        credentials=credentials,
        cache_discovery=False,
    )


# =========================================================
# 날짜 형식 변환
# =========================================================

def to_date_string(value):

    if isinstance(value, str):
        return value

    return value.strftime("%Y-%m-%d")


# =========================================================
# 채널 기간 요약 데이터
# =========================================================

def get_period_summary(
    analytics,
    start_date,
    end_date,
):

    start_date = to_date_string(
        start_date
    )

    end_date = to_date_string(
        end_date
    )

    response = (
        analytics
        .reports()
        .query(
            ids="channel==MINE",
            startDate=start_date,
            endDate=end_date,
            metrics=(
                "views,"
                "estimatedMinutesWatched,"
                "averageViewDuration,"
                "averageViewPercentage,"
                "likes,"
                "comments,"
                "shares,"
                "subscribersGained,"
                "subscribersLost"
            ),
        )
        .execute()
    )

    rows = response.get(
        "rows",
        []
    )

    if not rows:

        return {
            "views": 0,
            "watch_minutes": 0,
            "average_view_duration": 0,
            "average_view_percentage": 0,
            "likes": 0,
            "comments": 0,
            "shares": 0,
            "subscribers_gained": 0,
            "subscribers_lost": 0,
            "net_subscribers": 0,
        }

    headers = [
        header["name"]
        for header in response.get(
            "columnHeaders",
            []
        )
    ]

    data = dict(
        zip(
            headers,
            rows[0]
        )
    )

    subscribers_gained = int(
        data.get(
            "subscribersGained",
            0
        )
    )

    subscribers_lost = int(
        data.get(
            "subscribersLost",
            0
        )
    )

    return {

        "views":
            int(
                data.get(
                    "views",
                    0
                )
            ),

        "watch_minutes":
            float(
                data.get(
                    "estimatedMinutesWatched",
                    0
                )
            ),

        "average_view_duration":
            float(
                data.get(
                    "averageViewDuration",
                    0
                )
            ),

        "average_view_percentage":
            float(
                data.get(
                    "averageViewPercentage",
                    0
                )
            ),

        "likes":
            int(
                data.get(
                    "likes",
                    0
                )
            ),

        "comments":
            int(
                data.get(
                    "comments",
                    0
                )
            ),

        "shares":
            int(
                data.get(
                    "shares",
                    0
                )
            ),

        "subscribers_gained":
            subscribers_gained,

        "subscribers_lost":
            subscribers_lost,

        "net_subscribers":
            (
                subscribers_gained
                - subscribers_lost
            ),
    }


# =========================================================
# 일별 채널 데이터
# =========================================================

def get_daily_channel_data(
    analytics,
    start_date,
    end_date,
):

    start_date = to_date_string(
        start_date
    )

    end_date = to_date_string(
        end_date
    )

    response = (
        analytics
        .reports()
        .query(
            ids="channel==MINE",
            startDate=start_date,
            endDate=end_date,
            dimensions="day",
            metrics=(
                "views,"
                "estimatedMinutesWatched,"
                "likes,"
                "comments,"
                "shares,"
                "subscribersGained,"
                "subscribersLost"
            ),
            sort="day",
            maxResults=500,
        )
        .execute()
    )

    headers = [
        header["name"]
        for header in response.get(
            "columnHeaders",
            []
        )
    ]

    daily_data = []

    for row in response.get(
        "rows",
        []
    ):

        data = dict(
            zip(
                headers,
                row
            )
        )

        subscribers_gained = int(
            data.get(
                "subscribersGained",
                0
            )
        )

        subscribers_lost = int(
            data.get(
                "subscribersLost",
                0
            )
        )

        daily_data.append({

            "date":
                data.get(
                    "day"
                ),

            "views":
                int(
                    data.get(
                        "views",
                        0
                    )
                ),

            "watch_minutes":
                float(
                    data.get(
                        "estimatedMinutesWatched",
                        0
                    )
                ),

            "likes":
                int(
                    data.get(
                        "likes",
                        0
                    )
                ),

            "comments":
                int(
                    data.get(
                        "comments",
                        0
                    )
                ),

            "shares":
                int(
                    data.get(
                        "shares",
                        0
                    )
                ),

            "subscribers_gained":
                subscribers_gained,

            "subscribers_lost":
                subscribers_lost,

            "net_subscribers":
                (
                    subscribers_gained
                    - subscribers_lost
                ),
        })

    return daily_data


# =========================================================
# 특정 날짜에 영상별로 발생한 성과
# =========================================================

def get_video_performance_for_day(
    analytics,
    selected_date,
):

    selected_date = to_date_string(
        selected_date
    )

    response = (
        analytics
        .reports()
        .query(
            ids="channel==MINE",
            startDate=selected_date,
            endDate=selected_date,
            dimensions="video",
            metrics=(
                "views,"
                "estimatedMinutesWatched,"
                "likes,"
                "comments,"
                "shares,"
                "subscribersGained,"
                "subscribersLost"
            ),
            sort="-views",
            maxResults=200,
        )
        .execute()
    )

    headers = [
        header["name"]
        for header in response.get(
            "columnHeaders",
            []
        )
    ]

    videos = []

    for row in response.get(
        "rows",
        []
    ):

        data = dict(
            zip(
                headers,
                row
            )
        )

        subscribers_gained = int(
            data.get(
                "subscribersGained",
                0
            )
        )

        subscribers_lost = int(
            data.get(
                "subscribersLost",
                0
            )
        )

        videos.append({

            "video_id":
                data.get(
                    "video"
                ),

            "views":
                int(
                    data.get(
                        "views",
                        0
                    )
                ),

            "watch_minutes":
                float(
                    data.get(
                        "estimatedMinutesWatched",
                        0
                    )
                ),

            "likes":
                int(
                    data.get(
                        "likes",
                        0
                    )
                ),

            "comments":
                int(
                    data.get(
                        "comments",
                        0
                    )
                ),

            "shares":
                int(
                    data.get(
                        "shares",
                        0
                    )
                ),

            "subscribers_gained":
                subscribers_gained,

            "subscribers_lost":
                subscribers_lost,

            "net_subscribers":
                (
                    subscribers_gained
                    - subscribers_lost
                ),
        })

    return videos


# =========================================================
# 영상별 전체 Analytics
# =========================================================

def get_video_analytics(
    analytics,
    video_ids,
    start_date="2005-01-01",
    end_date=None,
):

    if not video_ids:
        return {}

    if end_date is None:
        end_date = date.today()

    start_date = to_date_string(
        start_date
    )

    end_date = to_date_string(
        end_date
    )

    analytics_by_video = {}

    # API 필터가 너무 길어지는 것을 막기 위해
    # 50개씩 나눠서 요청
    for i in range(
        0,
        len(video_ids),
        50
    ):

        batch = video_ids[
            i:i + 50
        ]

        response = (
            analytics
            .reports()
            .query(
                ids="channel==MINE",
                startDate=start_date,
                endDate=end_date,
                dimensions="video",
                filters=(
                    "video=="
                    + ",".join(batch)
                ),
                metrics=(
                    "views,"
                    "estimatedMinutesWatched,"
                    "averageViewDuration,"
                    "averageViewPercentage,"
                    "likes,"
                    "comments,"
                    "shares,"
                    "subscribersGained,"
                    "subscribersLost"
                ),
                maxResults=200,
            )
            .execute()
        )

        headers = [
            header["name"]
            for header in response.get(
                "columnHeaders",
                []
            )
        ]

        for row in response.get(
            "rows",
            []
        ):

            data = dict(
                zip(
                    headers,
                    row
                )
            )

            video_id = data.get(
                "video"
            )

            if not video_id:
                continue

            subscribers_gained = int(
                data.get(
                    "subscribersGained",
                    0
                )
            )

            subscribers_lost = int(
                data.get(
                    "subscribersLost",
                    0
                )
            )

            analytics_by_video[
                video_id
            ] = {

                "views":
                    int(
                        data.get(
                            "views",
                            0
                        )
                    ),

                "watch_minutes":
                    float(
                        data.get(
                            "estimatedMinutesWatched",
                            0
                        )
                    ),

                "average_view_duration":
                    float(
                        data.get(
                            "averageViewDuration",
                            0
                        )
                    ),

                "average_view_percentage":
                    float(
                        data.get(
                            "averageViewPercentage",
                            0
                        )
                    ),

                "likes":
                    int(
                        data.get(
                            "likes",
                            0
                        )
                    ),

                "comments":
                    int(
                        data.get(
                            "comments",
                            0
                        )
                    ),

                "shares":
                    int(
                        data.get(
                            "shares",
                            0
                        )
                    ),

                "subscribers_gained":
                    subscribers_gained,

                "subscribers_lost":
                    subscribers_lost,

                "net_subscribers":
                    (
                        subscribers_gained
                        - subscribers_lost
                    ),
            }

    return analytics_by_video


# =========================================================
# 이전 동일 기간 계산
#
# 예:
# 현재 8/26 ~ 9/1 = 7일
# 이전 8/19 ~ 8/25 = 7일
# =========================================================

def get_previous_period(
    start_date,
    end_date,
):

    period_days = (
        end_date - start_date
    ).days + 1

    previous_end = (
        start_date
        - timedelta(days=1)
    )

    previous_start = (
        previous_end
        - timedelta(
            days=period_days - 1
        )
    )

    return (
        previous_start,
        previous_end
    )


# =========================================================
# 증감률 계산
# =========================================================

def calculate_change(
    current,
    previous,
):

    current = float(
        current or 0
    )

    previous = float(
        previous or 0
    )

    if previous == 0:

        if current == 0:
            return 0

        return None

    return (
        (
            current - previous
        )
        / previous
    ) * 100
from googleapiclient.discovery import build

from utils import (
    duration_to_seconds,
    format_duration,
    format_youtube_datetime,
)


# =========================================================
# YouTube API 생성
# =========================================================

def create_youtube_api(credentials):
    return build(
        "youtube",
        "v3",
        credentials=credentials,
        cache_discovery=False,
    )


# =========================================================
# 내 채널 정보 가져오기
# =========================================================

def get_channel_info(youtube):

    response = youtube.channels().list(
        part="snippet,statistics,contentDetails",
        mine=True,
    ).execute()

    channels = response.get("items", [])

    if not channels:
        return None

    channel = channels[0]

    statistics = channel.get(
        "statistics",
        {}
    )

    return {
        "channel_id": channel["id"],

        "channel_name": (
            channel
            .get("snippet", {})
            .get("title", "채널 이름 없음")
        ),

        "subscribers": int(
            statistics.get(
                "subscriberCount",
                0
            )
        ),

        "total_views": int(
            statistics.get(
                "viewCount",
                0
            )
        ),

        "public_video_count": int(
            statistics.get(
                "videoCount",
                0
            )
        ),

        "uploads_playlist_id": (
            channel
            ["contentDetails"]
            ["relatedPlaylists"]
            ["uploads"]
        ),
    }


# =========================================================
# 업로드 재생목록의 영상 ID 전부 가져오기
# =========================================================

def get_upload_video_ids(
    youtube,
    uploads_playlist_id
):

    video_ids = []

    next_page_token = None

    while True:

        response = youtube.playlistItems().list(
            part="contentDetails",
            playlistId=uploads_playlist_id,
            maxResults=50,
            pageToken=next_page_token,
        ).execute()

        for item in response.get(
            "items",
            []
        ):

            video_id = (
                item
                .get("contentDetails", {})
                .get("videoId")
            )

            if video_id:
                video_ids.append(
                    video_id
                )

        next_page_token = response.get(
            "nextPageToken"
        )

        if not next_page_token:
            break

    return video_ids


# =========================================================
# 영상 상세 정보 가져오기
# =========================================================

def get_video_details(
    youtube,
    video_ids
):

    videos = []

    if not video_ids:
        return videos

    # YouTube videos.list는 한 번에 최대 50개
    for i in range(
        0,
        len(video_ids),
        50
    ):

        batch = video_ids[
            i:i + 50
        ]

        response = youtube.videos().list(
            part=(
                "snippet,"
                "statistics,"
                "contentDetails,"
                "status"
            ),
            id=",".join(batch),
            maxResults=50,
        ).execute()

        for video in response.get(
            "items",
            []
        ):

            video_id = video["id"]

            snippet = video.get(
                "snippet",
                {}
            )

            statistics = video.get(
                "statistics",
                {}
            )

            content_details = video.get(
                "contentDetails",
                {}
            )

            status = video.get(
                "status",
                {}
            )


            # ---------------------------------------------
            # 기본 정보
            # ---------------------------------------------

            title = snippet.get(
                "title",
                "제목 없음"
            )

            published_at = snippet.get(
                "publishedAt"
            )

            thumbnail = (
                snippet
                .get("thumbnails", {})
                .get("medium", {})
                .get("url", "")
            )


            # ---------------------------------------------
            # 조회수 / 좋아요 / 댓글
            # ---------------------------------------------

            views = int(
                statistics.get(
                    "viewCount",
                    0
                )
            )

            likes = int(
                statistics.get(
                    "likeCount",
                    0
                )
            )

            comments = int(
                statistics.get(
                    "commentCount",
                    0
                )
            )


            # ---------------------------------------------
            # 영상 길이
            # ---------------------------------------------

            duration_raw = (
                content_details.get(
                    "duration",
                    "PT0S"
                )
            )

            duration_seconds = (
                duration_to_seconds(
                    duration_raw
                )
            )

            duration_text = (
                format_duration(
                    duration_seconds
                )
            )


            # ---------------------------------------------
            # 공개 상태
            # ---------------------------------------------

            privacy_status = status.get(
                "privacyStatus",
                "unknown"
            )

            publish_at = status.get(
                "publishAt"
            )


            if (
                privacy_status == "private"
                and publish_at
            ):
                video_status = "🟡 예약"

            elif privacy_status == "public":
                video_status = "🟢 공개"

            elif privacy_status == "unlisted":
                video_status = "🔵 일부공개"

            elif privacy_status == "private":
                video_status = "🔒 비공개"

            else:
                video_status = "⚪ 기타"


            # ---------------------------------------------
            # 최종 데이터
            # ---------------------------------------------

            videos.append({

                "video_id":
                    video_id,

                "title":
                    title,

                "thumbnail":
                    thumbnail,

                "published":
                    format_youtube_datetime(
                        published_at
                    ),

                # 나중에 날짜 분석에서 사용
                "published_raw":
                    published_at,

                "scheduled":
                    format_youtube_datetime(
                        publish_at
                    ),

                "scheduled_raw":
                    publish_at,

                "views":
                    views,

                "likes":
                    likes,

                "comments":
                    comments,

                "duration":
                    duration_text,

                "duration_seconds":
                    duration_seconds,

                "status":
                    video_status,

                "privacy_status":
                    privacy_status,
            })

    return videos


# =========================================================
# 채널의 모든 영상 한 번에 가져오기
# =========================================================

def get_all_channel_videos(
    youtube,
    channel_info
):

    uploads_playlist_id = (
        channel_info[
            "uploads_playlist_id"
        ]
    )

    video_ids = get_upload_video_ids(
        youtube,
        uploads_playlist_id
    )

    videos = get_video_details(
        youtube,
        video_ids
    )

    return videos


# =========================================================
# 상태별 영상 분리
# =========================================================

def split_videos_by_status(
    videos
):

    public_videos = [
        video
        for video in videos
        if video["status"] == "🟢 공개"
    ]

    scheduled_videos = [
        video
        for video in videos
        if video["status"] == "🟡 예약"
    ]

    private_videos = [
        video
        for video in videos
        if video["status"] == "🔒 비공개"
    ]

    unlisted_videos = [
        video
        for video in videos
        if video["status"] == "🔵 일부공개"
    ]

    return {
        "public": public_videos,
        "scheduled": scheduled_videos,
        "private": private_videos,
        "unlisted": unlisted_videos,
    }
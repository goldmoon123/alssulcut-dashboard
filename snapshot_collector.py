import os
import sys
import requests
from datetime import datetime, timezone

YOUTUBE_API_KEY = os.environ["YOUTUBE_API_KEY"]
YOUTUBE_CHANNEL_ID = os.environ["YOUTUBE_CHANNEL_ID"]
SUPABASE_URL = os.environ["SUPABASE_URL"].rstrip("/")
SUPABASE_SERVICE_ROLE_KEY = os.environ["SUPABASE_SERVICE_ROLE_KEY"]

YT = "https://www.googleapis.com/youtube/v3"


def yt_get(endpoint, params):
    params = dict(params)
    params["key"] = YOUTUBE_API_KEY
    r = requests.get(f"{YT}/{endpoint}", params=params, timeout=30)
    r.raise_for_status()
    return r.json()


def get_uploads_playlist():
    data = yt_get(
        "channels",
        {
            "part": "contentDetails",
            "id": YOUTUBE_CHANNEL_ID,
        },
    )
    items = data.get("items", [])
    if not items:
        raise RuntimeError("채널을 찾지 못했습니다. YOUTUBE_CHANNEL_ID를 확인하세요.")
    return items[0]["contentDetails"]["relatedPlaylists"]["uploads"]


def get_public_video_ids(playlist_id):
    ids = []
    page_token = None

    while True:
        params = {
            "part": "contentDetails",
            "playlistId": playlist_id,
            "maxResults": 50,
        }
        if page_token:
            params["pageToken"] = page_token

        data = yt_get("playlistItems", params)

        for item in data.get("items", []):
            video_id = item.get("contentDetails", {}).get("videoId")
            if video_id:
                ids.append(video_id)

        page_token = data.get("nextPageToken")
        if not page_token:
            break

    return ids


def chunks(seq, size=50):
    for i in range(0, len(seq), size):
        yield seq[i:i + size]


def get_video_stats(video_ids):
    rows = []

    for batch in chunks(video_ids, 50):
        data = yt_get(
            "videos",
            {
                "part": "snippet,statistics,status",
                "id": ",".join(batch),
                "maxResults": 50,
            },
        )

        for item in data.get("items", []):
            status = item.get("status", {})
            if status.get("privacyStatus") != "public":
                continue

            stats = item.get("statistics", {})
            snippet = item.get("snippet", {})

            rows.append(
                {
                    "video_id": item["id"],
                    "title": snippet.get("title", ""),
                    "published_at": snippet.get("publishedAt"),
                    "view_count": int(stats.get("viewCount", 0)),
                    "like_count": int(stats.get("likeCount", 0)),
                    "comment_count": int(stats.get("commentCount", 0)),
                }
            )

    return rows


def save_snapshots(rows):
    captured_at = datetime.now(timezone.utc).isoformat()

    payload = [
        {
            **row,
            "channel_id": YOUTUBE_CHANNEL_ID,
            "captured_at": captured_at,
        }
        for row in rows
    ]

    if not payload:
        return 0

    url = f"{SUPABASE_URL}/rest/v1/video_snapshots"
    headers = {
        "apikey": SUPABASE_SERVICE_ROLE_KEY,
        "Authorization": f"Bearer {SUPABASE_SERVICE_ROLE_KEY}",
        "Content-Type": "application/json",
        "Prefer": "return=minimal",
    }

    r = requests.post(url, headers=headers, json=payload, timeout=30)
    r.raise_for_status()
    return len(payload)


def main():
    playlist_id = get_uploads_playlist()
    video_ids = get_public_video_ids(playlist_id)
    rows = get_video_stats(video_ids)
    saved = save_snapshots(rows)

    print(f"공개 영상 {saved}개 스냅샷 저장 완료")
    print(f"기록 시각(UTC): {datetime.now(timezone.utc).isoformat()}")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"수집 실패: {e}", file=sys.stderr)
        raise

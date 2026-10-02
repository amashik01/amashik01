"""TikTok Creative Center Top Ads (Bangladesh) -> data/tiktok_ads.jsonl
Opens the public Top Ads page and captures the JSON the page itself loads.
Usage: python tiktok_scraper.py [--show]
"""
import json, os, sys
from playwright.sync_api import sync_playwright

URL = ("https://ads.tiktok.com/business/creativecenter/inspiration/topads/pc/en"
       "?region=BD&period=30&industry=")
KEYWORDS = ["shoes", "heel", "sandal", "slipper", "footwear", "juta", "ladies shoes"]


def collect(obj, acc):
    if isinstance(obj, dict):
        if "material_id" in obj or ("ad_title" in obj and "video_info" in obj):
            acc.append(obj)
        for v in obj.values(): collect(v, acc)
    elif isinstance(obj, list):
        for v in obj: collect(v, acc)


def main(headless=True):
    os.makedirs("data", exist_ok=True)
    rows, seen = [], set()
    with sync_playwright() as p:
        b = p.chromium.launch(headless=headless)
        page = b.new_context(locale="en-US").new_page()
        bucket = []
        def on_resp(r):
            if "top_ads" in r.url or "creative_radar" in r.url:
                try: collect(r.json(), bucket)
                except Exception: pass
        page.on("response", on_resp)
        for kw in KEYWORDS:
            bucket.clear()
            page.goto(URL + f"&keyword={kw}", wait_until="domcontentloaded", timeout=60000)
            page.wait_for_timeout(5000)
            for _ in range(8):
                page.mouse.wheel(0, 5000); page.wait_for_timeout(2000)
            for a in bucket:
                mid = str(a.get("material_id") or a.get("id"))
                if mid in seen: continue
                seen.add(mid)
                vi = a.get("video_info") or {}
                rows.append({
                    "material_id": mid,
                    "detail_url": f"https://ads.tiktok.com/business/creativecenter/topads/{mid}/pc/en",
                    "brand": a.get("brand_name") or "", "title": a.get("ad_title") or "",
                    "likes": a.get("like"), "comments": a.get("comment"), "shares": a.get("share"),
                    "ctr": a.get("ctr"), "industry": a.get("industry_key"),
                    "cover": vi.get("cover"), "video": (vi.get("video_url") or {}).get("720p"),
                    "keyword": kw,
                })
            print(kw, len(rows))
        b.close()
    with open("data/tiktok_ads.jsonl", "w", encoding="utf-8") as f:
        for r in rows: f.write(json.dumps(r, ensure_ascii=False) + "\n")


if __name__ == "__main__":
    main("--show" not in sys.argv)

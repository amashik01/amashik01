"""Facebook Ad Library (Bangladesh) keyword scraper -> data/fb_ads.jsonl

Usage:
  python fb_scraper.py --limit 20            # pilot: first 20 keywords
  python fb_scraper.py                       # all keywords
Resumable: keywords already in data/done_keywords.txt are skipped.
Needs: pip install playwright && playwright install chromium
"""
import argparse, json, os, random, time
from urllib.parse import quote
from playwright.sync_api import sync_playwright
from keywords import build

OUT = "data/fb_ads.jsonl"
DONE = "data/done_keywords.txt"
URL = ("https://www.facebook.com/ads/library/?active_status=all&ad_type=all&country=BD"
       "&q={q}&search_type=keyword_unordered&media_type=all")


def find_ads(obj, acc):
    """Recursively collect dicts that look like an ad (have ad_archive_id)."""
    if isinstance(obj, dict):
        if "ad_archive_id" in obj and "snapshot" in obj:
            acc.append(obj)
        for v in obj.values():
            find_ads(v, acc)
    elif isinstance(obj, list):
        for v in obj:
            find_ads(v, acc)


def norm(ad, kw):
    s = ad.get("snapshot") or {}
    body = (s.get("body") or {}).get("text") or ""
    imgs = [i.get("original_image_url") or i.get("resized_image_url") for i in s.get("images") or []]
    vids = [v.get("video_hd_url") or v.get("video_sd_url") for v in s.get("videos") or []]
    cards = s.get("cards") or []
    for c in cards:
        if c.get("original_image_url"): imgs.append(c["original_image_url"])
        if c.get("video_hd_url"): vids.append(c["video_hd_url"])
    pid = str(ad.get("page_id") or s.get("page_id") or "")
    return {
        "ad_id": str(ad["ad_archive_id"]),
        "ad_library_url": f"https://www.facebook.com/ads/library/?id={ad['ad_archive_id']}",
        "page_id": pid,
        "page_name": ad.get("page_name") or s.get("page_name") or "",
        "page_url": f"https://www.facebook.com/{pid}" if pid else "",
        "ad_copy": body,
        "title": s.get("title") or "",
        "cta": s.get("cta_text") or "",
        "landing_url": s.get("link_url") or "",
        "display_format": s.get("display_format") or "",
        "start_date": ad.get("start_date"),
        "end_date": ad.get("end_date"),
        "is_active": ad.get("is_active"),
        "platforms": ad.get("publisher_platform") or [],
        "collation_count": ad.get("collation_count"),
        "image_urls": [i for i in imgs if i],   # CDN links expire; ad_library_url is permanent
        "video_urls": [v for v in vids if v],
        "keyword": kw,
    }


def run(limit, scrolls, headless):
    os.makedirs("data", exist_ok=True)
    done = set(open(DONE, encoding="utf-8").read().split("\n")) if os.path.exists(DONE) else set()
    seen = set()
    if os.path.exists(OUT):
        for line in open(OUT, encoding="utf-8"):
            seen.add((json.loads(line)["ad_id"], json.loads(line)["keyword"]))
    kws = [k for k in build() if k not in done][: limit or None]
    with sync_playwright() as p:
        b = p.chromium.launch(headless=headless)
        ctx = b.new_context(locale="en-US", viewport={"width": 1366, "height": 900})
        page = ctx.new_page()
        bucket = []

        def on_resp(r):
            if "graphql" in r.url or "ads/library" in r.url:
                try:
                    txt = r.text()
                except Exception:
                    return
                for chunk in txt.split("\n"):
                    chunk = chunk.strip()
                    if chunk.startswith("{") or chunk.startswith("for (;;);"):
                        try:
                            find_ads(json.loads(chunk.replace("for (;;);", "")), bucket)
                        except Exception:
                            pass
        page.on("response", on_resp)
        for kw in kws:
            bucket.clear()
            try:
                page.goto(URL.format(q=quote(kw)), wait_until="domcontentloaded", timeout=60000)
                page.wait_for_timeout(4000)
                # ads embedded in initial HTML
                for sc in page.eval_on_selector_all("script[type='application/json']", "e=>e.map(x=>x.textContent)"):
                    try: find_ads(json.loads(sc), bucket)
                    except Exception: pass
                for _ in range(scrolls):
                    page.mouse.wheel(0, 6000)
                    page.wait_for_timeout(random.randint(1500, 2800))
            except Exception as e:
                print("ERR", kw, e); continue
            n = 0
            with open(OUT, "a", encoding="utf-8") as f:
                for ad in bucket:
                    row = norm(ad, kw)
                    key = (row["ad_id"], kw)
                    if key in seen: continue
                    seen.add(key); n += 1
                    f.write(json.dumps(row, ensure_ascii=False) + "\n")
            with open(DONE, "a", encoding="utf-8") as f:
                f.write(kw + "\n")
            print(f"{kw!r}: +{n} ads")
            time.sleep(random.uniform(3, 7))
        b.close()


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--scrolls", type=int, default=12)
    ap.add_argument("--show", action="store_true", help="show browser window")
    a = ap.parse_args()
    run(a.limit, a.scrolls, not a.show)

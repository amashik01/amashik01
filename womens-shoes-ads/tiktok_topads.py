"""TikTok Creative Center Top Ads, Bangladesh -> data/tiktok_topads.jsonl
Drives the public search box (no login) and records the page's own list responses.
  python tiktok_topads.py [--period 180] [--limit 0]
"""
import argparse, json, os, sys
from urllib.parse import quote
from playwright.sync_api import sync_playwright
from tt_common import launch, close_popups

OUT = "data/tiktok_topads.jsonl"
URL = "https://ads.tiktok.com/business/creativecenter/inspiration/topads/pc/en?region=BD&period={p}"
KEYWORDS = ["shoes", "shoe", "heel", "high heel", "sandal", "slipper", "footwear", "flat", "loafer", "boot",
            "ladies", "women", "juta", "jutar", "fashion", "nagra", "kolhapuri", "ladies shoes", "women shoes",
            "block heel", "wedge", "pump", "bridal", "party", "leather", "comfort", "pinkwear", "nawabi",
            "ziziko", "beshati", "hermizon", "merkis", "city shop", "জুতা", "মেয়েদের জুতা", "স্যান্ডেল", "হিল",
            "ফুটওয়্যার", "cod", "cash on delivery"]


def main(period, limit, headless=True):
    os.makedirs("data", exist_ok=True)
    seen, rows = {}, []
    if os.path.exists(OUT):
        for l in open(OUT, encoding="utf-8"):
            r = json.loads(l); seen[r["id"]] = r
    kws = KEYWORDS[:limit or None]
    with sync_playwright() as p:
        b, pg = launch(p, headless)
        bucket = []

        def on(r):
            if "top_ads/v2/list" in r.url:
                try:
                    bucket.append(r.json()["data"])
                except Exception:
                    pass
        pg.on("response", on)
        try: pg.goto(URL.format(p=period), wait_until="commit", timeout=60000)
        except Exception as e: print("goto", str(e)[:80])
        pg.wait_for_timeout(8000); close_popups(pg); pg.wait_for_timeout(1000)
        for kw in kws:
            bucket.clear()
            try:
                inp = next(i for i in pg.query_selector_all("input[placeholder*='earch' i]") if i.is_visible())
                inp.click(force=True, timeout=5000); inp.fill(kw, timeout=8000)
                pg.keyboard.press("Enter"); pg.wait_for_timeout(5000)
                for _ in range(15):          # "View more" pagination
                    if not bucket or not bucket[-1]["pagination"].get("has_more"): break
                    btn = pg.query_selector("text=View More") or pg.query_selector("text=View more")
                    if not btn: break
                    btn.click(force=True); pg.wait_for_timeout(3500)
            except Exception as e:
                print("ERR", kw, str(e)[:100]); close_popups(pg); continue
            total = bucket[-1]["pagination"].get("total_count") if bucket else None
            new = 0
            for d in bucket:
                for m in d.get("materials") or []:
                    mid = str(m["id"])
                    if mid not in seen:
                        m["_kw"] = kw; m["_period"] = period; seen[mid] = m; new += 1
                    else:
                        seen[mid].setdefault("_kws", []).append(kw)
            print(f"{kw!r}: total_count={total} new={new} all={len(seen)}", flush=True)
            with open(OUT, "w", encoding="utf-8") as f:
                for m in seen.values(): f.write(json.dumps(m, ensure_ascii=False) + "\n")
        b.close()
    with open(OUT, "w", encoding="utf-8") as f:
        for m in seen.values(): f.write(json.dumps(m, ensure_ascii=False) + "\n")


if __name__ == "__main__":
    ap = argparse.ArgumentParser(); ap.add_argument("--period", type=int, default=180)
    ap.add_argument("--limit", type=int, default=0); ap.add_argument("--show", action="store_true")
    a = ap.parse_args(); main(a.period, a.limit, not a.show)

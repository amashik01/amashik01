"""Instagram ads in Meta Ad Library (Bangladesh, active, platform=Instagram) -> data/ig_ads.jsonl
Keywords come from pinkwear_keywords.xlsx (162 BN/EN/Banglish terms). Resumable, saves after every keyword.
  python ig_scraper.py [--limit N] [--scrolls 10] [--show]
"""
import argparse, json, os, re, time, random
from urllib.parse import quote
import openpyxl
from playwright.sync_api import sync_playwright
from tt_common import launch

OUT, DONE = "data/ig_ads.jsonl", "data/ig_done.txt"
URL = ("https://www.facebook.com/ads/library/?active_status=active&ad_type=all&country=BD&media_type=all"
       "&publisher_platforms[0]=instagram&sort_data[mode]=total_impressions&sort_data[direction]=desc"
       "&search_type=keyword_unordered&q={q}")
HDR = re.compile(r"Library ID:\s*(\d+)")


def keywords():
    ws = openpyxl.load_workbook("pinkwear_keywords.xlsx", read_only=True)["Keywords"]
    return [r[3] for r in list(ws.iter_rows(values_only=True))[1:] if r[3]]


def parse(text, kw):
    """Split the results page text into ads. Header block (Library ID) precedes its creative."""
    out, ms = [], list(HDR.finditer(text))
    for i, m in enumerate(ms):
        end = ms[i + 1].start() if i + 1 < len(ms) else len(text)
        seg = text[m.end():end]
        seg = re.sub(r"\n(Active|Inactive)\s*$", "", seg.rstrip())          # next ad's status line
        start = re.search(r"Started running on ([A-Za-z]+ \d+, \d{4})", seg)
        after = seg.split("See ad details", 1)[-1] if "See ad details" in seg else seg
        lines = [l.strip() for l in after.split("\n") if l.strip() and l.strip() != "​"]
        page = lines[0] if lines else ""
        copy_lines, dom, ig = [], "", ""
        for j, l in enumerate(lines[1:], 1):
            if l == "Sponsored": continue
            if l == "Visit Instagram profile" and j >= 1: ig = lines[j - 1]; continue
            copy_lines.append(l)
        multi = bool(re.search(r"multiple versions|This ad has \d+ versions", seg, re.I))
        out.append({"ad_id": m.group(1), "ad_url": f"https://www.facebook.com/ads/library/?id={m.group(1)}",
                    "start": start.group(1) if start else "", "page": page, "instagram_handle": ig,
                    "copy": "\n".join(copy_lines)[:1500], "multi_version": multi, "keyword": kw})
    return out


def main(limit, scrolls, headless):
    os.makedirs("data", exist_ok=True)
    done = set(open(DONE, encoding="utf-8").read().split("\n")) if os.path.exists(DONE) else set()
    seen = {}
    if os.path.exists(OUT):
        for l in open(OUT, encoding="utf-8"):
            r = json.loads(l); seen[r["ad_id"]] = r
    kws = [k for k in keywords() if k not in done][: limit or None]
    with sync_playwright() as p:
        b, pg = launch(p, headless)
        for kw in kws:
            try:
                pg.goto(URL.format(q=quote(kw)), wait_until="domcontentloaded", timeout=60000)
                pg.wait_for_timeout(6000)
                last = -1
                for _ in range(scrolls):
                    pg.mouse.wheel(0, 8000); pg.wait_for_timeout(2500)
                    n = pg.inner_text("body").count("Library ID")
                    if n == last: break
                    last = n
                text = pg.inner_text("body")
            except Exception as e:
                print("ERR", kw, str(e)[:80], flush=True); continue
            m = re.search(r"~?([\d,]+) results?", text)
            ads = parse(text, kw); new = 0
            for a in ads:
                if a["ad_id"] in seen: seen[a["ad_id"]].setdefault("keywords", []).append(kw)
                else: a["keywords"] = [kw]; seen[a["ad_id"]] = a; new += 1
            with open(OUT, "w", encoding="utf-8") as f:
                for a in seen.values(): f.write(json.dumps(a, ensure_ascii=False) + "\n")
            with open(DONE, "a", encoding="utf-8") as f: f.write(kw + "\n")
            print(f"{kw!r}: results={m.group(1) if m else '?'} parsed={len(ads)} new={new} total={len(seen)}", flush=True)
            time.sleep(random.uniform(2, 5))
        b.close()


if __name__ == "__main__":
    ap = argparse.ArgumentParser(); ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--scrolls", type=int, default=10); ap.add_argument("--show", action="store_true")
    a = ap.parse_args(); main(a.limit, a.scrolls, not a.show)

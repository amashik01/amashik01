# Women's shoes ad research (Bangladesh)

Pipeline: `keywords.py` (294 BN/EN/Banglish keywords) -> `fb_scraper.py` / `tiktok_scraper.py` -> `build_xlsx.py` -> two .xlsx files.

```
pip install playwright openpyxl && playwright install chromium
python fb_scraper.py --limit 20     # pilot
python fb_scraper.py                # full, resumable
python tiktok_scraper.py
python build_xlsx.py
```
Output links: permanent Ad Library / TikTok detail URLs are the main links; CDN image/video URLs are secondary and can expire.
Needs network access to facebook.com, *.fbcdn.net, ads.tiktok.com.

"""Shared Playwright helpers for TikTok scraping (works behind the sandbox proxy and on a normal PC)."""
import os
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
      "Chrome/130.0.0.0 Safari/537.36")
SANDBOX_CHROME = "/opt/pw-browsers/chromium"


def launch(p, headless=True):
    kw = dict(headless=headless, args=["--disable-blink-features=AutomationControlled"])
    if os.path.exists(SANDBOX_CHROME):
        kw["executable_path"] = SANDBOX_CHROME
    b = p.chromium.launch(**kw)
    ctx = b.new_context(user_agent=UA, locale="en-US", viewport={"width": 1400, "height": 1000})
    return b, ctx.new_page()


def close_popups(pg):
    for sel in ["[class*=Popup] [class*=close]", ".byted-modal-close-icon", "button[aria-label=Close]"]:
        try:
            pg.click(sel, timeout=1500)
        except Exception:
            pass
    pg.keyboard.press("Escape")

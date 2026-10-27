"""
One-time login script.
Run this ONCE to log into Facebook/Instagram.
Session is saved — you never need to run this again.

Usage:
  py login.py facebook
  py login.py instagram
"""

import sys
from pathlib import Path
from playwright.sync_api import sync_playwright

SESSION_DIR = Path(__file__).parent / ".browser_session"
SESSION_DIR.mkdir(exist_ok=True)

platform = sys.argv[1].lower() if len(sys.argv) > 1 else "facebook"
url = "https://www.instagram.com/accounts/login/" if platform == "instagram" else "https://www.facebook.com/login"

print(f"\nOpening {platform} login...")
print("Log in fully (including any verification codes), then close the browser window.")
print("Your session will be saved automatically.\n")

with sync_playwright() as p:
    context = p.chromium.launch_persistent_context(
        user_data_dir=str(SESSION_DIR),
        headless=False,
        args=["--disable-blink-features=AutomationControlled"],
    )
    page = context.pages[0] if context.pages else context.new_page()
    page.goto(url, wait_until="domcontentloaded", timeout=60000)

    # Stay open until user closes the browser window themselves
    print("Browser is open. Close it when you are fully logged in.\n")
    try:
        context.wait_for_event("close", timeout=600000)  # wait up to 10 min
    except Exception:
        pass

print("\nSession saved. You can now use the DM app normally.")

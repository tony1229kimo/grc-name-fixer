"""
版本資訊 — 給 launcher banner 顯示 + /api/version 比對 GitHub Releases 最新版用。

發新版步驟:
  1. 更新此處 VERSION
  2. git commit + git push
  3. .\build.ps1 -Release  (打包 + zip + gh release create)

GitHub Releases tag 命名: v{VERSION}  例如 v1.1.0
"""

VERSION = "1.1.0"

GITHUB_OWNER = "tony1229kimo"
GITHUB_REPO = "grc-name-fixer"
RELEASES_API_URL = (
    f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest"
)
RELEASES_PAGE_URL = (
    f"https://github.com/{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest"
)

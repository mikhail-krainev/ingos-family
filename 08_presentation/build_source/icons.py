"""Rasterize lucide icons to transparent PNGs for the deck: img/icons/<name>_<color>.png.
Usage: python3 icons.py"""
import pathlib, subprocess
from PIL import Image

HERE = pathlib.Path(__file__).resolve().parent
SRC = HERE / "node_modules" / "lucide-static" / "icons"
OUT = HERE / "img" / "icons"
CHROME = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
NAMES = ["user", "users", "wallet", "zap", "phone", "phone-call", "heart-handshake", "stethoscope", "building-2", "video",
         "bot", "file-text", "briefcase", "shield-check", "smartphone", "house", "bed-double", "clock", "bell", "medal",
         "badge-check", "id-card", "headset", "brain", "list-checks", "send", "hospital", "pill", "hand-coins",
         "calendar-check", "message-circle", "landmark", "ban", "circle-alert", "route", "user-round"]
COLORS = {"blue": "#1E3FD8", "white": "#FFFFFF", "navy": "#0A2A8A", "lime": "#B8F02A"}
CELL, SCALE, COLS = 96, 3, 12

cells = [(n, c) for n in NAMES for c in COLORS]
rows = (len(cells) + COLS - 1) // COLS
html = ["<html><body style='margin:0;background:transparent'><div style='display:grid;grid-template-columns:repeat(%d,%dpx)'>" % (COLS, CELL)]
for n, c in cells:
    svg = (SRC / f"{n}.svg").read_text().replace('width="24"', f'width="{CELL}"').replace('height="24"', f'height="{CELL}"')
    html.append(f"<div style='width:{CELL}px;height:{CELL}px;color:{COLORS[c]};line-height:0'>{svg}</div>")
html.append("</div></body></html>")
page = HERE / "icons_tmp.html"
page.write_text("".join(html), encoding="utf-8")
png = HERE / "icons_tmp.png"
subprocess.run([CHROME, "--headless", "--disable-gpu", "--hide-scrollbars", f"--force-device-scale-factor={SCALE}",
                "--default-background-color=00000000", f"--window-size={COLS * CELL},{rows * CELL}",
                f"--screenshot={png}", page.as_uri()], check=True, capture_output=True)
sheet = Image.open(png).convert("RGBA")
OUT.mkdir(parents=True, exist_ok=True)
px = CELL * SCALE
for i, (n, c) in enumerate(cells):
    x, y = (i % COLS) * px, (i // COLS) * px
    sheet.crop((x, y, x + px, y + px)).save(OUT / f"{n}_{c}.png")
page.unlink()
png.unlink()
print(len(cells), "icons")

"""Render prototype screens to PNG for the deck: iPhone frame, desktop browser, coordinator.
Usage: python3 shot.py  (run from this folder; writes ../img/*.png)"""
import pathlib, subprocess, sys
from PIL import Image

HERE = pathlib.Path(__file__).resolve().parent
SRC = HERE.parents[2] / "06_prototype" / "Прототип_ИНГОССЕМЬЯ.html"
OUT = HERE.parent / "img"
CHROME = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"

ISOLATE = ("var el=document.querySelector('{sel}');var w=document.createElement('div');w.style.cssText='padding:40px;display:inline-block';"
           "w.appendChild(el);document.body.replaceChildren(w);document.body.style.cssText='margin:0;background:transparent';"
           "document.documentElement.style.background='transparent';")
PHONE = ISOLATE.format(sel=".iphone")
DESK = "document.querySelector('[data-device=desk]').click();" + ISOLATE.format(sel=".browser")
COORD = "document.getElementById('tab-coord').click();" + ISOLATE.format(sel=".coord")

def jump(n): return f"document.querySelector('[data-jump={n}]').click();"

VIEWS = {  # name: (js, window w, h, scale)
    "ph_welcome": (jump("welcome") + PHONE, 900, 1000, 3),
    "ph_feed": (jump("feed") + PHONE, 900, 1000, 3),
    "ph_event": (jump("feed") + "document.querySelector('[data-event=\"3\"]').click();" + PHONE, 900, 1000, 3),
    "ph_summary": (jump("summary") + PHONE, 900, 1000, 3),
    "ph_tariff": (jump("tariff") + "document.querySelector('[data-benefit=combat]').click();" + PHONE, 900, 1000, 3),
    "ph_chat": (jump("chat") + PHONE, 900, 1000, 3),
    "desk_feed": (jump("feed") + "document.querySelector('[data-event=\"3\"]').click();" + DESK, 1280, 900, 2),
    "coord": (COORD, 1400, 1100, 2),
}

def shoot(name, js, w, h, scale):
    html = HERE / f"{name}.html"
    src = SRC.read_text(encoding="utf-8")
    hook = "<script>window.addEventListener('load',function(){setTimeout(function(){" + js + "},50)});</script></body>"
    html.write_text(src.replace("</body>", hook), encoding="utf-8")
    png = HERE / f"{name}.png"
    subprocess.run([CHROME, "--headless", "--disable-gpu", "--hide-scrollbars", f"--force-device-scale-factor={scale}",
                    "--default-background-color=00000000", f"--window-size={w},{h}", "--virtual-time-budget=3000",
                    f"--screenshot={png}", html.as_uri()], check=True, capture_output=True)
    im = Image.open(png).convert("RGBA")
    bbox = im.getchannel("A").getbbox()
    if bbox is None:
        sys.exit(f"{name}: empty screenshot")
    l, t, r, b = bbox
    im = im.crop((max(l - 2, 0), max(t - 2, 0), min(r + 2, im.width), min(b + 2, im.height)))
    im.save(OUT / f"{name}.png")
    html.unlink()
    png.unlink()
    print(name, im.size)

for n, v in VIEWS.items():
    shoot(n, *v)

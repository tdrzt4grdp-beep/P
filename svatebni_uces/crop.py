# Ořízne fotky na poměr stran rámečků v prezentaci
import json, os
from PIL import Image
D = os.path.dirname(os.path.abspath(__file__))
specs = {  # nazev: (sirka, vyska v palcich, fokus x, fokus y)
  "kytice":       (4.5, 5.625, 0.5, 0.45),
  "saty":         (2.0, 3.7, 0.5, 0.4),
  "titul":        (2.0, 3.7, 0.45, 0.5),
  "vlasy":        (4.0, 3.7, 0.45, 0.5),
  "navrh_zezadu": (2.6, 3.3, 0.5, 0.5),
  "navrh_zboku":  (2.6, 3.3, 0.45, 0.5),
  "kulma":        (2.2, 1.5, 0.6, 0.5),
  "vlasenky":     (2.2, 1.5, 0.5, 0.5),
  "cop":          (2.6, 3.7, 0.5, 0.5),
  "spona":        (2.85, 1.55, 0.5, 0.5),
  "kvetiny":      (2.85, 1.55, 0.45, 0.4),
  "zavoj":        (2.85, 1.55, 0.6, 0.35),
  "vysledek":     (2.95, 3.35, 0.5, 0.5),
  "dekuji":       (4.5, 5.625, 0.45, 0.5),
}
src = {"vysledek": "navrh_zezadu", "dekuji": "navrh_zboku"}
os.makedirs(f"{D}/foto/orez", exist_ok=True)
for n, (w, h, fx, fy) in specs.items():
    im = Image.open(f"{D}/foto/{src.get(n, n)}.jpg").convert("RGB")
    W, H = im.size; r = w / h
    if W / H > r: cw, ch = int(H * r), H
    else: cw, ch = W, int(W / r)
    x = min(max(int(fx * W - cw / 2), 0), W - cw); y = min(max(int(fy * H - ch / 2), 0), H - ch)
    im.crop((x, y, x + cw, y + ch)).save(f"{D}/foto/orez/{n}.jpg", quality=92)
    print(n, cw, ch)

#!/bin/bash
# Sestaví prezentaci ve dvou krocích:
# 1) grafická vrstva (kartičky, rámečky, ikonky, fotky) -> obrázky pozadí
# 2) finální prezentace = obrázek pozadí + upravitelný text
set -e
cd "$(dirname "$0")"
SKILL=../.claude/skills/pptx/scripts
mkdir -p pozadi
node build.js visual
python3 "$SKILL/office/soffice.py" --headless --convert-to pdf --outdir pozadi pozadi/graficka_vrstva.pptx >/dev/null 2>&1
rm -f pozadi/slide-*.png
pdftoppm -png -r 200 pozadi/graficka_vrstva.pdf pozadi/slide
for f in pozadi/slide-0*.png; do [ -e "$f" ] && mv "$f" "pozadi/slide-$(basename "$f" .png | sed 's/slide-0*//').png"; done
node build.js final
zip -dq Svatebni_uces_Habanova.pptx '*/' 2>/dev/null || true
python3 "$SKILL/office/validate.py" Svatebni_uces_Habanova.pptx

#!/bin/bash
# בונה סגנון, ממיר ל-PDF ולתמונות. שימוש: ./render.sh night
set -e
cd "$(dirname "$0")"
STYLE=${1:-night}
SK=/root/.claude/skills/synced/08c638a7-ddb8-46c5-be86-9b61f06d3064_3c3b1d69-4ee7-425a-92f6-908dec8fa724/pptx
python3 build.py "$STYLE"
NAME=superman; [ "$STYLE" != night ] && NAME=superman_$STYLE
python3 $SK/scripts/office/validate.py $NAME.pptx
python3 $SK/scripts/office/soffice.py --headless --convert-to pdf $NAME.pptx >/dev/null 2>&1
mkdir -p previews/$STYLE
find previews/$STYLE -name '*.png' -delete
pdftoppm -png -r 80 $NAME.pdf previews/$STYLE/slide

# Source this in every shell: source /home/user/lakehouse-tracker/BRAND_US_Launch/build/env.sh
export ROOT=/home/user/lakehouse-tracker/BRAND_US_Launch
export OUT=$ROOT/output
export BUILD=$ROOT/build
export NODE_PATH=$(npm root -g):${NODE_PATH:-}
export PUPPETEER_SKIP_DOWNLOAD=1
export MMDC_CFG=$BUILD/puppeteer.json
export SKILLS=/mnt/skills/public
export RECALC="python3 $SKILLS/xlsx/scripts/recalc.py"            # usage: $RECALC file.xlsx 120
export SOFFICE="python3 $SKILLS/pptx/scripts/office/soffice.py"  # usage: $SOFFICE --headless --convert-to pdf --outdir DIR file
export VALIDATE="python3 $SKILLS/pptx/scripts/office/validate.py" # usage: $VALIDATE deck.pptx
export APPLY_THEME=$SKILLS/pptx/scripts/apply_theme.js
export KO_FONT="Noto Sans KR"
# Render a mermaid file to PNG and PDF: render_mmd input.mmd
render_mmd() { local f="$1"; local b="${f%.mmd}"; mmdc -q -p "$MMDC_CFG" -c "$BUILD/mermaid_config.json" -i "$f" -o "$b.png" -b white --size 2400 -s 1 && mmdc -q -p "$MMDC_CFG" -c "$BUILD/mermaid_config.json" -i "$f" -o "$b.pdf" -b white --size 2400 -s 1; }
# Convert any office file to PDF next to it: to_pdf file.pptx
to_pdf() { local f="$1"; $SOFFICE --headless --convert-to pdf --outdir "$(dirname "$f")" "$f"; }
# Render a PDF to JPG pages at 110 dpi into a directory: to_jpg file.pdf outdir
to_jpg() { mkdir -p "$2"; rm -f "$2"/page-*.jpg; pdftoppm -jpeg -r 110 "$1" "$2/page"; ls -1 "$2"/page-*.jpg; }

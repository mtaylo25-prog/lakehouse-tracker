#!/usr/bin/env bash
# Build the Korean deck: pptx from build_deck.js, validate, PDF through LibreOffice UNO with Asian/Latin autospacing
# off (build/lo_pdf_ko.py; the plain to_pdf path leaves gaps between Hangul and Latin text), then QA renders and lint.
set -euo pipefail
source /home/user/lakehouse-tracker/BRAND_US_Launch/build/env.sh
cd "$ROOT"
python3 build/check_strings.py
node build/build_deck.js --lang ko
$VALIDATE output/02_Business_Proposal_KO.pptx
python3 build/lo_pdf_ko.py output/02_Business_Proposal_KO.pptx output/02_Business_Proposal_KO.pdf
to_jpg output/02_Business_Proposal_KO.pdf build/qa/deck_ko >/dev/null
python3 build/lint_hyphens.py output/02_Business_Proposal_KO.pptx output/02_Business_Proposal_KO.pdf

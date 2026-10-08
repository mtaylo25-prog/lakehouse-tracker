# BRAND US Launch package

Build spec: BUILD_BRIEF.md (what) and build/BUILD_SPEC.md (how). Decisions and defaults: inputs/DECISIONS.md.
Standing rules: no hyphens or dashes in any delivered text (official form and registration numbers keep their format); the brand is the placeholder BRAND; every number in a document comes from build/figures.json which is pulled from output/03_Financial_Model_5Y.xlsx; label every assumption Evidence based or Management estimate; never invent certifications, reviews, quotes or market sizes.
Shell setup: source build/env.sh. Lint: python3 build/lint_hyphens.py output/. Tie out: python3 build/tie_out.py.

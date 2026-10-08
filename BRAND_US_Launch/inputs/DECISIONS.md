# Open decisions and the defaults used in this build

Edit any line before running. The build reads this file instead of stopping to ask.

| # | Decision | Default used | Source |
|---|---|---|---|
| 1 | US brand name | BRAND (placeholder in every file) | Not chosen yet. Korean deck used MBF. |
| 2 | Ownership of the US LLC | Lucy (Hyejin Yoo) owns 100%. Mark Taylor paid under a written consulting or commission agreement, no equity. | Mark, October 2026 |
| 3 | Funding source | Self funded by Lucy to $39K (August ask). Any need above $39K is shown as a dated line of credit or investor amount in the model. | August 2026 deck |
| 4 | FX rate | 1,430 KRW per USD (input cell; scenarios at 1,300 and 1,550) | Mark, October 2026 |
| 5 | Duty scenario for the base case | Copper rate (HTS 8544.42.90: 2.6% MFN + 25% Section 301 List 3 + 12.5% Section 301 forced labor = 40.1%). Fiber rate and a +10 point stress case are modeled alongside. | Brief section 1, pending the CBP binding ruling |

Additional defaults the build uses when no input file answers them:

| Item | Default | Why |
|---|---|---|
| H2 Company legal form | Korean sole proprietorship (registration number middle digits 64) | Brief. Flagged for the 세무사. |
| Mark's role | Advisor and US operations support under a consulting agreement. Not an owner, not an agent who signs contracts. | Keeps the treaty position clean. |
| Entity state | Wyoming single member LLC | Brief |
| First SKUs | 5 hero SKUs: 8K 7M, 8K 10M, 8K 15M, 8K 20M, 4K 30M (management estimate from the August unit mix) | 8K short runs were 53% of units and 54% of gross profit |
| Launch month | Month 1 = first month with inventory live on Amazon, assumed early 2027 | Build sequence needs 8 to 10 weeks before first sale |

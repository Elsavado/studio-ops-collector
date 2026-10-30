# studio-ops-collector

GitHub Actions collector of operational research data for a **game development & fintech studio** (mobile/PC/web
games, live ops, payments, financial apps), across every field such a company works in: engineering, game
development, fintech, accounting, legal, HR, product, sales/marketing/support and security. Files go straight
from the runner to Google Drive with rclone (very large files stream directly from the source to Drive).

Everything is filed by topic under `Studio Operations Research/Public Data/`:

| Folder | Main sources |
|---|---|
| 01 Game Development | engine docs (Godot, Bevy, Defold, Phaser, raylib), game datasets, telemetry & player analytics |
| 02 Fintech & Payments | CFPB consumer complaints, open-banking specs, Apache Fineract, payment & fraud datasets |
| 03 Accounting & Finance | SEC Financial Statement Data Sets (every quarter since 2009), filings of listed game & fintech companies |
| 04 Legal & Compliance | real contracts filed as SEC exhibits (17 contract types), CUAD contracts, eCFR titles, case law |
| 05 HR & Workforce | O*NET occupations, salary / attrition / hiring datasets, public payrolls |
| 06 Software Engineering & DevOps | MDN, React, TypeScript, Kotlin, Kubernetes, Docker, PostgreSQL docs, handbooks, product lifecycles |
| 07 Product, Design & UX · 08 Sales, Marketing & Customer Support · 09 Security & Privacy | datasets by topic |
| 10 Books & Knowledge Library | open-access books, Wikipedia articles as DOCX, OWASP, CWE, licence texts |
| 11 GitHub Repositories · 12 Jupyter Notebooks · 13 Software & Apps | licensed repos, notebooks, app releases |
| 14 Q&A Archives | Stack Exchange dumps incl. Stack Overflow, Game Development, Law, Workplace, Money, Quant |
| 15 Videos & Talks | openly licensed talks and lectures (Internet Archive) |

Only data files are uploaded; each folder's source and licence are recorded in `_state`.

## Running

- Secret `RCLONE_CONF`: rclone.conf with a Google Drive remote named `[pax]`. Optional: `KAGGLE_API_TOKEN`, `SEC_USER_AGENT`.
- Variable `LIVE=true` turns on uploads (otherwise dry runs). `COMPANY_LIVE=true` adds the fictional studio (phase 2).
- Every 6 hours; each source gets a fair share of the run; state lives on Drive in `_state/`.

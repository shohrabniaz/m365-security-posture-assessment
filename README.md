# Microsoft 365 security posture assessment for SMEs

Fictional case: Coastal Managed IT (~60 staff modelled). Assessment uses public Microsoft and Australian guidance plus a **synthetic tenant baseline**. Live Microsoft Graph collection is optional and read-only; this pack does not require paid Secure Score licences or a production tenant.

## Research questions

1. Which M365 controls should an SME check first (MFA, admins, guests, audit, forwarding, risky sign-in)?
2. How can findings be scored into clear risk categories without inventing portal UI?
3. What remediation and governance steps are safe on Entra ID Free / Microsoft 365 Business-class estates?

## Deliverables

- Control framework and risk categories (`01_Framework`)
- Excel posture dashboard / scorecard (`02_Evaluation`)
- PowerShell collectors (live Graph when available) and a synthetic baseline (`scripts`, `evidence`)
- Report and 5-minute slides (`Final_Submission`)

## Scripts

1. `01_Assess_Identity_MFA.ps1` — MFA registration / per-user MFA notes (Graph or synthetic).
2. `02_Assess_Admins_Guests.ps1` — directory roles and guest inventory.
3. `03_Assess_Audit_Forwarding_SignIn.ps1` — audit/forwarding/risky-sign-in checklist fields.
4. `04_Build_Posture_Report.ps1` — merges findings into scored JSON for the workbook.

Default mode writes **synthetic** evidence so the pack remains reproducible without Graph. Pass `-Live` only after `Connect-MgGraph` to the laboratory tenant.

## Limits

Synthetic findings are labelled. Live Graph was skipped when directory access was unavailable. No production tenant. No invented Secure Score screenshots.

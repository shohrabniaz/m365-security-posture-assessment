"""Build Academic Project 15 — M365 Security Posture Assessment submission pack."""
from __future__ import annotations

from datetime import date
from pathlib import Path
import json
import shutil

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, RGBColor
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from pptx import Presentation
from pptx.dml.color import RGBColor as PRgb
from pptx.util import Inches, Pt as PPt

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "Final_Submission"
FRAME = ROOT / "01_Framework"
EVAL = ROOT / "02_Evaluation"
EVID = ROOT / "evidence"
NAVY = RGBColor(0x1F, 0x4E, 0x79)
TODAY = date.today().isoformat()


def heading(doc, text, level=1):
    h = doc.add_heading(text, level=level)
    for r in h.runs:
        r.font.color.rgb = NAVY


def load_summary():
    files = sorted(EVID.glob("*_Posture_Summary.json"), key=lambda p: p.stat().st_mtime, reverse=True)
    if not files:
        return None
    return json.loads(files[0].read_text(encoding="utf-8-sig"))


def build_framework():
    FRAME.mkdir(parents=True, exist_ok=True)

    d = Document()
    d.add_heading("M365 security posture assessment framework", 0)
    d.add_paragraph(f"Md Shohrab Hossen Niaz | Academic Project 15 | {TODAY}")
    heading(d, "Purpose", 1)
    d.add_paragraph(
        "Give an SME a first-pass check of Microsoft 365 identity and mailbox-adjacent "
        "controls using public Microsoft and Australian guidance. The framework is "
        "designed for Entra ID Free / Business-class estates and refuses to invent "
        "paid portal features."
    )
    heading(d, "Control families", 1)
    for x in [
        "MFA / authentication strength (privileged and standard users; legacy auth).",
        "Admin accounts (standing privilege, Global Admin count, orphan leavers).",
        "Guest access (inventory, invitation settings, review cadence).",
        "Audit logging (unified audit enabled; retention and review process).",
        "Mailbox forwarding (external SMTP / inbox rules detection process).",
        "Risky sign-in (Identity Protection if licensed; otherwise NotAvailable + sign-in log review).",
    ]:
        d.add_paragraph(x, style="List Number")
    heading(d, "Scoring", 1)
    d.add_paragraph(
        "Each control has Score / MaxScore (0–100). Aggregate posture percent = "
        "sum(Score) / sum(MaxScore). Risk bands: Critical <30%, Elevated 30–54%, "
        "Moderate 55–79%, Low residual ≥80%. Severity labels drive remediation order."
    )
    heading(d, "Evidence rules", 1)
    d.add_paragraph(
        "Synthetic mode is explicit in JSON. Live Graph mode must match the laboratory "
        "tenant ID and must not claim Conditional Access, PIM, or Identity Protection "
        "on Free tier. Screenshots must be real portal captures — none are fabricated here."
    )
    d.save(FRAME / "Posture_Assessment_Framework.docx")

    d = Document()
    d.add_heading("Risk categories and SME remediation guide", 0)
    d.add_paragraph(f"Md Shohrab Hossen Niaz | Academic Project 15 | {TODAY}")
    heading(d, "Categories", 1)
    rows = [
        ("Critical", "Disabled or unknown accounts in privileged groups; unrestricted Global Admin sprawl."),
        ("High", "No MFA on admins/users; no forwarding scan; Identity Protection NotAvailable without compensating review."),
        ("Medium", "Guest invite open; audit log on but no review process; weak sign-in review cadence."),
        ("Low", "Documented exceptions with owner, expiry, and compensating control."),
    ]
    for name, desc in rows:
        d.add_paragraph(f"{name}: {desc}", style="List Bullet")
    heading(d, "Safe remediation (Free-tier aware)", 1)
    for x in [
        "Remove leavers from privileged groups within 24 hours; weekly orphan report.",
        "Enable MFA for all users capable of MFA on the licence; separate admin UPNs.",
        "Limit guest invitations to admins; monthly guest export and disable unused guests.",
        "Confirm mailbox audit / unified audit; schedule a monthly sign-in anomaly review.",
        "Alert on new external forwarding; block auto-forwarding via transport rule where licensed.",
        "Do not purchase language: state NotAvailable for CA/ID Protection until budget exists.",
    ]:
        d.add_paragraph(x, style="List Number")
    heading(d, "Governance controls", 1)
    d.add_paragraph(
        "Name an identity owner. Keep a control register (this project's workbook). "
        "Review Critical/High monthly. Link joiner-mover-leaver and JIT privilege work "
        "so posture findings feed operational runbooks."
    )
    d.save(FRAME / "Risk_Categories_Remediation_Governance.docx")


def build_workbook(summary):
    EVAL.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    fill = PatternFill("solid", fgColor="1F4E79")
    white = Font(color="FFFFFF", bold=True)
    crit = PatternFill("solid", fgColor="C00000")
    high = PatternFill("solid", fgColor="ED7D31")
    med = PatternFill("solid", fgColor="FFC000")

    dash = wb.active
    dash.title = "Dashboard"
    dash["A1"] = "M365 posture scorecard — Coastal Managed IT (fictional)"
    dash["A1"].font = Font(bold=True, color="1F4E79", size=14)
    dash.merge_cells("A1:F1")
    dash["A3"] = "Posture %"
    dash["B3"] = summary["PosturePercent"] if summary else 0
    dash["A4"] = "Risk band"
    dash["B4"] = summary["RiskBand"] if summary else "n/a"
    dash["A5"] = "Mode"
    dash["B5"] = summary.get("Mode", "Synthetic") if summary else "Synthetic"
    dash["A6"] = "Generated (UTC)"
    dash["B6"] = summary.get("GeneratedUtc", "") if summary else ""
    dash["A8"] = "Severity"
    dash["B8"] = "Count"
    for c in (dash["A8"], dash["B8"]):
        c.fill = fill
        c.font = white
    sev = summary.get("SeverityCounts", {}) if summary else {}
    for i, name in enumerate(["Critical", "High", "Medium", "Low"], start=9):
        dash[f"A{i}"] = name
        dash[f"B{i}"] = sev.get(name, 0)
    dash["A14"] = "Top remediation themes"
    dash["A14"].font = Font(bold=True)
    rem = summary.get("RemediationTop5", []) if summary else []
    for i, line in enumerate(rem, start=15):
        dash[f"A{i}"] = line
        dash.merge_cells(f"A{i}:F{i}")
    for col in "ABCDEF":
        dash.column_dimensions[col].width = 18
    dash.column_dimensions["A"].width = 56

    chart = BarChart()
    chart.title = "Findings by severity"
    chart.y_axis.title = "Count"
    data = Reference(dash, min_col=2, min_row=8, max_row=12)
    cats = Reference(dash, min_col=1, min_row=9, max_row=12)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.shape = 4
    dash.add_chart(chart, "D3")

    find = wb.create_sheet("Findings")
    headers = ["Id", "Family", "Title", "Severity", "Score", "MaxScore", "Observed", "Mode"]
    find.append(headers)
    for c in find[1]:
        c.fill = fill
        c.font = white
    findings = summary.get("Findings", []) if summary else []
    for item in findings:
        # findings may be dicts from JSON
        if isinstance(item, dict):
            row = [
                item.get("Id"),
                item.get("Family"),
                item.get("Title"),
                item.get("Severity"),
                item.get("Score"),
                item.get("MaxScore"),
                item.get("Observed"),
                item.get("Mode"),
            ]
        else:
            row = [getattr(item, k, None) for k in ["Id", "Family", "Title", "Severity", "Score", "MaxScore", "Observed", "Mode"]]
        find.append(row)
        sev_name = row[3]
        fill_map = {"Critical": crit, "High": high, "Medium": med}
        if sev_name in fill_map:
            find.cell(row=find.max_row, column=4).fill = fill_map[sev_name]
    for col in find.columns:
        find.column_dimensions[col[0].column_letter].width = 22
    find.column_dimensions["G"].width = 70

    guide = wb.create_sheet("Control_Catalogue")
    guide.append(["Control ID", "Family", "Public guidance anchor", "Free-tier note", "Remediation hint"])
    for c in guide[1]:
        c.fill = fill
        c.font = white
    catalogue = [
        ("MFA-01", "MFA", "Require MFA for admins", "CA may be unavailable", "Per-user MFA / security defaults where applicable"),
        ("MFA-02", "MFA", "Require MFA for users", "Licence dependent", "Enable for all cloud users"),
        ("MFA-03", "MFA", "Block legacy auth", "Needs CA typically", "Record as gap on Free tier"),
        ("ADM-01", "Admin", "Restrict admin privilege", "Groups + Graph inventory", "Remove orphans; weekly review"),
        ("ADM-02", "Admin", "Minimise Global Admins", "Directory roles", "2 break-glass + JIT"),
        ("GST-01", "Guest", "Guest access reviews", "Governance paid", "Manual monthly export"),
        ("GST-02", "Guest", "Restrict invitations", "External collaboration settings", "Admins-only invite"),
        ("AUD-01", "Audit", "Unified audit log", "Usually on by default", "Define retention + review owner"),
        ("FWD-01", "Forwarding", "Detect external forward", "EXO cmdlets/licence", "Transport alert + monthly scan"),
        ("RSK-01", "Risky sign-in", "ID Protection", "P2 required", "Mark NotAvailable; use sign-in logs"),
        ("RSK-02", "Risky sign-in", "Sign-in monitoring", "Entra sign-in logs", "Weekly failed/unusual review"),
    ]
    for r in catalogue:
        guide.append(list(r))
    for col in guide.columns:
        guide.column_dimensions[col[0].column_letter].width = 28

    path = EVAL / "M365_Posture_Dashboard.xlsx"
    wb.save(path)
    return path


def build_report(summary):
    pct = summary.get("PosturePercent", 0) if summary else 0
    band = summary.get("RiskBand", "n/a") if summary else "n/a"
    crit = summary.get("SeverityCounts", {}).get("Critical", 0) if summary else 0
    high = summary.get("SeverityCounts", {}).get("High", 0) if summary else 0

    n_ctrl = summary.get("ControlCount", 0) if summary else 0
    sections = [
        (
            "Abstract",
            "I built a Microsoft 365 security posture assessment for a fictional managed-service "
            "SME (Coastal Managed IT) that operates on Entra ID Free–class capabilities. The "
            "framework checks MFA, admin accounts, guest access, audit logging, mailbox "
            "forwarding, and risky sign-in related controls. Each control is scored, rolled into "
            "an aggregate posture percent, and mapped to Critical/High/Medium/Low bands. An Excel "
            f"dashboard presents the scorecard for non-technical managers. The laboratory baseline "
            f"is synthetic and explicitly labelled; aggregate posture is {pct}% ({band}). "
            "Remediation advice stays honest about Free-tier limits: Conditional Access, Privileged "
            "Identity Management, and Identity Protection are marked NotAvailable where licences "
            "do not exist, rather than invented in screenshots or narrative.",
        ),
        (
            "1. Introduction",
            "Australian SMEs commonly adopt Microsoft 365 because it is familiar and bundled. "
            "Purchase does not equal posture. MFA may be partial, Global Administrators may "
            "accumulate, guests may linger, audit logs may exist without an owner, mailbox "
            "forwarding may quietly exfiltrate mail, and risky-sign-in automation may be assumed "
            "when it is a paid feature. Academic Project 15 asks for a security assessment "
            "framework grounded in public Microsoft guidance and synthetic tenant data, with "
            "defined risk categories, a tenant-friendly dashboard or report template, and safe "
            "remediation plus governance. The artefact is meant to be used before an SME buys "
            "every Entra add-on, and to stay usable when live Graph collection is blocked.",
        ),
        (
            "2. Related guidance",
            "Microsoft Secure Score and identity recommendations describe MFA for users and "
            "admins, reduction of standing privileged roles, and monitoring themes. The ASD "
            "Essential Eight elevates multi-factor authentication and restriction of "
            "administrative privileges as maturity-critical strategies. CIS Microsoft 365 "
            "Foundations Benchmarks add concrete checks for guests, auditing, and legacy "
            "protocols. NIST SP 800-63B informs authenticator strength language. This project "
            "does not reproduce vendor score UIs; it borrows the control ideas and restates "
            "them as an SME checklist that can run offline as JSON evidence.",
        ),
        (
            "3. Method",
            "The method is design science. First, six control families were fixed: MFA/"
            "authentication strength; admin accounts; guest access; audit logging; mailbox "
            "forwarding; risky sign-in and sign-in review. Second, PowerShell collectors write "
            "timestamped JSON. Default mode emits a synthetic Coastal Managed IT baseline so "
            "the pack is reproducible without Graph. Optional -Live mode is pinned to the "
            "laboratory tenant ID and falls back to synthetic if Connect-MgGraph fails or the "
            "tenant does not match. Third, a merger script computes sum(Score)/sum(MaxScore) "
            "and assigns a risk band. Fourth, an Excel workbook exposes a Dashboard sheet, a "
            "Findings sheet with severity colouring, and a Control_Catalogue sheet that records "
            "Free-tier notes beside each remediation hint. Live Graph was intentionally skipped "
            "in this environment after directory access failures; honesty about that limit is "
            "part of the method, not a defect to hide.",
        ),
        (
            "4. Results",
            f"The synthetic baseline scored {n_ctrl} controls with {crit} Critical and {high} "
            f"High findings. Aggregate posture is {pct}% — {band}. The Critical finding is "
            "standing privileged membership that still includes a disabled orphan account, a "
            "pattern aligned with joiner-mover-leaver failure modes studied in earlier identity "
            "work. High findings include incomplete MFA for privileged and standard users, "
            "absence of an automated external-forwarding scan, and Identity Protection marked "
            "at score zero because it is NotAvailable on Free tier. Medium findings cover open "
            "guest invitation settings, audit logging without a documented review process, and "
            "missing weekly sign-in review cadence. The dashboard lists five remediation themes: "
            "remove disabled privileged members weekly; enforce MFA for admins then users; "
            "shrink standing Global Admins in favour of break-glass plus time-boxed elevation; "
            "alert on new external forwarding; and document Free-tier NotAvailable controls "
            "instead of claiming them.",
        ),
        (
            "5. Discussion",
            "Separating process gaps from licence gaps matters for SME trust. A firm can fix "
            "orphan admins and MFA without buying P2. It cannot honestly claim risk-based "
            "Conditional Access or Identity Protection until it pays. The scorecard therefore "
            "allows a zero score with an explicit NotAvailable note, which is better governance "
            "than a green tick copied from marketing. Synthetic evidence is weaker than a live "
            "export, but it is stronger than fabricated portal screenshots. When Graph access "
            "returns, the same scripts can replace synthetic Items with read-only inventory "
            "while keeping the scoring and dashboard unchanged. Linking this assessment to "
            "joiner-mover-leaver automation and just-in-time privilege pilots closes the loop "
            "from finding to operational control.",
        ),
        (
            "6. Limits and ethics",
            "No production tenant was assessed. No passwords or client secrets are stored in "
            "the repository. No Secure Score or Entra portal screenshots were invented. Risky "
            "sign-in APIs and Conditional Access policy enumeration are out of Free-tier scope "
            "and are labelled accordingly. The fictional Coastal Managed IT case is for "
            "teaching and portfolio demonstration only.",
        ),
        (
            "7. Conclusion",
            "An SME can complete a first Microsoft 365 posture pass with a scored control "
            "catalogue, JSON evidence, and a manager-readable dashboard without purchasing "
            "every Entra identity add-on. Governance means naming an identity owner, reviewing "
            "Critical and High items monthly, and feeding findings into JML and JIT runbooks. "
            "The immediate next measurement is a live read-only Graph collection against the "
            "laboratory tenant when access is available, then a re-score that keeps the same "
            "bands and remediation language.",
        ),
        (
            "References",
            "Australian Signals Directorate. Essential Eight. Microsoft Learn: Secure Score; "
            "require multifactor authentication; mailbox auditing; detect external forwarding; "
            "Microsoft Entra ID Protection overview; restrict guest access. CIS Microsoft 365 "
            "Foundations Benchmark. NIST SP 800-63B Digital Identity Guidelines.",
        ),
    ]
    bw = sum(len(b.split()) for t, b in sections if t != "References")

    m = Document()
    tp = m.add_paragraph()
    tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tr = tp.add_run(
        "Microsoft 365 security posture assessment for SMEs\n"
        "A Free-tier-honest framework, scorecard, and remediation guide"
    )
    tr.bold = True
    tr.font.size = Pt(16)
    tr.font.color.rgb = NAVY
    mp = m.add_paragraph()
    mp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    mr = mp.add_run(
        f"Md Shohrab Hossen Niaz\nAcademic Project 15 — M365 Security Posture\n"
        f"Fictional SME case | Synthetic baseline (live Graph optional)\n"
        f"Word count (excluding title block and reference list): {bw}"
    )
    mr.font.size = Pt(11)
    for title, body in sections:
        heading(m, title, 1)
        m.add_paragraph(body)
    OUT.mkdir(parents=True, exist_ok=True)
    path = OUT / "01_AcademicProject15_Report.docx"
    m.save(path)
    print("BODY_WORDS", bw)
    return path


def build_slides(summary):
    pct = summary.get("PosturePercent", 0) if summary else 0
    band = summary.get("RiskBand", "n/a") if summary else "n/a"
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    navy = PRgb(0x1F, 0x4E, 0x79)

    def add(title, bullets, notes):
        s = prs.slides.add_slide(prs.slide_layouts[6])
        b = s.shapes.add_textbox(Inches(0.6), Inches(0.4), Inches(12), Inches(1))
        r = b.text_frame.paragraphs[0].add_run()
        r.text = title
        r.font.size = PPt(28)
        r.font.bold = True
        r.font.color.rgb = navy
        body = s.shapes.add_textbox(Inches(0.6), Inches(1.5), Inches(12), Inches(5))
        tf = body.text_frame
        tf.word_wrap = True
        first = True
        for line in bullets:
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            run = p.add_run()
            run.text = line
            run.font.size = PPt(18)
            p.space_after = PPt(10)
        s.notes_slide.notes_text_frame.text = notes

    add(
        "M365 posture assessment for SMEs",
        ["Framework + scored dashboard", "MFA, admins, guests, audit, forwarding, risky sign-in", "Free-tier honest — no invented portal UI"],
        "Academic Project 15.",
    )
    add(
        "Why it matters",
        ["SMEs buy M365 and trust defaults", "Standing admins, guests, forwarding stay invisible", "Managers need a simple risk band, not raw Secure Score jargon"],
        "Problem framing.",
    )
    add(
        "What I built",
        ["Control catalogue mapped to public Microsoft / ASD / CIS ideas", "PowerShell collectors (synthetic default, optional -Live)", f"Dashboard: posture {pct}% — {band}"],
        "Method and artefact.",
    )
    add(
        "Actions and limits",
        ["Clean privileged orphans; enforce MFA; shrink Global Admins", "Scan forwarding; review sign-ins weekly", "NotAvailable for CA / ID Protection on Free — stated, not faked"],
        "Questions?",
    )
    path = OUT / "02_AcademicProject15_Presentation.pptx"
    prs.save(path)
    return path


def assemble():
    summary = load_summary()
    build_framework()
    build_workbook(summary)
    build_report(summary)
    build_slides(summary)

    app_a = OUT / "Appendix" / "A_Evaluation"
    app_b = OUT / "Appendix" / "B_Framework"
    app_c = OUT / "Appendix" / "C_Scripts"
    app_d = OUT / "Appendix" / "D_Evidence"
    for p in (app_a, app_b, app_c, app_d):
        p.mkdir(parents=True, exist_ok=True)
    shutil.copy2(EVAL / "M365_Posture_Dashboard.xlsx", app_a / "M365_Posture_Dashboard.xlsx")
    for f in FRAME.glob("*.docx"):
        shutil.copy2(f, app_b / f.name)
    for f in ROOT.glob("*.ps1"):
        if not f.name.startswith("_"):
            shutil.copy2(f, app_c / f.name)
    if EVID.exists():
        for f in EVID.glob("*.json"):
            shutil.copy2(f, app_d / f.name)

    readme = Document()
    readme.add_heading("Project submission", 0)
    readme.add_paragraph(
        "Md Shohrab Hossen Niaz\nAcademic Project 15 — Microsoft 365 Security Posture Assessment\n"
        "Synthetic baseline JSON + Excel dashboard. Live Graph optional via -Live."
    )
    readme.save(OUT / "00_Read_Me_First.docx")
    print("assembled", OUT)


if __name__ == "__main__":
    assemble()

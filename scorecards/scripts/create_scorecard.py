#!/usr/bin/env python3
"""Creates the S.A.H NORDIC onboarding call scorecard as a .docx file."""

import zipfile
import os

OUTPUT = os.path.expanduser("~/Documents/Claude/SAH_Nordic_Onboarding_Scorecard.docx")

# ── colour palette ──────────────────────────────────────────────────────────
# Interceptly coral
CORAL     = "F27A7D"
CORAL_LT  = "FDEAEA"   # light coral for alternating rows / header bg
MAGENTA   = "FF00FF"
DARK      = "1F2937"   # near-black for body text
MID       = "4B5563"   # mid-grey for sub-labels
LIGHT_BG  = "F9FAFB"   # very light grey for row 2 shading
WHITE     = "FFFFFF"
YELLOW    = "FEF3C7"   # amber-tinted row
RED_LT    = "FEE2E2"
GREEN_LT  = "D1FAE5"
BLUE_LT   = "DBEAFE"

# ── helpers ──────────────────────────────────────────────────────────────────
def sz(pt):   return str(pt * 2)          # half-points
def twip(pt): return str(int(pt * 20))    # twentieths of a point (spacing)

def esc(t):
    return t.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")

def run(text, bold=False, italic=False, color=DARK, size_pt=11, font="Inter"):
    b  = "<w:b/>" if bold else ""
    i  = "<w:i/>" if italic else ""
    text = esc(text)
    return f"""<w:r>
      <w:rPr>
        <w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}"/>
        <w:sz w:val="{sz(size_pt)}"/><w:szCs w:val="{sz(size_pt)}"/>
        <w:color w:val="{color}"/>
        {b}{i}
      </w:rPr>
      <w:t xml:space="preserve">{text}</w:t>
    </w:r>"""

def para(children_xml, space_before=0, space_after=120, align="left",
         shading_color=None, left_indent=0, border_top=None, border_bottom=None):
    shading = f'<w:shd w:val="clear" w:color="auto" w:fill="{shading_color}"/>' if shading_color else ""
    indent  = f'<w:ind w:left="{left_indent}"/>' if left_indent else ""
    jc      = f'<w:jc w:val="{align}"/>' if align != "left" else ""
    bt = f'<w:top w:val="single" w:sz="6" w:color="{border_top}"/>' if border_top else ""
    bb = f'<w:bottom w:val="single" w:sz="4" w:color="{border_bottom}"/>' if border_bottom else ""
    pb_xml = f"<w:pBdr>{bt}{bb}</w:pBdr>" if (bt or bb) else ""
    return f"""<w:p>
    <w:pPr>
      <w:spacing w:before="{twip(space_before/20)}" w:after="{twip(space_after/20)}"/>
      {shading}{indent}{jc}{pb_xml}
    </w:pPr>
    {children_xml}
  </w:p>"""

def heading(text, level=1, color=CORAL, size_pt=None, space_before=160, space_after=80):
    sizes = {1: 22, 2: 14, 3: 12}
    sp = size_pt or sizes.get(level, 12)
    return para(run(text, bold=True, color=color, size_pt=sp, font="Gotham"),
                space_before=space_before, space_after=space_after,
                border_bottom=CORAL if level==1 else None)

def bullet(text, bold_prefix=None, color=DARK, indent=360):
    content = ""
    if bold_prefix:
        content += run(bold_prefix + " ", bold=True, color=color)
    content += run(text, color=color)
    return f"""<w:p>
    <w:pPr>
      <w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>
      <w:spacing w:before="0" w:after="60"/>
      <w:ind w:left="{indent}" w:hanging="240"/>
    </w:pPr>
    {content}
  </w:p>"""

def cell(children_xml, fill=WHITE, bold=False, color=DARK, size_pt=10,
         width=None, align="left", top=True, bottom=True, left=True, right=True,
         vAlign="center", colspan=None):
    b_attrs = lambda on: f'w:val="single" w:sz="4" w:color="E5E7EB"' if on else 'w:val="none"'
    b_top    = f'<w:top {b_attrs(top)}/>'
    b_bottom = f'<w:bottom {b_attrs(bottom)}/>'
    b_left   = f'<w:left {b_attrs(left)}/>'
    b_right  = f'<w:right {b_attrs(right)}/>'
    w_xml = f'<w:tcW w:w="{width}" w:type="dxa"/>' if width else ""
    sp_xml = f'<w:gridSpan w:val="{colspan}"/>' if colspan else ""
    va_xml = f'<w:vAlign w:val="{vAlign}"/>'
    return f"""<w:tc>
      <w:tcPr>
        {sp_xml}
        {w_xml}
        <w:shd w:val="clear" w:color="auto" w:fill="{fill}"/>
        <w:tcBorders>{b_top}{b_bottom}{b_left}{b_right}</w:tcBorders>
        <w:tcMar><w:top w:w="80" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/><w:left w:w="120" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tcMar>
        {va_xml}
      </w:tcPr>
      <w:p>
        <w:pPr>
          <w:jc w:val="{align}"/>
          <w:spacing w:before="0" w:after="0"/>
        </w:pPr>
        {children_xml}
      </w:p>
    </w:tc>"""

def row(*cells_xml, height=None):
    h = f'<w:trHeight w:val="{height}"/>' if height else ""
    return f"<w:tr><w:trPr>{h}</w:trPr>{''.join(cells_xml)}</w:tr>"

def table(*rows_xml, total_width=9360, col_widths=None):
    if col_widths:
        grid = "".join(f'<w:gridCol w:w="{w}"/>' for w in col_widths)
    else:
        grid = ""
    return f"""<w:tbl>
    <w:tblPr>
      <w:tblW w:w="{total_width}" w:type="dxa"/>
      <w:tblBorders>
        <w:insideH w:val="single" w:sz="4" w:color="E5E7EB"/>
        <w:insideV w:val="single" w:sz="4" w:color="E5E7EB"/>
      </w:tblBorders>
      <w:tblCellMar>
        <w:top w:w="80" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/>
        <w:left w:w="120" w:type="dxa"/><w:right w:w="120" w:type="dxa"/>
      </w:tblCellMar>
    </w:tblPr>
    <w:tblGrid>{grid}</w:tblGrid>
    {''.join(rows_xml)}
  </w:tbl>"""

def spacer(pt=6):
    return f'<w:p><w:pPr><w:spacing w:before="0" w:after="{twip(pt)}"/></w:pPr></w:p>'

# ── STATUS / BADGE helpers ────────────────────────────────────────────────────
def status(symbol, label, fill, text_color=DARK):
    return run(f" {symbol} {label} ", bold=True, color=text_color, size_pt=9)

# ── DOCUMENT BODY ─────────────────────────────────────────────────────────────
def build_body():
    parts = []

    # ── COVER HEADER ──────────────────────────────────────────────────────────
    parts.append(para(
        run("INTERCEPTLY", bold=True, color=WHITE, size_pt=11, font="Gotham"),
        shading_color=CORAL, align="center", space_before=0, space_after=0
    ))
    parts.append(para(
        run("Onboarding Call Scorecard", bold=True, color=DARK, size_pt=22, font="Gotham"),
        align="center", space_before=160, space_after=40
    ))
    parts.append(para(
        run("S.A.H NORDIC  ×  Chall | Interceptly", color=MID, size_pt=13, font="Inter"),
        align="center", space_before=0, space_after=200,
        border_bottom=CORAL
    ))

    # ── META TABLE ────────────────────────────────────────────────────────────
    W = [1560, 3120, 1560, 3120]
    def meta_row(l1, v1, l2, v2):
        return row(
            cell(run(l1, bold=True, color=MID, size_pt=9), fill=LIGHT_BG, width=W[0]),
            cell(run(v1, color=DARK, size_pt=10), fill=WHITE, width=W[1]),
            cell(run(l2, bold=True, color=MID, size_pt=9), fill=LIGHT_BG, width=W[2]),
            cell(run(v2, color=DARK, size_pt=10), fill=WHITE, width=W[3]),
        )
    parts.append(table(
        meta_row("DATE",     "14 March 2026",               "CSM",      "Chall"),
        meta_row("DURATION", "1h 2m",                       "SEGMENT",  "SMB"),
        meta_row("CALL TYPE","Onboarding / Platform Training","ACCOUNT", "S.A.H NORDIC"),
        meta_row("PACKAGE",  "Engage Pro",                  "CONTACT",  "Sharif"),
        total_width=9360, col_widths=W
    ))
    parts.append(spacer(12))

    # ── OVERALL SCORE BANNER ──────────────────────────────────────────────────
    # Big score row
    score_content = (
        run("Overall Onboarding Health Score", bold=True, color=WHITE, size_pt=13, font="Gotham") +
        run("      ", color=WHITE, size_pt=13) +
        run("72 / 100", bold=True, color=WHITE, size_pt=20, font="Gotham") +
        run("    ", color=WHITE, size_pt=13) +
        run("🟡  YELLOW — Needs Attention  (revised from 58 post-call)", bold=True, color=WHITE, size_pt=12, font="Gotham")
    )
    parts.append(para(score_content, shading_color=CORAL, align="center",
                      space_before=80, space_after=80))
    parts.append(spacer(14))

    # ── SECTION 1 ─────────────────────────────────────────────────────────────
    parts.append(heading("1.  Platform Setup & Configuration", level=1))
    parts.append(para(
        run("Score: ", bold=True, color=MID) + run("23 / 30", bold=True, color=CORAL, size_pt=12),
        space_before=0, space_after=60
    ))

    # Setup table
    W2 = [3600, 1200, 4560]
    def setup_row(item, status_sym, notes, fill=WHITE):
        return row(
            cell(run(item, color=DARK, size_pt=10), fill=fill, width=W2[0]),
            cell(run(status_sym, bold=True, color=DARK, size_pt=11), fill=fill, width=W2[1], align="center"),
            cell(run(notes, color=MID, size_pt=9, italic=True), fill=fill, width=W2[2]),
        )
    parts.append(table(
        row(cell(run("ITEM", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=W2[0]),
            cell(run("STATUS", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=W2[1], align="center"),
            cell(run("NOTES", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=W2[2])),
        setup_row("LinkedIn account connected (via proxy)", "✅", "Accountify proxy configured"),
        setup_row("Warm-up settings configured", "✅", "5/day start, +1 increment, max 20", fill=LIGHT_BG),
        setup_row("Connection Request Booster", "✅", "Toggled on; grayed out during warm-up"),
        setup_row("Pending connection request limit", "✅", "900 cap configured", fill=LIGHT_BG),
        setup_row("Sales Navigator confirmed linked", "✅", "System confirmed SN detected"),
        setup_row("Training videos accessible", "⚠️", "Missing during call; Chall followed up in post-call message", fill=YELLOW),
        setup_row("Campaign sequence ready", "⚠️", "Template duplicated; Interceptly team committed to providing content", fill=YELLOW),
        total_width=9360, col_widths=W2
    ))
    parts.append(para(
        run("Score rationale: ", bold=True, color=MID, size_pt=9) +
        run("Warm-up and account settings handled well. Training videos were unresolved during the call but Chall proactively followed up in the post-call message. Content is now committed by the Interceptly team.", color=MID, size_pt=9),
        space_before=60, space_after=120
    ))

    # ── SECTION 2 ─────────────────────────────────────────────────────────────
    parts.append(heading("2.  Product Knowledge Transfer", level=1))
    parts.append(para(
        run("Score: ", bold=True, color=MID) + run("20 / 25", bold=True, color=CORAL, size_pt=12),
        space_before=0, space_after=60
    ))
    W3 = [3000, 1200, 5160]
    def kt_row(topic, depth, notes, fill=WHITE):
        return row(
            cell(run(topic, color=DARK, size_pt=10), fill=fill, width=W3[0]),
            cell(run(depth, bold=True, size_pt=11, color=DARK), fill=fill, width=W3[1], align="center"),
            cell(run(notes, color=MID, size_pt=9, italic=True), fill=fill, width=W3[2]),
        )
    parts.append(table(
        row(cell(run("TOPIC", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=W3[0]),
            cell(run("DEPTH", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=W3[1], align="center"),
            cell(run("NOTES", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=W3[2])),
        kt_row("Dashboard overview", "✅ Full", "Activity feed, stats, tasks explained clearly"),
        kt_row("Campaign creation & sequence logic", "✅ Full", "Template duplication, actions, conditions, approval steps", fill=LIGHT_BG),
        kt_row("Prospects & Inbox", "✅ Full", "Saved replies feature highlighted — relevant to customer concern"),
        kt_row("Reports", "✅ Full", "Daily activity tracking explained", fill=LIGHT_BG),
        kt_row("Blacklist / safety settings", "✅ Full", "Automation stop-on-reply behaviour covered"),
        kt_row("Sales Navigator walkthrough", "✅ Full", "Account → Lead filtering, list creation, URL paste to campaign", fill=LIGHT_BG),
        kt_row("AI personalisation feature", "ℹ️ N/A", "Not included in Engage Pro package — Chall should have stated this clearly vs. 'we'll double check'", fill=LIGHT_BG),
        kt_row("Email + LinkedIn multi-channel", "✅ Full", "Customer chose LinkedIn-only for now — appropriate decision", fill=LIGHT_BG),
        kt_row("Integrations", "✅ Covered", "Shown; customer confirmed not needed at this stage"),
        total_width=9360, col_widths=W3
    ))
    parts.append(para(
        run("Score rationale: ", bold=True, color=MID, size_pt=9) +
        run("Coverage was comprehensive. AI features are confirmed not included in Engage Pro — the only gap was that Chall was not immediately confident in stating this during the call. A CSM should know their customer's package before the session.", color=MID, size_pt=9),
        space_before=60, space_after=120
    ))

    # ── SECTION 3 ─────────────────────────────────────────────────────────────
    parts.append(heading("3.  Customer Goals & Expectation Alignment", level=1))
    parts.append(para(
        run("Score: ", bold=True, color=MID) + run("14 / 20", bold=True, color=CORAL, size_pt=12),
        space_before=0, space_after=60
    ))
    W4 = [4200, 720, 4440]
    def goal_row(signal, icon, assessment, fill=WHITE):
        return row(
            cell(run(signal, color=DARK, size_pt=10), fill=fill, width=W4[0]),
            cell(run(icon, size_pt=11), fill=fill, width=W4[1], align="center"),
            cell(run(assessment, color=MID, size_pt=9, italic=True), fill=fill, width=W4[2]),
        )
    parts.append(table(
        row(cell(run("SIGNAL", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=W4[0]),
            cell(run("", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=W4[1]),
            cell(run("ASSESSMENT", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=W4[2])),
        goal_row("Customer's primary goal understood", "✅", "Generating sales calls via LinkedIn outreach"),
        goal_row("Target audience defined", "✅", "SMB B2B owners, IT/tech sector, US market", fill=LIGHT_BG),
        goal_row("Budget sensitivity acknowledged", "✅", "Budget-constrained; full-time job alongside this"),
        goal_row("Realistic timeline & KPIs set", "⚠️", "Customer expects results in a few months — no success milestone established", fill=YELLOW),
        goal_row("Content creation expectation", "✅", "Resolved post-call — Interceptly team committed to providing campaign content in after-call message", fill=GREEN_LT),
        goal_row("Package limitations set (no AI, no follow-up calls)", "⚠️", "Customer not explicitly told during call; relies on in-platform chat for ongoing support", fill=YELLOW),
        goal_row("Email migration pathway", "✅", "Handled well — no pressure, clear cost comparison vs. Instantly", fill=LIGHT_BG),
        goal_row("Reply handling strategy", "⚠️", "Saved replies shown but no reply playbook provided; customer flagged anxiety about not being a salesperson", fill=YELLOW),
        goal_row("Workspace decision", "✅", "Separate workspace confirmed and well explained"),
        total_width=9360, col_widths=W4
    ))
    parts.append(para(
        run("Score rationale: ", bold=True, color=MID, size_pt=9) +
        run("Content mismatch resolved post-call — Chall committed to providing the content, removing the biggest activation blocker. Key remaining gap: package limitations (no AI, no scheduled follow-up calls) were not set as explicit expectations during the session, and no reply handling plan was given to a customer who self-identified as not a salesperson.", color=MID, size_pt=9),
        space_before=60, space_after=120
    ))

    # ── SECTION 4 ─────────────────────────────────────────────────────────────
    parts.append(heading("4.  Next Steps & Accountability", level=1))
    parts.append(para(
        run("Score: ", bold=True, color=MID) + run("11 / 15", bold=True, color=CORAL, size_pt=12),
        space_before=0, space_after=60
    ))
    W5 = [5040, 4320]
    def ns_row(item, status_sym, fill=WHITE):
        return row(
            cell(run(item, color=DARK, size_pt=10), fill=fill, width=W5[0]),
            cell(run(status_sym, size_pt=11, bold=True), fill=fill, width=W5[1], align="center"),
        )
    parts.append(table(
        row(cell(run("ITEM", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=W5[0]),
            cell(run("STATUS", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=W5[1], align="center")),
        ns_row("Follow-up call scheduled", "ℹ️  N/A — not included in Engage Pro package"),
        ns_row("Help articles sent (3 articles)", "✅  Sent post-call", fill=GREEN_LT),
        ns_row("Content creation committed", "✅  Interceptly team to provide content", fill=GREEN_LT),
        ns_row("Training video follow-up", "✅  Asked in post-call message; awaiting response", fill=LIGHT_BG),
        ns_row("AI feature status resolved", "✅  Confirmed not in Engage Pro package", fill=GREEN_LT),
        ns_row("Warm-up timeline communicated", "✅  ~3 weeks to 20/day confirmed", fill=LIGHT_BG),
        ns_row("2 separate prospect lists explained", "✅  Noted in post-call message per Lee's recommendation"),
        ns_row("Reply handling playbook provided", "❌  Not sent — customer flagged anxiety; still outstanding", fill=RED_LT),
        total_width=9360, col_widths=W5
    ))
    parts.append(para(
        run("Score rationale: ", bold=True, color=MID, size_pt=9) +
        run("Strong post-call message resolved most open items — content committed, articles sent, video follow-up asked, and 2-list structure explained. No follow-up calls is a package limitation, not a failure. One remaining gap: no reply playbook sent despite the customer explicitly flagging they are not a salesperson.", color=MID, size_pt=9),
        space_before=60, space_after=120
    ))

    # ── SECTION 5 ─────────────────────────────────────────────────────────────
    parts.append(heading("5.  Relationship & Rapport", level=1))
    parts.append(para(
        run("Score: ", bold=True, color=MID) + run("5 / 10", bold=True, color=CORAL, size_pt=12),
        space_before=0, space_after=60
    ))
    W6 = [4200, 720, 4440]
    def rr_row(signal, icon, assessment, fill=WHITE):
        return row(
            cell(run(signal, color=DARK, size_pt=10), fill=fill, width=W6[0]),
            cell(run(icon, size_pt=11), fill=fill, width=W6[1], align="center"),
            cell(run(assessment, color=MID, size_pt=9, italic=True), fill=fill, width=W6[2]),
        )
    parts.append(table(
        row(cell(run("SIGNAL", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=W6[0]),
            cell(run("", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=W6[1]),
            cell(run("ASSESSMENT", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=W6[2])),
        rr_row("Customer comfort level", "✅", "Relaxed tone; asked questions freely throughout"),
        rr_row("CSM responsiveness", "✅", "Patient, clear, non-pushy", fill=LIGHT_BG),
        rr_row("Technical issues managed", "⚠️", "Chall's internet quality impacted the call on two occasions", fill=YELLOW),
        rr_row("Dual role (partner + customer) acknowledged", "✅", "Noted naturally mid-call", fill=LIGHT_BG),
        rr_row("Call overran", "⚠️", "Ran ~10 min over — flagged but customer agreed to continue", fill=YELLOW),
        total_width=9360, col_widths=W6
    ))
    parts.append(spacer(10))

    # ── OPEN ISSUES ──────────────────────────────────────────────────────────
    parts.append(heading("Open Issues — Action Required", level=1, color="DC2626"))
    W7 = [360, 3240, 960, 2040, 2760]
    def issue_row(num, issue, priority, owner, deadline, fill=WHITE):
        p_color = {"🔴 High": "DC2626", "🟡 Medium": "D97706"}.get(priority, DARK)
        return row(
            cell(run(num, bold=True, color=CORAL, size_pt=10), fill=fill, width=W7[0]),
            cell(run(issue, color=DARK, size_pt=10), fill=fill, width=W7[1]),
            cell(run(priority, bold=True, color=p_color, size_pt=9), fill=fill, width=W7[2]),
            cell(run(owner, color=DARK, size_pt=9), fill=fill, width=W7[3]),
            cell(run(deadline, color=MID, size_pt=9, italic=True), fill=fill, width=W7[4]),
        )
    parts.append(table(
        row(cell(run("#", bold=True, color=WHITE, size_pt=9), fill="DC2626", width=W7[0]),
            cell(run("ISSUE", bold=True, color=WHITE, size_pt=9), fill="DC2626", width=W7[1]),
            cell(run("PRIORITY", bold=True, color=WHITE, size_pt=9), fill="DC2626", width=W7[2]),
            cell(run("OWNER", bold=True, color=WHITE, size_pt=9), fill="DC2626", width=W7[3]),
            cell(run("DEADLINE", bold=True, color=WHITE, size_pt=9), fill="DC2626", width=W7[4])),
        issue_row("1", "Training videos still missing — awaiting customer confirmation", "🟡 Medium", "Chall / Support", "Confirm this week", fill=YELLOW),
        issue_row("2", "Reply handling playbook not sent — customer is not a salesperson", "🔴 High", "Chall", "Send this week", fill=RED_LT),
        issue_row("3", "Package limitations not set explicitly during call (no AI, no follow-up calls)", "🟡 Medium", "Chall", "Clarify before first reply comes in", fill=YELLOW),
        issue_row("4", "Campaign content — timing and format of delivery not confirmed", "🔴 High", "Chall + Lee", "Confirm before warm-up completes (Week 3)", fill=RED_LT),
        total_width=9360, col_widths=W7
    ))
    parts.append(spacer(10))

    # ── RISK FLAGS ────────────────────────────────────────────────────────────
    parts.append(heading("Risk Flags", level=1, color="92400E"))
    for title, detail in [
        ("Capacity risk", "Customer has a full-time job — limited availability to respond to LinkedIn replies promptly, especially across time zones."),
        ("Reply handling anxiety (no AI assist)", "Customer explicitly said 'I'm not a salesperson' and asked about AI reply suggestions — which are not in Engage Pro. Without a reply playbook, early responses may be mishandled and kill momentum before traction is established."),
        ("Activation timing", "Content delivery from Interceptly must land before warm-up completes (~3 weeks). If not, Time to First Value slips past 30 days and churn risk rises sharply."),
        ("No follow-up calls in package", "Customer is onboarded via a single session with in-platform support only. Risk of abandonment if they get stuck — particularly given their limited sales background."),
        ("Email platform lock-in", "Not in Instantly's integration catalog. If LinkedIn outreach is slow to show results, customer may deprioritise the platform entirely."),
    ]:
        parts.append(bullet(detail, bold_prefix=title, color=DARK))
    parts.append(spacer(10))

    # ── WHAT WENT WELL ────────────────────────────────────────────────────────
    parts.append(heading("What Went Well", level=1, color="065F46"))
    for item in [
        "Warm-up configured correctly within Accountify's safety limits — good attention to account health.",
        "Sales Navigator walkthrough was practical and hands-on; customer left with a real search list already added to their campaign.",
        "Workspace decision explained clearly and proactively — handled before it became a point of confusion.",
        "Saved replies feature surfaced at exactly the right moment in response to the customer's reply-handling concern.",
        "Strong after-call message: clear next steps, 3 relevant help articles, training video follow-up, content commitment, and 2-list structure explained.",
        "Tone throughout was relaxed and trust-building; customer finished the call engaged and motivated.",
    ]:
        parts.append(bullet(item, color=DARK))
    parts.append(spacer(10))

    # ── RECOMMENDED NEXT STEPS ────────────────────────────────────────────────
    parts.append(heading("Recommended Next Steps", level=1, color=CORAL))
    steps = [
        ("This week", "Confirm training videos are now visible — follow up if Sharif hasn't responded to the post-call message."),
        ("This week", "Send a reply playbook: 3–5 saved reply templates covering the most common LinkedIn responses (interested, not now, wrong person, tell me more)."),
        ("This week", "Confirm content delivery timeline with Lee — Sharif needs both sets of campaign content (owners + other roles) ready before Week 3."),
        ("This week", "Clarify Engage Pro support model with Sharif — set expectation that in-platform chat is the primary support channel going forward."),
        ("Week 3", "Warm-up check-in — if at 15–20/day, confirm campaign content is uploaded and ready to launch."),
        ("Day 30", "First value milestone: at least one campaign live with first connection requests sent and first reply received."),
    ]
    W8 = [1440, 7920]
    parts.append(table(
        *[row(
            cell(run(when, bold=True, color=CORAL, size_pt=9), fill=CORAL_LT if i%2==0 else WHITE, width=W8[0]),
            cell(run(action, color=DARK, size_pt=10), fill=CORAL_LT if i%2==0 else WHITE, width=W8[1]),
          ) for i, (when, action) in enumerate(steps)],
        total_width=9360, col_widths=W8
    ))
    parts.append(spacer(16))

    # ── FOOTER NOTE ───────────────────────────────────────────────────────────
    parts.append(para(
        run("Scored against Interceptly CSM onboarding framework — SMB segment thresholds.  |  Health: 🟡 Yellow (50–74)  |  Package: Engage Pro (no AI, no follow-up calls)  |  Revised 14 March 2026",
            color=MID, size_pt=8, italic=True),
        align="center", space_before=80, space_after=0,
        border_top="E5E7EB"
    ))

    return "\n".join(parts)


# ── NUMBERING ─────────────────────────────────────────────────────────────────
NUMBERING_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="&#x2022;"/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="360" w:hanging="240"/></w:pPr>
      <w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/></w:rPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>"""

# ── DOCUMENT XML ──────────────────────────────────────────────────────────────
DOCUMENT_XML = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
  <w:body>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080"
               w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
    {build_body()}
  </w:body>
</w:document>"""

# ── STYLES XML ────────────────────────────────────────────────────────────────
STYLES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="Inter" w:hAnsi="Inter" w:cs="Inter"/>
      <w:sz w:val="22"/><w:szCs w:val="22"/>
      <w:color w:val="1F2937"/>
    </w:rPr></w:rPrDefault>
  </w:docDefaults>
  <w:style w:type="character" w:styleId="Hyperlink">
    <w:name w:val="Hyperlink"/>
    <w:rPr><w:color w:val="F27A7D"/><w:u w:val="single"/></w:rPr>
  </w:style>
</w:styles>"""

# ── RELATIONSHIPS ─────────────────────────────────────────────────────────────
RELS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
</Relationships>"""

CONTENT_TYPES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
</Types>"""

DOTRELS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

# ── WRITE DOCX ────────────────────────────────────────────────────────────────
with zipfile.ZipFile(OUTPUT, "w", zipfile.ZIP_DEFLATED) as zf:
    zf.writestr("[Content_Types].xml",         CONTENT_TYPES_XML)
    zf.writestr("_rels/.rels",                 DOTRELS_XML)
    zf.writestr("word/document.xml",           DOCUMENT_XML)
    zf.writestr("word/styles.xml",             STYLES_XML)
    zf.writestr("word/numbering.xml",          NUMBERING_XML)
    zf.writestr("word/_rels/document.xml.rels",RELS_XML)

print(f"Created: {OUTPUT}")

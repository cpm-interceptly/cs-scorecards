#!/usr/bin/env python3
"""Serve First — SDR Training & Platform Overview scorecard."""

import zipfile, os

OUTPUT = os.path.expanduser(
    "~/Documents/Claude/ServeFirst_Training_Scorecard_Mar10.docx"
)

# ── colour palette ─────────────────────────────────────────────────────────
CORAL    = "F27A7D"
CORAL_LT = "FDEAEA"
DARK     = "1F2937"
MID      = "4B5563"
LIGHT_BG = "F9FAFB"
WHITE    = "FFFFFF"
YELLOW   = "FEF3C7"
RED_LT   = "FEE2E2"
GREEN_LT = "D1FAE5"
BLUE_LT  = "DBEAFE"

def sz(pt):   return str(pt * 2)
def twip(pt): return str(int(pt * 20))
def esc(t):
    return t.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")

def run(text, bold=False, italic=False, color=DARK, size_pt=11, font="Inter"):
    b = "<w:b/>" if bold else ""
    i = "<w:i/>" if italic else ""
    return f"""<w:r>
      <w:rPr>
        <w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}"/>
        <w:sz w:val="{sz(size_pt)}"/><w:szCs w:val="{sz(size_pt)}"/>
        <w:color w:val="{color}"/>{b}{i}
      </w:rPr>
      <w:t xml:space="preserve">{esc(text)}</w:t>
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
                border_bottom=CORAL if level == 1 else None)

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

def cell(children_xml, fill=WHITE, width=None, align="left",
         top=True, bottom=True, left=True, right=True, colspan=None):
    b_attrs = lambda on: 'w:val="single" w:sz="4" w:color="E5E7EB"' if on else 'w:val="none"'
    borders = (f'<w:top {b_attrs(top)}/><w:bottom {b_attrs(bottom)}/>'
               f'<w:left {b_attrs(left)}/><w:right {b_attrs(right)}/>')
    w_xml  = f'<w:tcW w:w="{width}" w:type="dxa"/>' if width else ""
    sp_xml = f'<w:gridSpan w:val="{colspan}"/>' if colspan else ""
    return f"""<w:tc>
      <w:tcPr>
        {sp_xml}{w_xml}
        <w:shd w:val="clear" w:color="auto" w:fill="{fill}"/>
        <w:tcBorders>{borders}</w:tcBorders>
        <w:tcMar>
          <w:top w:w="80" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/>
          <w:left w:w="120" w:type="dxa"/><w:right w:w="120" w:type="dxa"/>
        </w:tcMar>
        <w:vAlign w:val="center"/>
      </w:tcPr>
      <w:p>
        <w:pPr><w:jc w:val="{align}"/><w:spacing w:before="0" w:after="0"/></w:pPr>
        {children_xml}
      </w:p>
    </w:tc>"""

def row(*cells_xml, height=None):
    h = f'<w:trHeight w:val="{height}"/>' if height else ""
    return f"<w:tr><w:trPr>{h}</w:trPr>{''.join(cells_xml)}</w:tr>"

def table(*rows_xml, total_width=9360, col_widths=None):
    grid = "".join(f'<w:gridCol w:w="{w}"/>' for w in col_widths) if col_widths else ""
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


# ── DOCUMENT BODY ──────────────────────────────────────────────────────────────
def build_body():
    parts = []

    # ── HEADER BANNER ─────────────────────────────────────────────────────────
    parts.append(para(
        run("INTERCEPTLY", bold=True, color=WHITE, size_pt=11, font="Gotham"),
        shading_color=CORAL, align="center", space_before=0, space_after=0
    ))
    parts.append(para(
        run("SDR Training & Platform Overview — Call Scorecard", bold=True, color=DARK, size_pt=20, font="Gotham"),
        align="center", space_before=160, space_after=40
    ))
    parts.append(para(
        run("Serve First  ×  Alison | Interceptly", color=MID, size_pt=13, font="Inter"),
        align="center", space_before=0, space_after=200, border_bottom=CORAL
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
        meta_row("DATE",      "10 March 2026",                  "CSM",       "Alison"),
        meta_row("DURATION",  "39 minutes",                     "SUPPORT",   "Chall (technical)"),
        meta_row("CALL TYPE", "SDR Training / Platform Overview","ACCOUNT",  "Serve First"),
        meta_row("PACKAGE",   "Pipeline Builder — 5 seats",     "ATTENDEES", "Tim, Oliver, Paul, Connor, Katie, Alex"),
        total_width=9360, col_widths=W
    ))
    parts.append(spacer(12))

    # ── ATTENDEE CONTEXT BOX ──────────────────────────────────────────────────
    parts.append(heading("Attendee Context", level=1, color=MID))
    Wa = [2400, 2160, 4800]
    parts.append(table(
        row(cell(run("PERSON", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=Wa[0]),
            cell(run("ROLE", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=Wa[1]),
            cell(run("CONTEXT", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=Wa[2])),
        row(cell(run("Tim", color=DARK, size_pt=10), fill=LIGHT_BG, width=Wa[0]),
            cell(run("Board / Senior stakeholder", color=DARK, size_pt=10), fill=LIGHT_BG, width=Wa[1]),
            cell(run("Familiar with platform; ran Errol's campaign; driving adoption internally", color=MID, size_pt=9, italic=True), fill=LIGHT_BG, width=Wa[2])),
        row(cell(run("Oliver Ayoub", color=DARK, size_pt=10), fill=WHITE, width=Wa[0]),
            cell(run("AE", color=DARK, size_pt=10), fill=WHITE, width=Wa[1]),
            cell(run("On the platform; already has a Deloitte meeting booked from campaign", color=MID, size_pt=9, italic=True), fill=WHITE, width=Wa[2])),
        row(cell(run("Paul", color=DARK, size_pt=10), fill=LIGHT_BG, width=Wa[0]),
            cell(run("New Marketing Director", color=DARK, size_pt=10), fill=LIGHT_BG, width=Wa[1]),
            cell(run("NEVER seen platform — first contact with Interceptly team; key evaluator for expansion", color=MID, size_pt=9, italic=True), fill=RED_LT, width=Wa[2])),
        row(cell(run("Connor", color=DARK, size_pt=10), fill=WHITE, width=Wa[0]),
            cell(run("AE (existing seat)", color=DARK, size_pt=10), fill=WHITE, width=Wa[1]),
            cell(run("Could not log in during call", color=MID, size_pt=9, italic=True), fill=YELLOW, width=Wa[2])),
        row(cell(run("Johnny", color=DARK, size_pt=10), fill=LIGHT_BG, width=Wa[0]),
            cell(run("AE (existing seat)", color=DARK, size_pt=10), fill=LIGHT_BG, width=Wa[1]),
            cell(run("Could not join the call", color=MID, size_pt=9, italic=True), fill=YELLOW, width=Wa[2])),
        row(cell(run("Katie", color=DARK, size_pt=10), fill=WHITE, width=Wa[0]),
            cell(run("New AE (potential seat)", color=DARK, size_pt=10), fill=WHITE, width=Wa[1]),
            cell(run("Introduced at start — not engaged or addressed again during call", color=MID, size_pt=9, italic=True), fill=RED_LT, width=Wa[2])),
        row(cell(run("Alex", color=DARK, size_pt=10), fill=LIGHT_BG, width=Wa[0]),
            cell(run("New AE (potential seat)", color=DARK, size_pt=10), fill=LIGHT_BG, width=Wa[1]),
            cell(run("Introduced at start — not engaged or addressed again during call", color=MID, size_pt=9, italic=True), fill=RED_LT, width=Wa[2])),
        total_width=9360, col_widths=Wa
    ))
    parts.append(spacer(12))

    # ── OVERALL SCORE BANNER ──────────────────────────────────────────────────
    score_content = (
        run("Overall Training Call Health Score", bold=True, color=WHITE, size_pt=13, font="Gotham") +
        run("      ", color=WHITE, size_pt=13) +
        run("60 / 100", bold=True, color=WHITE, size_pt=20, font="Gotham") +
        run("    ", color=WHITE, size_pt=13) +
        run("🟡  YELLOW — Needs Attention  (revised from 54 post-call)", bold=True, color=WHITE, size_pt=12, font="Gotham")
    )
    parts.append(para(score_content, shading_color=CORAL, align="center",
                      space_before=80, space_after=80))
    parts.append(spacer(14))

    # ── SECTION 1: PRE-CALL PREPARATION ───────────────────────────────────────
    parts.append(heading("1.  Pre-Call Preparation", level=1))
    parts.append(para(
        run("Score: ", bold=True, color=MID) + run("8 / 20", bold=True, color=CORAL, size_pt=12),
        space_before=0, space_after=60
    ))
    W2 = [3600, 1200, 4560]
    def prep_row(item, sym, notes, fill=WHITE):
        return row(
            cell(run(item, color=DARK, size_pt=10), fill=fill, width=W2[0]),
            cell(run(sym, bold=True, size_pt=11, color=DARK), fill=fill, width=W2[1], align="center"),
            cell(run(notes, color=MID, size_pt=9, italic=True), fill=fill, width=W2[2]),
        )
    parts.append(table(
        row(cell(run("ITEM", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=W2[0]),
            cell(run("STATUS", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=W2[1], align="center"),
            cell(run("NOTES", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=W2[2])),
        prep_row("Login access verified pre-call", "❌", "3 of 5 users could not log in — rebranding broke URLs/passwords; should have been pre-checked", fill=RED_LT),
        prep_row("Structured session agenda shared", "❌", "No agenda communicated at the start; session was reactive throughout", fill=RED_LT),
        prep_row("Paul (new MD) briefed / intro prepared", "❌", "Paul had never seen the platform — no tailored intro or framing for a first-time senior stakeholder", fill=RED_LT),
        prep_row("New AEs (Katie & Alex) considered", "⚠️", "Mentioned as reason for the call but no specific content or demo prepared for them", fill=YELLOW),
        prep_row("Technical support (Chall) on standby", "✅", "Chall joined and resolved password resets during the call"),
        prep_row("Screen sharing worked", "✅", "Alison's screen share was functional throughout", fill=LIGHT_BG),
        total_width=9360, col_widths=W2
    ))
    parts.append(para(
        run("Score rationale: ", bold=True, color=MID, size_pt=9) +
        run("The login failure for 3 users is the single biggest failure of the call — you cannot run a platform training session when your attendees cannot log in. This was preventable with a 10-minute pre-call access check. The absence of a structured agenda and no preparation for Paul (the most important new stakeholder) further compounds the preparation gap.", color=MID, size_pt=9),
        space_before=60, space_after=120
    ))

    # ── SECTION 2: PLATFORM DEMO & KNOWLEDGE TRANSFER ─────────────────────────
    parts.append(heading("2.  Platform Demo & Knowledge Transfer", level=1))
    parts.append(para(
        run("Score: ", bold=True, color=MID) + run("16 / 25", bold=True, color=CORAL, size_pt=12),
        space_before=0, space_after=60
    ))
    W3 = [3000, 1200, 5160]
    def demo_row(topic, depth, notes, fill=WHITE):
        return row(
            cell(run(topic, color=DARK, size_pt=10), fill=fill, width=W3[0]),
            cell(run(depth, bold=True, size_pt=11, color=DARK), fill=fill, width=W3[1], align="center"),
            cell(run(notes, color=MID, size_pt=9, italic=True), fill=fill, width=W3[2]),
        )
    parts.append(table(
        row(cell(run("TOPIC", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=W3[0]),
            cell(run("DEPTH", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=W3[1], align="center"),
            cell(run("NOTES", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=W3[2])),
        demo_row("Company / service intro", "✅ Full", "Clear explanation of what Interceptly does; Alison's role and experience framed well"),
        demo_row("Dashboard overview", "✅ Full", "Active campaigns shown; stats and layout explained", fill=LIGHT_BG),
        demo_row("Lead Intercept feature", "✅ Full", "Concept, signals, and filter workflow explained; dashboard shown"),
        demo_row("Inbox & reply management", "✅ Full", "Paused campaigns, reply workflow, live examples shown (Oliver's inbox)", fill=LIGHT_BG),
        demo_row("Comment approval workflow", "✅ Full", "Approval step shown; Oliver responded positively — strong product moment"),
        demo_row("Campaign sequence creation / editing", "❌ Not shown", "Oliver asked about cadence templates — deferred to individual sessions; gap for new AEs", fill=RED_LT),
        demo_row("Messaging personalisation walkthrough", "❌ Not shown", "Discussed conceptually but not demonstrated; Oliver specifically asked for this", fill=RED_LT),
        demo_row("HubSpot integration", "⚠️ Brief", "Confirmed it exists and is working; no walkthrough", fill=YELLOW),
        demo_row("Signal quality explanation (Lead Intercept)", "⚠️ Weak", "Oliver challenged the value of signals — Alison's answer was vague; Tim stepped in to explain", fill=YELLOW),
        total_width=9360, col_widths=W3
    ))
    parts.append(para(
        run("Score rationale: ", bold=True, color=MID, size_pt=9) +
        run("Core features were demonstrated effectively and the inbox/reply flow landed well. The gaps are campaign creation and message editing — which are exactly what AEs need to use daily. Oliver had to be directed to book a separate session for something that should have been in this training. Tim also carried too much of the explanation; the CSM should be leading the demo.", color=MID, size_pt=9),
        space_before=60, space_after=120
    ))

    # ── SECTION 3: STAKEHOLDER MANAGEMENT ─────────────────────────────────────
    parts.append(heading("3.  Stakeholder Management", level=1))
    parts.append(para(
        run("Score: ", bold=True, color=MID) + run("10 / 20", bold=True, color=CORAL, size_pt=12),
        space_before=0, space_after=60
    ))
    W4 = [4200, 720, 4440]
    def sh_row(signal, icon, notes, fill=WHITE):
        return row(
            cell(run(signal, color=DARK, size_pt=10), fill=fill, width=W4[0]),
            cell(run(icon, size_pt=11), fill=fill, width=W4[1], align="center"),
            cell(run(notes, color=MID, size_pt=9, italic=True), fill=fill, width=W4[2]),
        )
    parts.append(table(
        row(cell(run("SIGNAL", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=W4[0]),
            cell(run("", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=W4[1]),
            cell(run("ASSESSMENT", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=W4[2])),
        sh_row("Paul (new MD) formally oriented", "❌", "Never individually acknowledged, briefed, or asked for input — most important new stakeholder", fill=RED_LT),
        sh_row("Katie & Alex (potential seats) engaged", "❌", "Introduced at the start and then completely dropped; no demo tailored to them", fill=RED_LT),
        sh_row("Oliver (AE) engaged effectively", "✅", "Active dialogue, real questions answered, left with clear next step", fill=GREEN_LT),
        sh_row("Tim's co-presentation managed", "⚠️", "Tim explained Lead Intercept value and ROI himself — Alison should have led this", fill=YELLOW),
        sh_row("Login issues handled gracefully", "⚠️", "Chall resolved passwords live; awkward but managed without derailing call", fill=YELLOW),
        sh_row("Multi-audience balance (new vs. familiar)", "❌", "Session defaulted to Oliver's level — new users (Paul, Katie, Alex) were lost or ignored", fill=RED_LT),
        sh_row("Closing checked in with all attendees", "❌", "Wrap-up addressed Tim only; Paul, Katie, Alex received no individual close", fill=RED_LT),
        sh_row("Post-call email sent to Tim with individual AE sessions proposed", "✅", "Structured recap with clear action items split by party; individual sessions formally offered to all AEs", fill=GREEN_LT),
        total_width=9360, col_widths=W4
    ))
    parts.append(para(
        run("Score rationale: ", bold=True, color=MID, size_pt=9) +
        run("Paul is the new marketing director who had never seen the platform — if he walks away unimpressed or confused, that is a direct risk to seat expansion. Katie and Alex are the reason the call was called; they were never spoken to again after introductions. The post-call email goes to Tim only and does not name or address Paul, Katie, or Alex individually — a missed opportunity to close the stakeholder gap.", color=MID, size_pt=9),
        space_before=60, space_after=120
    ))

    # ── SECTION 4: PRODUCT POSITIONING & VALUE COMMUNICATION ──────────────────
    parts.append(heading("4.  Product Positioning & Value Communication", level=1))
    parts.append(para(
        run("Score: ", bold=True, color=MID) + run("10 / 15", bold=True, color=CORAL, size_pt=12),
        space_before=0, space_after=60
    ))
    W5 = [4200, 720, 4440]
    def pv_row(signal, icon, notes, fill=WHITE):
        return row(
            cell(run(signal, color=DARK, size_pt=10), fill=fill, width=W5[0]),
            cell(run(icon, size_pt=11), fill=fill, width=W5[1], align="center"),
            cell(run(notes, color=MID, size_pt=9, italic=True), fill=fill, width=W5[2]),
        )
    parts.append(table(
        row(cell(run("SIGNAL", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=W5[0]),
            cell(run("", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=W5[1]),
            cell(run("ASSESSMENT", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=W5[2])),
        pv_row("Value proposition explained clearly at start", "✅", "Interceptly's role in filling the outreach strategy/management gap was articulated well"),
        pv_row("Rebranding confusion managed", "⚠️", "Login URL issues flagged the rebrand problem; Alison referenced Outreachly at times — inconsistent", fill=YELLOW),
        pv_row("ROI / results communicated", "⚠️", "Tim gave the numbers (3-4 AQLs/month, 1,000 connections for Errol) — Alison should have led this", fill=YELLOW),
        pv_row("Concrete proof point surfaced (Oliver's Deloitte meeting)", "✅", "Strong ROI moment — Oliver himself confirmed a booked meeting from the campaign", fill=GREEN_LT),
        pv_row("'Marginal gains' framing handled", "✅", "Tim set the right expectation; Alison reinforced it — good alignment on realistic outcomes"),
        pv_row("Expansion opportunity (Katie, Alex, Paul) actively positioned", "❌", "No explicit pitch or case made for adding the new seats during the call", fill=RED_LT),
        pv_row("Rebrand to Interceptly explained proactively", "⚠️", "Only mentioned when asked; should have been framed positively at the start of the call", fill=YELLOW),
        total_width=9360, col_widths=W5
    ))
    parts.append(para(
        run("Score rationale: ", bold=True, color=MID, size_pt=9) +
        run("The core value proposition landed and the Deloitte proof point was a genuine highlight. The main gap is that the CSM never led the ROI narrative — Tim did. For a call where adding 2-3 new seats was the stated objective, there was no deliberate positioning of the expansion case to Paul or the new AEs.", color=MID, size_pt=9),
        space_before=60, space_after=120
    ))

    # ── SECTION 5: NEXT STEPS & ACCOUNTABILITY ─────────────────────────────────
    parts.append(heading("5.  Next Steps & Accountability", level=1))
    parts.append(para(
        run("Score: ", bold=True, color=MID) + run("16 / 20", bold=True, color=CORAL, size_pt=12),
        space_before=0, space_after=60
    ))
    W6 = [5040, 4320]
    def ns_row(item, sym, fill=WHITE):
        return row(
            cell(run(item, color=DARK, size_pt=10), fill=fill, width=W6[0]),
            cell(run(sym, bold=True, size_pt=10, color=DARK), fill=fill, width=W6[1], align="center"),
        )
    parts.append(table(
        row(cell(run("ITEM", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=W6[0]),
            cell(run("STATUS", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=W6[1], align="center")),
        ns_row("Platform access — Oliver & Connor", "✅  Reset on call; confirmed working", fill=GREEN_LT),
        ns_row("Platform access — Johnny & Conrad", "✅  Chall contacted Tim post-call; both resolved", fill=GREEN_LT),
        ns_row("Formal written recap with action items sent", "✅  Structured email with Meeting Recap, Key Points, Action Items, and Next Steps", fill=GREEN_LT),
        ns_row("Action items split by party (Serve First vs Interceptly)", "✅  Clear ownership documented in post-call email", fill=GREEN_LT),
        ns_row("Individual AE sessions formally proposed", "✅  Written commitment to schedule 1-1 sessions per AE covering campaigns, messaging, and strategy", fill=GREEN_LT),
        ns_row("Competitors / existing clients / active company lists formally requested", "✅  Explicitly itemised in post-call email; awaiting client response", fill=GREEN_LT),
        ns_row("Campaign prospect assignment check committed", "✅  Alison confirmed in email; will verify AEs are on correct campaigns", fill=LIGHT_BG),
        ns_row("Messaging tone and cadence feedback loop established", "✅  AEs directed to review templates and share feedback; Alison to incorporate", fill=LIGHT_BG),
        ns_row("Individual session booked or email sent directly to Paul (new MD)", "❌  Post-call email addressed Tim only; Paul not named or contacted", fill=RED_LT),
        ns_row("Katie & Alex directly addressed post-call", "❌  Not mentioned in post-call email despite being the stated reason for the call", fill=RED_LT),
        ns_row("Individual sessions booked (dates confirmed)", "❌  Proposed but not yet scheduled", fill=YELLOW),
        ns_row("Next group review session scheduled", "❌  No date set for group follow-up", fill=YELLOW),
        total_width=9360, col_widths=W6
    ))
    parts.append(para(
        run("Score rationale: ", bold=True, color=MID, size_pt=9) +
        run("Strong post-call effort — the written recap with split action items, formally proposed individual sessions, and explicit blacklist request are all exactly right. Chall's post-call access resolution for Johnny and Conrad removes a blocker cleanly. The remaining gap: Paul, Katie, and Alex were not addressed personally. The email going to Tim only means the three most at-risk stakeholders received no direct communication after a call they attended.", color=MID, size_pt=9),
        space_before=60, space_after=120
    ))

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
        issue_row("1", "Paul (new MD) — not addressed in post-call email; no individual session proposed to him directly", "🔴 High", "Alison", "Send personal follow-up this week", fill=RED_LT),
        issue_row("2", "Katie & Alex — not mentioned in post-call email; call objective (evaluate for new seats) has no post-call momentum", "🔴 High", "Alison", "Contact individually this week"),
        issue_row("3", "Individual AE sessions proposed but not yet booked — no dates or calendar invites sent", "🔴 High", "Alison", "Book all 4 sessions this week", fill=RED_LT),
        issue_row("4", "Competitors + client + active company blacklist not yet received; campaigns may still target the wrong people", "🟡 Medium", "Tim (client)", "Chase before next campaign run", fill=YELLOW),
        issue_row("5", "Campaign messaging not personalised per AE — feedback loop established but content not yet updated", "🟡 Medium", "Alison", "Prior to HR campaign launch"),
        issue_row("6", "Stagecoach prospect showing in Oliver's campaign — Alison committed to checking assignment", "🟡 Medium", "Alison", "Confirm this week", fill=YELLOW),
        total_width=9360, col_widths=W7
    ))
    parts.append(spacer(10))

    # ── RISK FLAGS ────────────────────────────────────────────────────────────
    parts.append(heading("Risk Flags", level=1, color="92400E"))
    for title, detail in [
        ("Paul disengaged or unconvinced", "The new marketing director never saw the platform in action, never had a question answered, and was not addressed in the post-call email. If Paul is the decision-maker on adding seats, this is a serious gap with no current recovery plan."),
        ("Katie & Alex seat conversion at risk", "The stated purpose of the call was to evaluate adding Katie and Alex. They were introduced and then invisible — not engaged during the call, not mentioned in the follow-up. The individual sessions proposed in the email do not name them specifically, which reduces urgency."),
        ("Tim dependency", "The post-call email goes to Tim only. All action items are routed through Tim. If Tim is the only internal advocate and he becomes unavailable or deprioritises the tool, the account has no direct CSM-to-AE relationship to fall back on."),
        ("Cadence not personalised — adoption risk", "Oliver explicitly said the generic cadence is not how he communicates. The post-call email requests feedback but does not set a deadline or own the rewrite. If AEs review templates and do not feel represented, they will disengage from the inbox."),
        ("Rebranding confusion — login trust", "Password reset links were still pointing to Outreachly domains. Access is now resolved but the experience creates a credibility question for new users (Paul, Katie, Alex) who have not yet logged in."),
    ]:
        parts.append(bullet(detail, bold_prefix=title, color=DARK))
    parts.append(spacer(10))

    # ── WHAT WENT WELL ────────────────────────────────────────────────────────
    parts.append(heading("What Went Well", level=1, color="065F46"))
    for item in [
        "Strong opening intro — Alison's background and Interceptly's value proposition were clearly framed at the start.",
        "Oliver's Deloitte meeting moment — a concrete, live proof point that the platform works; the most compelling minute of the call.",
        "Comment approval workflow — Oliver's concern about AI commenting was immediately resolved; turned a potential objection into a product strength.",
        "Lead Intercept demo — the concept of tracking competitor activity landed well; Tim's enthusiasm validated the feature.",
        "Chall's post-call access resolution — Johnny and Conrad both reconnected to the platform after the call; clean, proactive follow-through.",
        "Post-call email quality — structured recap with Meeting Recap, Key Points, split Action Items, and individual session proposal. Clear ownership and professional tone.",
        "Blacklist process formalised — competitors, existing clients, and active company exclusions explicitly itemised and requested in writing for the first time.",
    ]:
        parts.append(bullet(item, color=DARK))
    parts.append(spacer(10))

    # ── RECOMMENDED NEXT STEPS ────────────────────────────────────────────────
    parts.append(heading("Recommended Next Steps", level=1, color=CORAL))
    steps = [
        ("This week", "Send a personal email directly to Paul: brief intro to Interceptly, reference Oliver's Deloitte meeting as a proof point, and invite him to his own 30-minute session. Do not route this through Tim."),
        ("This week", "Contact Katie and Alex directly (not via Tim) — confirm their access, acknowledge they are being onboarded, and book their individual sessions with dates."),
        ("This week", "Chase Tim for the three blacklists (competitors, existing clients, active companies) — campaigns are live now and may still be reaching the wrong people."),
        ("This week", "Book all individual AE sessions with calendar invites — proposed in the email but not yet confirmed with dates. Johnny's session should include an account-level catch-up since he missed the group call."),
        ("Before AE sessions", "Prepare AE-specific session structure: cadence review and rewrite, inbox management walkthrough, Lead Intercept signal assignment, and how to resume a paused campaign."),
        ("HR campaign launch", "Get Oliver's tone-of-voice feedback before writing HR campaign content — do not repeat the generic cadence mistake. Agree a review step before any new campaign goes live."),
    ]
    W8 = [1440, 7920]
    parts.append(table(
        *[row(
            cell(run(when, bold=True, color=CORAL, size_pt=9), fill=CORAL_LT if i % 2 == 0 else WHITE, width=W8[0]),
            cell(run(action, color=DARK, size_pt=10), fill=CORAL_LT if i % 2 == 0 else WHITE, width=W8[1]),
          ) for i, (when, action) in enumerate(steps)],
        total_width=9360, col_widths=W8
    ))
    parts.append(spacer(16))

    # ── FOOTER ────────────────────────────────────────────────────────────────
    parts.append(para(
        run("Scored against Interceptly CSM training call framework.  |  Health: 🟡 Yellow (50–74)  |  Package: Pipeline Builder — 5 seats  |  Revised 14 March 2026 (post-call email + access resolution)",
            color=MID, size_pt=8, italic=True),
        align="center", space_before=80, space_after=0, border_top="E5E7EB"
    ))

    return "\n".join(parts)


# ── SUPPORTING XML ─────────────────────────────────────────────────────────────
NUMBERING_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/><w:numFmt w:val="bullet"/>
      <w:lvlText w:val="&#x2022;"/><w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="360" w:hanging="240"/></w:pPr>
      <w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/></w:rPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>"""

DOCUMENT_XML = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
  <w:body>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080"
               w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
    {build_body()}
  </w:body>
</w:document>"""

STYLES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults><w:rPrDefault><w:rPr>
    <w:rFonts w:ascii="Inter" w:hAnsi="Inter" w:cs="Inter"/>
    <w:sz w:val="22"/><w:szCs w:val="22"/><w:color w:val="1F2937"/>
  </w:rPr></w:rPrDefault></w:docDefaults>
</w:styles>"""

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

with zipfile.ZipFile(OUTPUT, "w", zipfile.ZIP_DEFLATED) as zf:
    zf.writestr("[Content_Types].xml",          CONTENT_TYPES_XML)
    zf.writestr("_rels/.rels",                  DOTRELS_XML)
    zf.writestr("word/document.xml",            DOCUMENT_XML)
    zf.writestr("word/styles.xml",              STYLES_XML)
    zf.writestr("word/numbering.xml",           NUMBERING_XML)
    zf.writestr("word/_rels/document.xml.rels", RELS_XML)

print(f"Created: {OUTPUT}")

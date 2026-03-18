#!/usr/bin/env python3
"""Oliver Ayoub — Campaign Strategy Call scorecard. March 18 2026."""

import zipfile, os

OUTPUT = os.path.expanduser(
    "~/Documents/Claude/OliverAyoub_CampaignStrategy_Scorecard_Mar18.docx"
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

def bullet(text, bold_prefix=None, color=DARK):
    content = ""
    if bold_prefix:
        content += run(bold_prefix + " ", bold=True, color=color)
    content += run(text, color=color)
    return f"""<w:p>
    <w:pPr>
      <w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>
      <w:spacing w:before="0" w:after="60"/>
      <w:ind w:left="360" w:hanging="240"/>
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


def build_body():
    parts = []

    # ── HEADER ────────────────────────────────────────────────────────────────
    parts.append(para(
        run("INTERCEPTLY", bold=True, color=WHITE, size_pt=11, font="Gotham"),
        shading_color=CORAL, align="center", space_before=0, space_after=0
    ))
    parts.append(para(
        run("Customer Success Call Scorecard", bold=True, color=DARK, size_pt=20, font="Gotham"),
        align="center", space_before=160, space_after=40
    ))
    parts.append(para(
        run("Oliver Ayoub  x  Alison | Interceptly", color=MID, size_pt=13, font="Inter"),
        align="center", space_before=0, space_after=200, border_bottom=CORAL
    ))

    # ── META ──────────────────────────────────────────────────────────────────
    W = [1560, 3120, 1560, 3120]
    def meta_row(l1, v1, l2, v2):
        return row(
            cell(run(l1, bold=True, color=MID, size_pt=9), fill=LIGHT_BG, width=W[0]),
            cell(run(v1, color=DARK, size_pt=10), fill=WHITE, width=W[1]),
            cell(run(l2, bold=True, color=MID, size_pt=9), fill=LIGHT_BG, width=W[2]),
            cell(run(v2, color=DARK, size_pt=10), fill=WHITE, width=W[3]),
        )
    parts.append(table(
        meta_row("DATE",      "18 March 2026",                  "CSM",      "Alison"),
        meta_row("DURATION",  "14 minutes",                     "CUSTOMER", "Oliver Ayoub"),
        meta_row("CALL TYPE", "Campaign Strategy Session",      "ACCOUNT",  "Oliver Ayoub (AE — retail / hospitality ICP)"),
        meta_row("PACKAGE",   "Not specified",                  "SEGMENT",  "Active — CX + OPS Excellence campaigns live"),
        total_width=9360, col_widths=W
    ))
    parts.append(spacer(12))

    # ── OVERALL SCORE ─────────────────────────────────────────────────────────
    score_content = (
        run("Overall CS Call Health Score", bold=True, color=WHITE, size_pt=13, font="Gotham") +
        run("      ", color=WHITE) +
        run("70 / 100", bold=True, color=WHITE, size_pt=20, font="Gotham") +
        run("    ", color=WHITE) +
        run("🟡  YELLOW — Needs Attention", bold=True, color=WHITE, size_pt=12, font="Gotham")
    )
    parts.append(para(score_content, shading_color=CORAL, align="center",
                      space_before=80, space_after=80))
    parts.append(spacer(14))

    # ── CALL CONTEXT ──────────────────────────────────────────────────────────
    parts.append(heading("Call Context", level=1, color=MID))
    parts.append(para(
        run("Oliver is an AE (retail/hospitality ICP) with two active campaigns — CX and OPS Excellence. He called wanting a simple framework: how many touchpoints, what content format, and how to get a new bespoke HR campaign built and uploaded. Alison had proactively built a new campaign template ahead of the call. The session was primarily a campaign strategy walkthrough, but ran into a friction point when Oliver found the platform UI demo too complex. The call resolved with agreement on a next step (Alison to send a simple content template) and a target to upload next week.", color=MID, size_pt=10),
        space_before=0, space_after=120
    ))

    # ── SECTION 1: PRE-CALL PREPARATION & PROACTIVITY ─────────────────────────
    parts.append(heading("1.  Pre-Call Preparation & Proactivity", level=1))
    parts.append(para(
        run("Score: ", bold=True, color=MID) + run("17 / 20", bold=True, color=CORAL, size_pt=12),
        space_before=0, space_after=60
    ))
    W2 = [3960, 720, 4680]
    def prep_row(signal, icon, notes, fill=WHITE):
        return row(
            cell(run(signal, color=DARK, size_pt=10), fill=fill, width=W2[0]),
            cell(run(icon, size_pt=11), fill=fill, width=W2[1], align="center"),
            cell(run(notes, color=MID, size_pt=9, italic=True), fill=fill, width=W2[2]),
        )
    parts.append(table(
        row(cell(run("SIGNAL", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=W2[0]),
            cell(run("", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=W2[1]),
            cell(run("ASSESSMENT", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=W2[2])),
        prep_row("New campaign template built before the call", "✅", "Alison had proactively created the HR campaign workflow based on their last conversation — arrived with something to show, not just to discuss. Strongest pre-call action", fill=GREEN_LT),
        prep_row("Knew active campaign status (CX and OPS Excellence live)", "✅", "Immediately referenced which campaigns were live. No time wasted confirming basics", fill=GREEN_LT),
        prep_row("Clear call agenda set", "✅", "Alison led with a clear objective: walk through the new campaign flow and get Oliver's feedback on content direction. No ambiguity about what the session was for", fill=GREEN_LT),
        prep_row("Simple content template not prepared in advance", "⚠️", "Oliver asked for a template to fill in. Alison agreed to send one after the call — but a basic Google Doc template could have been prepared before the call. It was a predictable ask for a customer writing their own content", fill=YELLOW),
        total_width=9360, col_widths=W2
    ))
    parts.append(para(
        run("Score rationale: ", bold=True, color=MID, size_pt=9) +
        run("Excellent proactivity — arriving with a fully built campaign template is the mark of a CSM who anticipates rather than reacts. The only gap is that the simple content template (what Oliver actually needed to write messages) wasn't ready. Given Oliver confirmed on a prior call that he wanted to write his own content, this should have been prepared.", color=MID, size_pt=9),
        space_before=60, space_after=120
    ))

    # ── SECTION 2: CAMPAIGN KNOWLEDGE & CONSULTATION ─────────────────────────
    parts.append(heading("2.  Campaign Knowledge & Consultation", level=1))
    parts.append(para(
        run("Score: ", bold=True, color=MID) + run("17 / 25", bold=True, color=CORAL, size_pt=12),
        space_before=0, space_after=60
    ))
    W3 = [3600, 1200, 4560]
    def camp_row(topic, depth, notes, fill=WHITE):
        return row(
            cell(run(topic, color=DARK, size_pt=10), fill=fill, width=W3[0]),
            cell(run(depth, bold=True, size_pt=11, color=DARK), fill=fill, width=W3[1], align="center"),
            cell(run(notes, color=MID, size_pt=9, italic=True), fill=fill, width=W3[2]),
        )
    parts.append(table(
        row(cell(run("TOPIC", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=W3[0]),
            cell(run("DEPTH", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=W3[1], align="center"),
            cell(run("NOTES", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=W3[2])),
        camp_row("Campaign structure (11 touches, 8 days, dual-branch)", "✅ Full", "Comprehensive walkthrough of the workflow: two branches (open/closed profile), timing gaps, total touch count. Logic well explained"),
        camp_row("AI comment feature — activity-based timing", "✅ Full", "Correctly explained that AI comments post on the prospect's most recent activity regardless of time gap. Proactively raised the stale-post scenario and recommended an approval step", fill=LIGHT_BG),
        camp_row("Message sequencing philosophy (no sell early, CTA at touch 3-4)", "✅ Full", "Good consulting-level explanation: early messages spark interest, CTA introduced mid-sequence. 'Earning the right to get a response' framing was effective"),
        camp_row("Replies peak at touch 3-4 — evidence cited", "✅ Full", "Data-backed rationale for sequence length. Pushed back professionally when Oliver wanted to cut to 7-8 touches", fill=LIGHT_BG),
        camp_row("Platform UI demo — complexity vs. customer mental model", "❌ Mismatch", "Oliver said 'none of it makes sense to me' and 'it's a whole load of boxes'. The screen-share was a live workflow view, which maps to platform logic not to Oliver's question: what messages go out, when, and what do I write? The visual demo was the wrong format for this customer", fill=RED_LT),
        camp_row("Template format guidance (first response to Oliver's ask)", "⚠️ Vague", "'You can use the resources section to create templates' — this sent Oliver to figure it out himself. He had to push a second time ('is it all right if you send me a template?') before Alison committed to sending something clear", fill=YELLOW),
        camp_row("Evidence-based pushback on sequence reduction", "✅ Full", "When Oliver wanted to cut to 7-8 touches, Alison pushed back citing proven results across industries. Oliver immediately accepted and deferred. Demonstrates credibility", fill=LIGHT_BG),
        camp_row("Retail/hospitality ICP context captured", "✅ Partial", "Alison noted the ICP at the close. Was not probed earlier — Oliver volunteered it. An earlier check-in on ICP could have shaped the template recommendation", fill=LIGHT_BG),
        total_width=9360, col_widths=W3
    ))
    parts.append(para(
        run("Score rationale: ", bold=True, color=MID, size_pt=9) +
        run("Strong campaign knowledge demonstrated: sequence logic, AI comment mechanics, evidence-based sequencing rationale. The significant gap is the demo approach. Sharing a live workflow UI with a customer who said he just needed to know 'what touch points are going out, the time between them, and what content to give' created unnecessary friction. The first response to Oliver's template request was also a deflection — Alison recovered well but only after Oliver pushed twice.", color=MID, size_pt=9),
        space_before=60, space_after=120
    ))

    # ── SECTION 3: CUSTOMER UNDERSTANDING & NEEDS ALIGNMENT ──────────────────
    parts.append(heading("3.  Customer Understanding & Needs Alignment", level=1))
    parts.append(para(
        run("Score: ", bold=True, color=MID) + run("13 / 20", bold=True, color=CORAL, size_pt=12),
        space_before=0, space_after=60
    ))
    W4 = [4200, 720, 4440]
    def cu_row(signal, icon, notes, fill=WHITE):
        return row(
            cell(run(signal, color=DARK, size_pt=10), fill=fill, width=W4[0]),
            cell(run(icon, size_pt=11), fill=fill, width=W4[1], align="center"),
            cell(run(notes, color=MID, size_pt=9, italic=True), fill=fill, width=W4[2]),
        )
    parts.append(table(
        row(cell(run("SIGNAL", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=W4[0]),
            cell(run("", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=W4[1]),
            cell(run("ASSESSMENT", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=W4[2])),
        cu_row("Oliver's core ask understood at close", "✅", "By the end of the call, Alison accurately captured what Oliver needed: a simple template, him writing the content, Alison providing feedback and uploading. Correct resolution", fill=GREEN_LT),
        cu_row("Bespoke vs. shared campaign — clarifying question asked", "✅", "Alison proactively asked whether the HR campaign was bespoke to Oliver or shared across all AEs. Important operational question that changed the approach. Well timed", fill=GREEN_LT),
        cu_row("Oliver's ask was clear from minute one — demo wasn't the right response", "❌", "Oliver's opening ask was explicit: 'could you provide me with the framework... I'll go get them created'. He was asking for a format/template, not a platform walkthrough. Alison defaulted to a screen-share demo rather than leading with the simpler answer", fill=RED_LT),
        cu_row("Oliver's communication style / personality not explored", "⚠️", "Oliver volunteered that he's 'quirky', 'more out there', 'not afraid to be high risk'. This is highly relevant for tailoring the content template — Alison should have drawn this out proactively to shape the template she's building", fill=YELLOW),
        cu_row("Approval step recommendation aligned to Oliver's risk tolerance", "✅", "Alison recommended adding an approval step for AI comments — conservative but correct. Oliver accepted immediately, suggesting he appreciated the guidance"),
        cu_row("Oliver's time constraints acknowledged and respected", "✅", "Oliver said 'I've got a nip off' — Alison wrapped cleanly. The call didn't overrun despite the content volume", fill=LIGHT_BG),
        total_width=9360, col_widths=W4
    ))
    parts.append(para(
        run("Score rationale: ", bold=True, color=MID, size_pt=9) +
        run("The call resolved correctly but the path there had a significant mismatch: Oliver wanted a simple template framework from the start, and Alison responded with a platform screen-share. Oliver had to say 'none of it makes sense to me' before the pivot happened. The save was good — but a CSM who reads the customer earlier would have skipped the friction entirely. Oliver's personality/style — which he volunteered unprompted — is critical input for the content template; Alison should have mined this.", color=MID, size_pt=9),
        space_before=60, space_after=120
    ))

    # ── SECTION 4: ACTION ITEMS & ACCOUNTABILITY ──────────────────────────────
    parts.append(heading("4.  Action Items & Accountability", level=1))
    parts.append(para(
        run("Score: ", bold=True, color=MID) + run("12 / 20", bold=True, color=CORAL, size_pt=12),
        space_before=0, space_after=60
    ))
    W5 = [5040, 4320]
    def ai_row(item, sym, fill=WHITE):
        return row(
            cell(run(item, color=DARK, size_pt=10), fill=fill, width=W5[0]),
            cell(run(sym, bold=True, size_pt=10, color=DARK), fill=fill, width=W5[1], align="center"),
        )
    parts.append(table(
        row(cell(run("ITEM", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=W5[0]),
            cell(run("STATUS", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=W5[1], align="center")),
        ai_row("Alison to send content template (Google Doc format) with touchpoints, timing, and content slots", "✅  Agreed on call — delivery not confirmed", fill=GREEN_LT),
        ai_row("Template tailored to retail/hospitality ICP with Interceptly's best-performing sequence", "✅  Verbally agreed", fill=GREEN_LT),
        ai_row("Timeline for sending template (same day? by end of week?)", "❌  Not confirmed. Oliver expects it 'this weekend' — no explicit commitment on when Alison will send", fill=RED_LT),
        ai_row("Oliver to review template over the weekend and write content", "✅  Oliver self-committed: 'I'll have a good look at it this weekend'", fill=LIGHT_BG),
        ai_row("Target to upload campaign content next week", "⚠️  Agreed directionally but no specific date set — 'next week' is broad", fill=YELLOW),
        ai_row("Follow-up call booked to review Oliver's content before upload", "❌  No next session scheduled. Content feedback will presumably happen async — not confirmed", fill=RED_LT),
        ai_row("Alison to provide feedback on Oliver's content draft", "⚠️  Mentioned ('you can decide whether or not you can have a say on it') but no formal commitment on turnaround or process", fill=YELLOW),
        total_width=9360, col_widths=W5
    ))
    parts.append(para(
        run("Score rationale: ", bold=True, color=MID, size_pt=9) +
        run("The core action item is clear (Alison sends template, Oliver writes content) but the execution is underdefined. No date set for template delivery, no follow-up session booked, and the content review process is vague ('you can have a say on it'). For a customer who is self-motivated and moving fast, these gaps are unlikely to derail the week — but they create opportunities for the momentum to stall if the template doesn't land by Friday.", color=MID, size_pt=9),
        space_before=60, space_after=120
    ))

    # ── SECTION 5: COMMUNICATION & RAPPORT ───────────────────────────────────
    parts.append(heading("5.  Communication & Rapport", level=1))
    parts.append(para(
        run("Score: ", bold=True, color=MID) + run("11 / 15", bold=True, color=CORAL, size_pt=12),
        space_before=0, space_after=60
    ))
    W6 = [4200, 720, 4440]
    def rr_row(signal, icon, notes, fill=WHITE):
        return row(
            cell(run(signal, color=DARK, size_pt=10), fill=fill, width=W6[0]),
            cell(run(icon, size_pt=11), fill=fill, width=W6[1], align="center"),
            cell(run(notes, color=MID, size_pt=9, italic=True), fill=fill, width=W6[2]),
        )
    parts.append(table(
        row(cell(run("SIGNAL", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=W6[0]),
            cell(run("", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=W6[1]),
            cell(run("ASSESSMENT", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=W6[2])),
        rr_row("Natural opening rapport — location, humour", "✅", "Quick, easy rapport: Colombia sun hat / 'wet weather' banter. Set a relaxed tone in under a minute. Oliver was at ease immediately", fill=GREEN_LT),
        rr_row("Listened without defensiveness to Oliver's 'none of it makes sense'", "✅", "When Oliver gave blunt feedback on the UI, Alison pivoted cleanly without explaining herself or being defensive. Good composure", fill=GREEN_LT),
        rr_row("Handled evidence-based pushback confidently", "✅", "When Oliver wanted fewer touches, Alison pushed back politely but with conviction ('we've tested this with many industries, proven record'). Oliver deferred immediately — sign of a trusted advisor relationship", fill=GREEN_LT),
        rr_row("Oliver had to repeat himself twice to get a straight answer on the template", "⚠️", "First answer was 'you can use the resources section'. Oliver said 'none of it makes sense' and had to ask again more directly before getting 'I'll send you a template'. One clear answer earlier would have been stronger", fill=YELLOW),
        rr_row("Pace and brevity appropriate for 14-minute session", "✅", "Alison matched Oliver's pace — direct, no fluff. Clean close when Oliver said he needed to leave", fill=LIGHT_BG),
        rr_row("Oliver's personality/style could have been surfaced and mirrored", "⚠️", "Oliver described himself as 'quirky', 'high risk', 'out there'. Alison didn't pick up on this or adapt her tone accordingly. A line like 'perfect — let me build you something that matches your style' would have reinforced the personal connection", fill=YELLOW),
        total_width=9360, col_widths=W6
    ))
    parts.append(spacer(10))

    # ── OPEN ISSUES ───────────────────────────────────────────────────────────
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
        issue_row("1", "Send Oliver the content template (Google Doc) — touchpoints, timing gaps, and content slots — tailored to retail/hospitality ICP with 8-11 touch best-practice sequence", "🔴 High", "Alison", "Today (Mar 18)", fill=RED_LT),
        issue_row("2", "Include brief note on Oliver's personality/tone in template cover note — he writes 'quirky, high-risk, personable'. Template framing should reflect this", "🟡 Medium", "Alison", "With template"),
        issue_row("3", "Book a 20-min content review call for next week to review Oliver's draft before uploading", "🟡 Medium", "Alison", "Book this week", fill=YELLOW),
        issue_row("4", "Confirm whether Oliver's HR campaign uses the same leads as CX/OPS or a separate list — not confirmed on the call", "🟡 Medium", "Alison", "In template follow-up"),
        total_width=9360, col_widths=W7
    ))
    parts.append(spacer(10))

    # ── WHAT WENT WELL ────────────────────────────────────────────────────────
    parts.append(heading("What Went Well", level=1, color="065F46"))
    for item in [
        "Proactively built the new HR campaign template before the call — arrived with something to show, not just to discuss. This is the standard all pre-call prep should be held to.",
        "AI comment feature explanation was excellent: correctly flagged the stale-post scenario, recommended an approval step, and explained the timing logic accurately. Oliver accepted immediately.",
        "Evidence-based pushback on sequence reduction was confident and credible. Oliver deferred right away — a sign that Alison has earned authority in this relationship.",
        "Clean, fast rapport at the open. Oliver was relaxed and direct throughout — a good indicator of a healthy working relationship.",
        "Pivoted cleanly when Oliver gave blunt feedback on the UI demo. No defensiveness, no over-explanation — just adapted and gave him what he needed.",
        "Correctly identified that Oliver's campaign is bespoke and separate from other AEs — an important operational question that changed the setup plan.",
    ]:
        parts.append(bullet(item, color=DARK))
    parts.append(spacer(10))

    # ── RECOMMENDED NEXT STEPS ────────────────────────────────────────────────
    parts.append(heading("Recommended Next Steps", level=1, color=CORAL))
    steps = [
        ("Today", "Send Oliver the content template as a clean Google Doc. Include: step number, channel (email / LinkedIn message / AI comment etc.), timing gap, character/word guidance, and a blank content field. Keep it to one row per touch. Attach a short note: 'Based on your style (direct, high-risk, personable) — feel free to be yourself in the content. I'll review before we upload.'"),
        ("Today", "Include the best-performing sequence structure for retail/hospitality in the template — pre-filled timing gaps and touch logic so Oliver only needs to fill in the words."),
        ("This week", "Book a 20-minute content review slot for next week — aim for mid-week so there's time to adjust before upload. Send the invite alongside the template."),
        ("Next week", "Review Oliver's content draft before uploading. Check: messaging logic (no sell in early touches), CTA placement (touch 3-4), tone consistency, and LinkedIn character limits."),
        ("Ongoing", "For Oliver's future campaigns: lead with a 1-page visual summary (step, channel, timing) before the platform screen share. Customers with a sales background respond better to a sequence map than a live workflow UI."),
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
        run("Scored against Interceptly CSM success call framework.  |  Health: 🟡 Yellow (50-74)  |  Call type: Campaign Strategy Session  |  Generated 19 March 2026",
            color=MID, size_pt=8, italic=True),
        align="center", space_before=80, space_after=0, border_top="E5E7EB"
    ))

    return "\n".join(parts)


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
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
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

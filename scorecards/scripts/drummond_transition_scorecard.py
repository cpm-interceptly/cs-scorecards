#!/usr/bin/env python3
"""Drummond Group LLC — Account Transition Call scorecard. March 16 2026."""

import zipfile, os

OUTPUT = os.path.expanduser(
    "~/Documents/Claude/Drummond_Transition_Scorecard_Mar16.docx"
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
        run("Drummond Group LLC  x  Alison | Interceptly", color=MID, size_pt=13, font="Inter"),
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
        meta_row("DATE",      "16 March 2026",                       "CSM",      "Alison"),
        meta_row("DURATION",  "28 minutes",                          "CUSTOMER", "Drummond Group LLC"),
        meta_row("CALL TYPE", "Account Transition — User Handover",  "CONTACTS", "Michael Henry, Lucy Railton, Peyton Kirkley, Carl LB"),
        meta_row("PACKAGE",   "Lead Intercept + Campaign Management", "SEGMENT", "SMB — Active / Multi-Campaign"),
        total_width=9360, col_widths=W
    ))
    parts.append(spacer(12))

    # ── ATTENDEES ─────────────────────────────────────────────────────────────
    parts.append(heading("Attendees", level=1, color=MID))
    WA = [3000, 2160, 4200]
    parts.append(table(
        row(cell(run("NAME", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=WA[0]),
            cell(run("ORG", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=WA[1]),
            cell(run("ROLE / CONTEXT", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=WA[2])),
        row(cell(run("Michael Henry", color=DARK, size_pt=10), fill=WHITE, width=WA[0]),
            cell(run("Drummond Group LLC", color=DARK, size_pt=10), fill=WHITE, width=WA[1]),
            cell(run("Primary decision-maker. Owns outreach strategy. Previously managed Sunitha.", color=MID, size_pt=9, italic=True), fill=WHITE, width=WA[2])),
        row(cell(run("Lucy Railton", color=DARK, size_pt=10), fill=LIGHT_BG, width=WA[0]),
            cell(run("Drummond Group LLC", color=DARK, size_pt=10), fill=LIGHT_BG, width=WA[1]),
            cell(run("Senior stakeholder. Active in strategic decisions (campaign sequencing, ICP).", color=MID, size_pt=9, italic=True), fill=LIGHT_BG, width=WA[2])),
        row(cell(run("Peyton Kirkley", color=DARK, size_pt=10), fill=WHITE, width=WA[0]),
            cell(run("Drummond Group LLC", color=DARK, size_pt=10), fill=WHITE, width=WA[1]),
            cell(run("New outreach rep replacing Sunitha. Experienced with voicemail drops. First contact with Interceptly platform.", color=MID, size_pt=9, italic=True), fill=WHITE, width=WA[2])),
        row(cell(run("Carl LB", color=DARK, size_pt=10), fill=LIGHT_BG, width=WA[0]),
            cell(run("Drummond Group LLC", color=DARK, size_pt=10), fill=LIGHT_BG, width=WA[1]),
            cell(run("Internal ops/campaign owner. Holds the PCI campaign brief. Liaising with Alison on campaign launch.", color=MID, size_pt=9, italic=True), fill=LIGHT_BG, width=WA[2])),
        total_width=9360, col_widths=WA
    ))
    parts.append(spacer(12))

    # ── OVERALL SCORE ─────────────────────────────────────────────────────────
    score_content = (
        run("Overall CS Call Health Score", bold=True, color=WHITE, size_pt=13, font="Gotham") +
        run("      ", color=WHITE) +
        run("72 / 100", bold=True, color=WHITE, size_pt=20, font="Gotham") +
        run("    ", color=WHITE) +
        run("🟡  YELLOW — Needs Attention", bold=True, color=WHITE, size_pt=12, font="Gotham")
    )
    parts.append(para(score_content, shading_color=CORAL, align="center",
                      space_before=80, space_after=80))
    parts.append(spacer(14))

    # ── CALL CONTEXT ──────────────────────────────────────────────────────────
    parts.append(heading("Call Context", level=1, color=MID))
    parts.append(para(
        run("Drummond Group LLC is transitioning their outreach rep from Sunitha (departing) to Peyton Kirkley (incoming). Two active campaigns are in play: a mental health care clinic outreach (email + LinkedIn + voicemail drops) and a Lead Intercept program (LinkedIn only, weekly Drada-engagement signals). A third PCI campaign is on hold pending internal alignment. This call was a stakeholder alignment session to introduce Peyton, agree the transition plan, and unblock the Lead Intercept transfer.", color=MID, size_pt=10),
        space_before=0, space_after=120
    ))

    # ── SECTION 1: ACCOUNT TRANSITION MANAGEMENT ──────────────────────────────
    parts.append(heading("1.  Account Transition Management", level=1))
    parts.append(para(
        run("Score: ", bold=True, color=MID) + run("20 / 25", bold=True, color=CORAL, size_pt=12),
        space_before=0, space_after=60
    ))
    W2 = [3960, 720, 4680]
    def tr_row(signal, icon, notes, fill=WHITE):
        return row(
            cell(run(signal, color=DARK, size_pt=10), fill=fill, width=W2[0]),
            cell(run(icon, size_pt=11), fill=fill, width=W2[1], align="center"),
            cell(run(notes, color=MID, size_pt=9, italic=True), fill=fill, width=W2[2]),
        )
    parts.append(table(
        row(cell(run("SIGNAL", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=W2[0]),
            cell(run("", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=W2[1]),
            cell(run("ASSESSMENT", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=W2[2])),
        tr_row("Clear transition sequence established", "✅", "Prioritised Lead Intercept (LinkedIn only — no warm-up needed) immediately; paused HIPAA and email campaigns pending warm-up; PCI on hold pending Carl/Peyton alignment. Logical and risk-aware sequencing", fill=GREEN_LT),
        tr_row("Data export risk flagged proactively", "✅", "Alison raised the need to export Sunitha's prospect data before removing her — before Michael realised it was needed. Strongest moment of the call. Identified what's included (labels, tags, status) and what isn't (conversation history)", fill=GREEN_LT),
        tr_row("Workspace duplication plan clear", "✅", "Confirmed campaigns will be duplicated into Peyton's new workspace with updated sender/copy. Data saved before swap. Post-call email reinforced this", fill=GREEN_LT),
        tr_row("Bi-weekly training cadence proposed", "✅", "Proactively suggested continuing Sunitha's bi-weekly training schedule with Peyton — continuity of support structure", fill=GREEN_LT),
        tr_row("Export ownership ambiguous during call", "⚠️", "'I need to confirm with the team' — Alison hesitated on whether she could do the export herself. Michael ended up doing it himself after she walked him through the process. Ownership should have been clearer", fill=YELLOW),
        tr_row("Same-day Lead Intercept transfer committed", "✅", "Committed to completing the switch before end of day so no further days are lost (campaigns had been cold since Wednesday)", fill=GREEN_LT),
        total_width=9360, col_widths=W2
    ))
    parts.append(para(
        run("Score rationale: ", bold=True, color=MID, size_pt=9) +
        run("Strong transition planning — proactively caught the data export risk before the customer did, established a logical campaign sequencing plan, and committed to a same-day transfer for the priority campaign. The only gap is some hesitancy on export ownership ('I need to confirm') which created brief ambiguity that Michael resolved himself.", color=MID, size_pt=9),
        space_before=60, space_after=120
    ))

    # ── SECTION 2: PLATFORM KNOWLEDGE & FEATURE GUIDANCE ─────────────────────
    parts.append(heading("2.  Platform Knowledge & Feature Guidance", level=1))
    parts.append(para(
        run("Score: ", bold=True, color=MID) + run("14 / 20", bold=True, color=CORAL, size_pt=12),
        space_before=0, space_after=60
    ))
    W3 = [3600, 1200, 4560]
    def feat_row(topic, depth, notes, fill=WHITE):
        return row(
            cell(run(topic, color=DARK, size_pt=10), fill=fill, width=W3[0]),
            cell(run(depth, bold=True, size_pt=11, color=DARK), fill=fill, width=W3[1], align="center"),
            cell(run(notes, color=MID, size_pt=9, italic=True), fill=fill, width=W3[2]),
        )
    parts.append(table(
        row(cell(run("TOPIC", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=W3[0]),
            cell(run("DEPTH", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=W3[1], align="center"),
            cell(run("NOTES", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=W3[2])),
        feat_row("Duplicate lead in Lead Intercept explained", "✅ Full", "Accurate technical explanation: person appeared twice because they interacted with two tracked profiles. Different IDs triggered dual entry. Confirmed internal fix applied"),
        feat_row("Email warm-up requirements", "✅ Full", "Correctly identified: primary email (warm), 4-5 aliases (not warm, just created). LinkedIn-only suggested as interim — good risk management", fill=LIGHT_BG),
        feat_row("Prospect CSV export — contents and process", "✅ Full", "Explained process (prospect section), what's included (labels, tags, status, all activity), what's excluded (conversation history). Walked Michael through in real-time"),
        feat_row("Voicemail drop — initial recommendation", "⚠️ Flip", "Initially recommended against voicemail for PCI campaign without strong justification. When Michael pushed back, agreed to add 2 drops mid-sequence. Reversal was reasonable but the initial 'no' lacked conviction and created unnecessary friction", fill=YELLOW),
        feat_row("Data skewing from untracked voicemail activity", "✅ Full", "Correctly explained that off-platform calls aren't tracked, which skews engagement data. Framed clearly as a data awareness issue, not a blocker", fill=LIGHT_BG),
        feat_row("Conversation history not included in export", "✅ Full", "Proactively flagged that conversation history stays with Sunitha's inbox — important data loss awareness before the swap"),
        feat_row("Lead Intercept weekly workflow mechanics", "✅ Full", "Understood and confirmed the Wednesday population process, signal assignment, and campaign add flow. Did not over-explain", fill=LIGHT_BG),
        feat_row("Voicemail re-recording process for Peyton", "⚠️ Partial", "Mentioned Peyton would need to re-record the HIPAA voicemail drop, and that she'd walk him through it at training — but no specifics given on the platform step or timing. Post-call email omitted this entirely", fill=YELLOW),
        total_width=9360, col_widths=W3
    ))
    parts.append(para(
        run("Score rationale: ", bold=True, color=MID, size_pt=9) +
        run("Solid technical knowledge demonstrated across export process, email warm-up, and Lead Intercept mechanics. Two gaps: the initial voicemail drop reversal undermined confidence, and the voicemail re-recording process was left vague and missing from the post-call email — a task Peyton needs to action before HIPAA can relaunch.", color=MID, size_pt=9),
        space_before=60, space_after=120
    ))

    # ── SECTION 3: STAKEHOLDER ENGAGEMENT ────────────────────────────────────
    parts.append(heading("3.  Stakeholder Engagement", level=1))
    parts.append(para(
        run("Score: ", bold=True, color=MID) + run("13 / 20", bold=True, color=CORAL, size_pt=12),
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
        sh_row("Michael — primary relationship managed well", "✅", "Strong existing relationship evident. Alison deferred to Michael to introduce the context to Peyton — smart move that reinforced Michael's authority", fill=GREEN_LT),
        sh_row("Lucy — strategic contributions picked up", "✅", "Lucy's campaign sequencing suggestion (voicemail first, then email) was acknowledged and built on. Alison looped Carl in correctly when Lucy raised the PCI question", fill=GREEN_LT),
        sh_row("Peyton — proactive onboarding to platform early in call", "❌", "Peyton said almost nothing for the first 12 minutes beyond pleasantries. Alison didn't draw him into the conversation or ask about his background/experience until near the end. For a handover call where Peyton is the new primary user, this is a missed opportunity", fill=RED_LT),
        sh_row("Carl — included appropriately when relevant", "⚠️", "Carl only spoke when Lucy directly asked him a question (at 5:26). Alison didn't proactively address him or confirm his readiness. He's a key stakeholder for the PCI campaign", fill=YELLOW),
        sh_row("Voicemail drop discussion managed well", "✅", "When Michael pushed back on the voicemail recommendation, Alison handled it gracefully — validated his instinct, asked about his experience, adapted the strategy. Good collaborative problem-solving", fill=GREEN_LT),
        sh_row("Peyton's voicemail drop experience acknowledged and used", "✅", "When Peyton shared his experience (voicemail + follow-up email), Lucy immediately used it to shape the campaign strategy. Alison didn't pick up on this and redirect it — Lucy did the work", fill=YELLOW),
        sh_row("Closing — Peyton welcomed and supported", "✅", "Strong close with Peyton: 'feel free to reach out at any point, I always try to respond as soon as I can' — appropriate for a new contact who will need frequent support early", fill=GREEN_LT),
        total_width=9360, col_widths=W4
    ))
    parts.append(para(
        run("Score rationale: ", bold=True, color=MID, size_pt=9) +
        run("Strong management of the primary relationship (Michael and Lucy). The significant gap is Peyton — the incoming user and the person who matters most for platform adoption going forward — was essentially a passenger for the first half of the call. Alison should have drawn him out earlier: his experience, his current tools, his confidence level with outreach platforms. A user who feels engaged from the start is far more likely to self-serve and succeed.", color=MID, size_pt=9),
        space_before=60, space_after=120
    ))

    # ── SECTION 4: ACTION ITEMS & ACCOUNTABILITY ──────────────────────────────
    parts.append(heading("4.  Action Items & Accountability", level=1))
    parts.append(para(
        run("Score: ", bold=True, color=MID) + run("15 / 20", bold=True, color=CORAL, size_pt=12),
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
        ai_row("Login links sent to Peyton", "✅  Confirmed in post-call email — sent same day", fill=GREEN_LT),
        ai_row("Platform walkthrough training session booked", "✅  Booked on the call — tomorrow 1pm Eastern", fill=GREEN_LT),
        ai_row("Lead Intercept campaign transferred to Peyton's workspace", "✅  Committed before end of day. Post-call email confirmed plan", fill=GREEN_LT),
        ai_row("HIPAA and email campaigns paused pending warm-up", "✅  Confirmed in call and post-call email", fill=GREEN_LT),
        ai_row("Prospect data export (replied prospects from Sunitha's workspace)", "⚠️  Alison walked Michael through how to do it himself. Ownership stayed with Michael — not confirmed as a joint task or Alison-owned", fill=YELLOW),
        ai_row("Carl + Peyton to align on PCI campaign sequence by Monday", "⚠️  Agreed on call, but not confirmed in post-call email — no paper trail for this deadline", fill=YELLOW),
        ai_row("Voicemail re-recording for HIPAA campaign (Peyton)", "❌  Mentioned verbally but no process steps given. Absent from post-call email entirely", fill=RED_LT),
        ai_row("Bi-weekly training cadence confirmed with Peyton", "❌  Proposed on the call but not confirmed in post-call email — no calendar invite or written commitment", fill=YELLOW),
        ai_row("Post-call message sent to Michael", "✅  Sent same day — concise, covered key points", fill=GREEN_LT),
        total_width=9360, col_widths=W5
    ))
    parts.append(para(
        run("Score rationale: ", bold=True, color=MID, size_pt=9) +
        run("Core commitments were made and followed through (login links, training booked, campaigns paused, Lead Intercept transfer). The post-call email covered the essentials clearly. The gaps are three items that dropped off: voicemail re-recording (critical for HIPAA relaunch), the PCI Monday deadline (no written confirmation), and the bi-weekly cadence (proposed but not confirmed in writing). For a transition with a new user, the written trail matters — Peyton needs clarity on what to expect.", color=MID, size_pt=9),
        space_before=60, space_after=120
    ))

    # ── SECTION 5: CUSTOMER RELATIONSHIP & COMMUNICATION ──────────────────────
    parts.append(heading("5.  Customer Relationship & Communication", level=1))
    parts.append(para(
        run("Score: ", bold=True, color=MID) + run("10 / 15", bold=True, color=CORAL, size_pt=12),
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
        rr_row("Trusted advisor dynamic with Michael and Lucy", "✅", "Michael deferred to Alison on campaign decisions; Lucy engaged openly. Strong existing trust base", fill=GREEN_LT),
        rr_row("Peyton warmly welcomed and supported at close", "✅", "Strong closing message — open availability, proactive support commitment. Good first impression for the new primary user", fill=GREEN_LT),
        rr_row("Platform described as 'Outreachly' in Fathom transcript", "⚠️", "Fathom tagged Alison as 'Alison | Outreachly (Outreachly AI)' — a legacy branding alias. Worth ensuring the platform name is consistently Interceptly in all external-facing comms", fill=YELLOW),
        rr_row("Voicemail drop reversal — brief credibility dip", "⚠️", "Recommending no voicemail drops, then agreeing to include them when challenged, creates a perception that recommendations aren't fully conviction-driven. Minor but worth noting as a pattern", fill=YELLOW),
        rr_row("Michael's enthusiasm about platform results reinforced", "✅", "Michael voluntarily praised Interceptly's consolidated dashboard and LinkedIn engagement results vs. prior year — Alison should have explicitly acknowledged this and linked it to the value of continuing with Peyton. Missed a quick win", fill=YELLOW),
        rr_row("Post-call email tone — professional and clear", "✅", "Good tone, correctly addressed to Michael, referenced the key decisions. Appropriate length for a transition update", fill=GREEN_LT),
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
        issue_row("1", "Voicemail re-recording for HIPAA campaign — Peyton must record in his own voice before HIPAA can relaunch. Walk Peyton through platform step at training session", "🔴 High", "Alison + Peyton", "Training session (Mar 17)", fill=RED_LT),
        issue_row("2", "Peyton's email aliases (4-5) need warm-up monitoring — confirm warm-up start date and estimated readiness for HIPAA and PCI email campaigns", "🔴 High", "Alison", "Confirm this week"),
        issue_row("3", "Carl + Peyton PCI campaign alignment — workflow brief due by Monday. Alison needs 48hr notice to prep strategy. No written confirmation of this deadline", "🟡 Medium", "Carl + Peyton → Alison", "Monday (Mar 23)", fill=YELLOW),
        issue_row("4", "Bi-weekly training cadence — confirm in writing with Peyton (recurring calendar invite)", "🟡 Medium", "Alison", "This week"),
        issue_row("5", "Sunitha's LinkedIn replied prospect export — confirm Michael completed this before workspace swap. If not, window to retrieve may have closed", "🔴 High", "Alison (confirm)", "Immediate", fill=RED_LT),
        total_width=9360, col_widths=W7
    ))
    parts.append(spacer(10))

    # ── WHAT WENT WELL ────────────────────────────────────────────────────────
    parts.append(heading("What Went Well", level=1, color="065F46"))
    for item in [
        "Proactively flagged the data export risk before Michael realised — this was the most valuable CSM intervention of the call. Without Alison raising it, prospect data could have been lost at the workspace swap.",
        "Clear, logical transition sequence: Lead Intercept first (LinkedIn only, no warm-up), HIPAA/email paused, PCI held. Reduced risk of email deliverability issues for Peyton's new accounts.",
        "Duplicate lead technical explanation was accurate and reassuring — confirmed an internal fix had already been applied. Built confidence in the platform's reliability.",
        "Training session booked in real time on the call — no back-and-forth needed post-call.",
        "Strong relationship with Michael and Lucy evident — they trust Alison's guidance and openly defer to her on tactical decisions.",
        "Warm, professional close with Peyton — 'feel free to reach out at any point' sets the right tone for a new user who will need support in the first few weeks.",
    ]:
        parts.append(bullet(item, color=DARK))
    parts.append(spacer(10))

    # ── RECOMMENDED NEXT STEPS ────────────────────────────────────────────────
    parts.append(heading("Recommended Next Steps", level=1, color=CORAL))
    steps = [
        ("Today", "Confirm Peyton successfully logged in and account connected. If any issues, resolve before end of day."),
        ("Today", "Confirm Michael completed the Sunitha LinkedIn replied prospect export before the workspace swap was finalised. If the window has closed, escalate to the team."),
        ("Mar 17", "At Peyton's training session: walk through the platform fully, show the voicemail drop recording step, and get the HIPAA voicemail recorded in his voice."),
        ("Mar 17", "Send Peyton a written summary after his training session: what to do in the platform, how Lead Intercept works on Wednesdays, and the voicemail process."),
        ("Mar 17", "Send Peyton a recurring calendar invite for bi-weekly check-ins (continuing the cadence from Sunitha's onboarding)."),
        ("Mar 23", "Chase Carl and Peyton for PCI campaign workflow brief. Give 48hr buffer before prep start — aim for brief by Monday to launch strategy session mid-week."),
        ("Ongoing", "Monitor Peyton's 4-5 alias email warm-up. Confirm readiness date and update Michael/Lucy with an estimated relaunch timeline for HIPAA and email campaigns."),
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
        run("Scored against Interceptly CSM success call framework.  |  Health: 🟡 Yellow (50-74)  |  Call type: Account Transition (user handover)  |  Generated 17 March 2026",
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

#!/usr/bin/env python3
"""Tyson Schroeder — Interceptly Success Call scorecard. March 17 2026."""

import zipfile, os

OUTPUT = os.path.expanduser(
    "~/Documents/Claude/Tyson_SuccessCall_Scorecard_Mar17.docx"
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
        run("Tyson Schroeder  x  Chall | Interceptly", color=MID, size_pt=13, font="Inter"),
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
        meta_row("DATE",      "17 March 2026",              "CSM",       "Chall"),
        meta_row("DURATION",  "64 minutes",                 "CUSTOMER",  "Tyson Schroeder"),
        meta_row("CALL TYPE", "Customer Success Check-in",  "ACCOUNT",   "Interceptly (Tyson's business)"),
        meta_row("PACKAGE",   "Engage AI (no recurring meetings)", "SEGMENT", "SMB — Active / Engaged"),
        total_width=9360, col_widths=W
    ))
    parts.append(spacer(12))

    # ── OVERALL SCORE ─────────────────────────────────────────────────────────
    score_content = (
        run("Overall CS Call Health Score  (Revised — post-call message reviewed)", bold=True, color=WHITE, size_pt=13, font="Gotham") +
        run("      ", color=WHITE) +
        run("81 / 100", bold=True, color=WHITE, size_pt=20, font="Gotham") +
        run("    ", color=WHITE) +
        run("🟢  GREEN — Good Performance", bold=True, color=WHITE, size_pt=12, font="Gotham")
    )
    parts.append(para(score_content, shading_color=CORAL, align="center",
                      space_before=80, space_after=80))
    parts.append(spacer(14))

    # ── CALL CONTEXT BOX ──────────────────────────────────────────────────────
    parts.append(heading("Call Context", level=1, color=MID))
    parts.append(para(
        run("Tyson is an active, self-sufficient customer on the Engage AI package. AI messaging is his most-used feature and is producing measurable results. He came with a structured list of questions — primarily around SMS/dialer expansion and pipeline usability issues. No recurring meetings are included in the package; this session was customer-initiated. Chall also had a prepared item (voicemail drops) to introduce.", color=MID, size_pt=10),
        space_before=0, space_after=120
    ))

    # ── SECTION 1: PLATFORM EXPERTISE & FEATURE GUIDANCE ─────────────────────
    parts.append(heading("1.  Platform Expertise & Feature Guidance", level=1))
    parts.append(para(
        run("Score: ", bold=True, color=MID) + run("19 / 25", bold=True, color=CORAL, size_pt=12),
        space_before=0, space_after=60
    ))
    W2 = [3600, 1200, 4560]
    def feat_row(topic, depth, notes, fill=WHITE):
        return row(
            cell(run(topic, color=DARK, size_pt=10), fill=fill, width=W2[0]),
            cell(run(depth, bold=True, size_pt=11, color=DARK), fill=fill, width=W2[1], align="center"),
            cell(run(notes, color=MID, size_pt=9, italic=True), fill=fill, width=W2[2]),
        )
    parts.append(table(
        row(cell(run("TOPIC", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=W2[0]),
            cell(run("DEPTH", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=W2[1], align="center"),
            cell(run("NOTES", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=W2[2])),
        feat_row("Voicemail drop — mechanism explained", "✅ Full", "Two-number trick explained correctly; credit cost, one-time number fee, pre-recording all covered"),
        feat_row("Workflow modification on active campaign", "✅ Full", "Correct: adding steps to bottom of active campaign picks up prospects at end of sequence", fill=LIGHT_BG),
        feat_row("Label strategy for segmentation", "✅ Full", "Proactively introduced 'completed' label to separate nurture vs. new campaigns — added value"),
        feat_row("Intent data search — concept and mechanics", "✅ Full", "Browser tracking / cookie-based intent well explained; daily refresh and 7-day window covered", fill=LIGHT_BG),
        feat_row("Platform database search (269M profiles)", "✅ Full", "Correctly distinguished from Sales Navigator; real-time vs. scrape caveat stated transparently"),
        feat_row("All search source types", "✅ Full", "Sales Nav, CSV, intent, database, post URL, LinkedIn event, recruiter all covered", fill=LIGHT_BG),
        feat_row("Credit system and cost structure", "✅ Full", "Credit-based for SMS/voicemail/enrichment; 500/month threshold explained"),
        feat_row("Dialer — in-app functionality", "✅ Full", "Correctly identified phone icon in platform; explained as upgrade requiring activation"),
        feat_row("SMS — pricing, compliance, forwarding, automation", "⚠️ Partial", "24 min on SMS during the call with no live answers. Post-call message confirmed: virtual number one-time cost, credit usage (from 500/mo), calls forwarding to personal number, voicemail drop pre-recorded mechanism. SMS feature currently unavailable (update in progress) — compliance and automation answers pending feature return", fill=YELLOW),
        feat_row("Recruiter project vs. recruiter search URL", "⚠️ Weak", "Explanation looped and confused Tyson; eventually clarified but unnecessary friction", fill=YELLOW),
        feat_row("AI voice / auto-dialer with LLM", "✅ Honest", "Correctly stated not available; honest about platform current state — no over-promising"),
        feat_row("Make-a-call step (manual reminder, not auto-dial)", "✅ Full", "Clarified clearly that 'make a call' step is a reminder, not automatic dialing", fill=LIGHT_BG),
        total_width=9360, col_widths=W2
    ))
    parts.append(para(
        run("Score rationale: ", bold=True, color=MID, size_pt=9) +
        run("Comprehensive coverage across nearly all features. SMS caused 24 minutes of deferred answers on the call — however, the post-call message confirmed the virtual number pricing, credit model, forwarding behaviour, and voicemail mechanics, and proactively disclosed that SMS is temporarily unavailable. This context softens the call-time gap: some hesitancy was likely navigating the unavailability disclosure. Score revised from 17 to 19. All other topics handled accurately with added value (labels, intent data depth, scrape transparency).", color=MID, size_pt=9),
        space_before=60, space_after=120
    ))

    # ── SECTION 2: ISSUE IDENTIFICATION & RESOLUTION ──────────────────────────
    parts.append(heading("2.  Issue Identification & Resolution", level=1))
    parts.append(para(
        run("Score: ", bold=True, color=MID) + run("15 / 20", bold=True, color=CORAL, size_pt=12),
        space_before=0, space_after=60
    ))
    W3 = [4200, 720, 4440]
    def iss_row(signal, icon, notes, fill=WHITE):
        return row(
            cell(run(signal, color=DARK, size_pt=10), fill=fill, width=W3[0]),
            cell(run(icon, size_pt=11), fill=fill, width=W3[1], align="center"),
            cell(run(notes, color=MID, size_pt=9, italic=True), fill=fill, width=W3[2]),
        )
    parts.append(table(
        row(cell(run("SIGNAL", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=W3[0]),
            cell(run("", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=W3[1]),
            cell(run("ASSESSMENT", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=W3[2])),
        iss_row("Pipeline contact card disappearing on stage move", "✅", "Bug replicated live and escalated to product/tech team post-call. Tyson informed of investigation. Stage reorder fix included in same investigation", fill=GREEN_LT),
        iss_row("Pipeline stage reorder not intuitive", "⚠️", "Noted and added to list; workaround (probability change) already found by customer. Needs documentation or product fix", fill=YELLOW),
        iss_row("SMS features — compliance, forwarding, automation", "⚠️", "Post-call message confirmed virtual number pricing, credits, and forwarding. SMS currently unavailable — compliance and automation answers pending feature return. Tyson will be notified when available", fill=YELLOW),
        iss_row("Can voicemail drop + make-a-call be combined", "⚠️", "Correctly identified two numbers needed; honest about not having tried it; committed to check", fill=YELLOW),
        iss_row("Zero issues resolved in real-time during the call", "❌", "All four open items deferred to post-call. A 64-minute call with nothing resolved in real-time is a pattern risk", fill=RED_LT),
        iss_row("Issues acknowledged and commitments made clearly", "✅", "Chall stated clear commitments on all items: 'I will message you right away' — appropriate expectation-setting", fill=GREEN_LT),
        total_width=9360, col_widths=W3
    ))
    parts.append(para(
        run("Score rationale: ", bold=True, color=MID, size_pt=9) +
        run("Issues identified correctly and meaningful post-call follow-through delivered: pipeline bug escalated to the product team, SMS pricing and forwarding answered, stage reorder folded into the same investigation. The call itself still had zero real-time resolutions across 64 minutes, but the promptness and completeness of the post-call message substantially closes the gap. Score revised from 11 to 15.", color=MID, size_pt=9),
        space_before=60, space_after=120
    ))

    # ── SECTION 3: ACTION ITEMS & ACCOUNTABILITY ──────────────────────────────
    parts.append(heading("3.  Action Items & Accountability", level=1))
    parts.append(para(
        run("Score: ", bold=True, color=MID) + run("18 / 20", bold=True, color=CORAL, size_pt=12),
        space_before=0, space_after=60
    ))
    W4 = [5040, 4320]
    def ai_row(item, sym, fill=WHITE):
        return row(
            cell(run(item, color=DARK, size_pt=10), fill=fill, width=W4[0]),
            cell(run(sym, bold=True, size_pt=10, color=DARK), fill=fill, width=W4[1], align="center"),
        )
    parts.append(table(
        row(cell(run("ITEM", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=W4[0]),
            cell(run("STATUS", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=W4[1], align="center")),
        ai_row("SMS — pricing, credits, forwarding, voicemail mechanism", "✅  Answered in post-call message. SMS unavailable — will notify when back online", fill=GREEN_LT),
        ai_row("SMS — compliance ('say stop') and full automation as campaign step", "⚠️  Pending SMS feature return. Committed to update", fill=YELLOW),
        ai_row("Pipeline contact card bug — investigate and report", "✅  Escalated to product/tech team. Update committed to Tyson", fill=GREEN_LT),
        ai_row("Pipeline stage reorder — investigate fix", "✅  Included in pipeline investigation. Post-call message confirmed", fill=GREEN_LT),
        ai_row("Voicemail drop + make-a-call combined (two numbers) — check with team", "⚠️  Not explicitly confirmed in post-call message — still open", fill=YELLOW),
        ai_row("Post-call message promised to Tyson", "✅  Sent — comprehensive written summary covering all 5 topics", fill=GREEN_LT),
        ai_row("Action items were platform-flagged during call (ACTION ITEM markers)", "✅  Both SMS and pipeline items auto-tagged", fill=LIGHT_BG),
        ai_row("Deadlines given for follow-up items", "❌  'Right away' and 'after this call' — no specific SLA stated", fill=RED_LT),
        ai_row("Written summary / structured action list confirmed", "✅  Post-call message sent — 5 structured topic sections, 3 explicit action items all Chall-owned", fill=GREEN_LT),
        ai_row("Next success check-in or touchpoint planned", "❌  No follow-up session discussed; not in package but no alternative touchpoint set", fill=YELLOW),
        total_width=9360, col_widths=W4
    ))
    parts.append(para(
        run("Score rationale: ", bold=True, color=MID, size_pt=9) +
        run("Post-call message delivered promptly and comprehensively: 5 topic areas covered (AI messaging, phone features, workflow, data sources, pipeline), 3 explicit action items listed, all Chall-owned. Pipeline bug escalated. SMS pricing and forwarding answered; unavailability disclosed clearly. Score revised from 14 to 18. Remaining gaps: no specific SLA on timeline for open items, no follow-up touchpoint scheduled for a package with no recurring meetings.", color=MID, size_pt=9),
        space_before=60, space_after=120
    ))

    # ── SECTION 4: CUSTOMER HEALTH & EXPANSION SIGNALS ────────────────────────
    parts.append(heading("4.  Customer Health & Expansion Signals", level=1))
    parts.append(para(
        run("Score: ", bold=True, color=MID) + run("19 / 20", bold=True, color=CORAL, size_pt=12),
        space_before=0, space_after=60
    ))
    W5 = [4200, 720, 4440]
    def ch_row(signal, icon, notes, fill=WHITE):
        return row(
            cell(run(signal, color=DARK, size_pt=10), fill=fill, width=W5[0]),
            cell(run(icon, size_pt=11), fill=fill, width=W5[1], align="center"),
            cell(run(notes, color=MID, size_pt=9, italic=True), fill=fill, width=W5[2]),
        )
    parts.append(table(
        row(cell(run("SIGNAL", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=W5[0]),
            cell(run("", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=W5[1]),
            cell(run("ASSESSMENT", bold=True, color=WHITE, size_pt=9), fill=CORAL, width=W5[2])),
        ch_row("Customer engagement level", "✅", "Came with a prepared list, drove most of the agenda, asked advanced questions — highly active user", fill=GREEN_LT),
        ch_row("AI feature satisfaction", "✅", "Strong positive feedback: 'I like it', 'I sought you guys out because of this'. Platform champion for AI messaging", fill=GREEN_LT),
        ch_row("Campaign results confirmed", "✅", "Chall confirmed Tyson's campaigns have high acceptance rates and positive responses — acknowledged by Lee", fill=GREEN_LT),
        ch_row("No churn signals", "✅", "Zero intent to leave; customer is building out more use cases (SMS, dialer, recruiting)", fill=GREEN_LT),
        ch_row("Expansion signal: SMS + dialer + virtual number", "✅", "Explicitly asked about cost, setup, and next steps for phone number purchase — warm expansion opportunity", fill=GREEN_LT),
        ch_row("Expansion opportunity actively scoped", "✅", "Post-call message included virtual number one-time cost, credit model (500/mo), and dialer forwarding details — Tyson has the information to activate. Proactive delivery rather than waiting for customer to ask", fill=GREEN_LT),
        ch_row("Customer self-sufficiency is high", "✅", "Tyson is already running campaigns, resuming sequences, and doing his own research — low support burden"),
        ch_row("Platform stickiness (AI learning from responses)", "✅", "Customer noticed the AI learning his communication style — this is a deep retention signal; Chall reinforced it correctly"),
        total_width=9360, col_widths=W5
    ))
    parts.append(para(
        run("Score rationale: ", bold=True, color=MID, size_pt=9) +
        run("High-health, low-churn account with a warm expansion signal that was followed up on post-call. Tyson is a genuine product champion. Post-call message proactively included virtual number pricing and credit details — expansion information delivered without Tyson needing to chase. Score revised from 18 to 19. The only remaining gap is that SMS is currently unavailable, which creates a natural delay in conversion.", color=MID, size_pt=9),
        space_before=60, space_after=120
    ))

    # ── SECTION 5: RAPPORT & ENGAGEMENT QUALITY ───────────────────────────────
    parts.append(heading("5.  Rapport & Engagement Quality", level=1))
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
        rr_row("Natural, relaxed conversation", "✅", "Customer spoke freely and at length; no formal awkwardness — established relationship"),
        rr_row("Chall responsive and patient", "✅", "Followed Tyson's agenda; picked up on each question without redirecting unnecessarily", fill=LIGHT_BG),
        rr_row("Positive performance feedback surfaced and acknowledged", "✅", "Chall told Tyson his campaigns are performing well — proactive positive reinforcement"),
        rr_row("'To be honest, we're not really using the SMS feature'", "⚠️", "Poor opening on the call's primary topic — undermines customer confidence in platform expertise", fill=YELLOW),
        rr_row("Excessive filler ('yeah', 'yep', 'okay') without adding value", "⚠️", "Chall often responded with single filler words while Tyson spoke at length — reduces perceived engagement", fill=YELLOW),
        rr_row("Internet drop at ~35 min", "⚠️", "Call briefly disconnected; recovered smoothly but affected flow", fill=YELLOW),
        rr_row("CSM-initiated feature (voicemail drops) introduced proactively", "✅", "Chall had a prepared item to share — shows preparation and adds value beyond reactive Q&A", fill=GREEN_LT),
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
        issue_row("1", "SMS compliance ('say stop') and full automation as campaign step — pending SMS feature return. Notify Tyson when available", "🟡 Medium", "Chall", "On SMS return", fill=YELLOW),
        issue_row("2", "Pipeline contact card disappears on stage move — escalated to product/tech team. Update due to Tyson", "🟡 Medium", "Chall + Product", "In progress"),
        issue_row("3", "Pipeline stage reorder — folded into pipeline investigation above", "🟡 Medium", "Chall", "In progress", fill=YELLOW),
        issue_row("4", "Voicemail drop + make-a-call combined on same lead — not confirmed in post-call message. Still open", "🟡 Medium", "Chall", "This week"),
        issue_row("5", "SMS/dialer activation — pricing/credits sent to Tyson. Follow up to confirm whether he wants to activate", "🟡 Medium", "Chall", "Follow up within 3 days", fill=YELLOW),
        total_width=9360, col_widths=W7
    ))
    parts.append(spacer(10))

    # ── RISK FLAGS ────────────────────────────────────────────────────────────
    parts.append(heading("Risk Flags", level=1, color="92400E"))
    for title, detail in [
        ("SMS knowledge gap as pattern risk", "SMS came up as the primary topic and Chall had no answers. If a customer on Engage AI — a premium package — repeatedly encounters 'I don't know, I'll check' on core features, confidence in the CSM erodes over time. This should trigger internal SMS training."),
        ("Expansion opportunity at risk of stalling", "Tyson clearly wants the SMS/dialer functionality. 'Just let me know' puts the initiative back on the customer. For a customer with no recurring meetings in their package, there is no scheduled touchpoint to pick this up — if Chall doesn't follow up proactively, this upsell stalls."),
        ("No follow-up touchpoint in package", "Engage AI has no recurring meetings. All continuity relies on the post-call message and customer-initiated contact. If the post-call message is incomplete or delayed, there is no safety net."),
        ("Pipeline bug may cause data loss trust issues", "The contact card disappearing when moving pipeline stages is a UX issue that affects how Tyson manages his leads. If this isn't escalated and communicated back quickly, it erodes trust in the platform as a reliable CRM layer."),
    ]:
        parts.append(bullet(detail, bold_prefix=title, color=DARK))
    parts.append(spacer(10))

    # ── WHAT WENT WELL ────────────────────────────────────────────────────────
    parts.append(heading("What Went Well", level=1, color="065F46"))
    for item in [
        "Tyson's AI satisfaction is exceptional — he is a genuine product champion, comparing the platform favourably to having an SDR. Chall reinforced this with specific campaign performance data (acceptance rates, positive responses).",
        "Voicemail drop was introduced proactively and explained clearly — a good example of CSM-led value delivery, not just reactive Q&A.",
        "Label strategy for segmentation was introduced unprompted — added real strategic value beyond the question asked.",
        "Intent data explanation was one of the strongest moments of the call — clearly explained the cookie/browser tracking mechanism and the 7-day refresh logic.",
        "Pipeline bug was replicated live in real time — demonstrates good diagnostic instinct even without an immediate resolution.",
        "Honest transparency throughout — never over-promised on AI voice, SMS capability, or data freshness. Trust-building through accurate limitations.",
    ]:
        parts.append(bullet(item, color=DARK))
    parts.append(spacer(10))

    # ── RECOMMENDED NEXT STEPS ────────────────────────────────────────────────
    parts.append(heading("Recommended Next Steps", level=1, color=CORAL))
    steps = [
        ("Today", "Send post-call message with full answers to: SMS compliance, natural message format, forwarding to personal phone, and whether SMS can be a fully automated campaign step."),
        ("Today", "Confirm voicemail drop + make-a-call on same lead — two numbers needed or single number sufficient?"),
        ("Today", "Send Tyson the next steps for activating the virtual phone number — do not wait for him to initiate. Frame as: 'Based on today, here is exactly how to get this set up.'"),
        ("This week", "Escalate pipeline contact card bug to product/tech team with replication steps. Send Tyson an update with timeline for fix."),
        ("This week", "Confirm pipeline stage reorder — is there a drag-and-drop fix, or document the probability workaround as official guidance?"),
        ("Internal", "Brief the CS team on SMS features: pricing, compliance rules, forwarding behaviour, and automation capability. This should not be a knowledge gap for any CSM on Engage AI accounts."),
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
        run("Scored against Interceptly CSM success call framework.  |  Health: 🟢 Green (75-100)  |  Revised: post-call message reviewed  |  Package: Engage AI (no recurring meetings)  |  Generated 17 March 2026",
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

---
name: call-scorecard
description: >
  Interceptly CS call scorecard generator. Analyses a meeting transcript and produces
  a scored .docx scorecard in the Interceptly brand format. Use this skill whenever
  the user says "scorecard", "score this call", "analyse my call", "create a scorecard",
  or pastes a call transcript and asks for analysis. Supports all call types: onboarding,
  SDR training, customer success check-in, account transition, campaign strategy session.
  Optionally accepts after-call email context to revise the score post-call.
---

# Call Scorecard Generator

Produces a scored, formatted `.docx` scorecard from a call transcript. Scores are
evidence-based — every point awarded or deducted is tied to a specific moment in the
transcript. The output file follows the Interceptly brand palette and is saved locally.

---

## Step 1 — Gather inputs

Ask the user for any missing context before proceeding. Required inputs:

| Input | How to get it |
|---|---|
| **Transcript** | Pasted directly into the conversation (Fathom copy, Close link, or plain text) |
| **Call date** | From the transcript header or the user |
| **CSM name** | Who ran the call (Chall, Alison, Lee) |
| **Customer / account name** | Company or individual |
| **Call type** | See call types below |
| **Package** | e.g. Engage AI, Engage Pro, Pipeline Builder — affects scoring rubric |

Optional (improves accuracy):
- After-call email or Slack message sent to the customer
- Any additional context the user provides (e.g. "this was a one-off training session")

**Do not start scoring until you have the transcript and call type.**

---

## Call types and scoring rubrics

Choose the rubric that best matches the call. If uncertain, ask the user.

### CS Success / Check-in Call (100 pts)
| Section | Max | What it measures |
|---|---|---|
| Platform Expertise & Feature Guidance | 25 | Accuracy and depth of feature knowledge |
| Issue Identification & Resolution | 20 | Bug handling, unanswered questions, workarounds |
| Action Items & Accountability | 20 | Commitments made, post-call message quality |
| Customer Health & Expansion Signals | 20 | Engagement, churn risk, upsell signals |
| Rapport & Engagement Quality | 15 | Communication style, responsiveness, tone |

### Onboarding Call (100 pts)
| Section | Max | What it measures |
|---|---|---|
| Platform Setup & Configuration | 30 | Technical setup completeness, configuration accuracy |
| Product Knowledge Transfer | 25 | Feature explanation clarity and depth |
| Customer Goals & Expectation Alignment | 20 | ICP discussion, success metrics, package fit |
| Next Steps & Accountability | 15 | Training plan, follow-up scheduled, action items |
| Relationship & Rapport | 10 | First impression, trust building, engagement |

### SDR Training / Platform Overview (100 pts)
| Section | Max | What it measures |
|---|---|---|
| Pre-Call Preparation | 20 | Agenda clarity, attendee knowledge, materials ready |
| Platform Demo & Knowledge Transfer | 25 | Feature coverage, explanation quality, Q&A handling |
| Stakeholder Management | 20 | All attendees engaged, key decision-makers addressed |
| Product Positioning | 15 | Value prop communicated, differentiators highlighted |
| Next Steps & Accountability | 20 | Follow-up confirmed, action items assigned |

### Account Transition / User Handover (100 pts)
| Section | Max | What it measures |
|---|---|---|
| Account Transition Management | 25 | Data migration, access handover, continuity plan |
| Platform Knowledge & Feature Guidance | 20 | Correct answers to technical questions |
| Stakeholder Engagement | 20 | All parties engaged, new user onboarded |
| Action Items & Accountability | 20 | Clear next steps, written follow-up |
| Customer Relationship & Communication | 15 | Trust maintained, tone, rapport with new contact |

### Campaign Strategy Session (100 pts)
| Section | Max | What it measures |
|---|---|---|
| Pre-Call Preparation & Proactivity | 20 | Templates ready, account context known |
| Campaign Knowledge & Consultation | 25 | Sequence logic, best practices, evidence-based guidance |
| Customer Understanding & Needs Alignment | 20 | Customer ask correctly identified, format matched |
| Action Items & Accountability | 20 | Template/asset delivery confirmed, next steps clear |
| Communication & Rapport | 15 | Tone, pace, pushback handling |

---

## Score health bands

| Score | Label | Colour |
|---|---|---|
| 75–100 | 🟢 GREEN — Good Performance | Green |
| 50–74 | 🟡 YELLOW — Needs Attention | Yellow |
| 0–49 | 🔴 RED — Significant Issues | Red |

---

## Step 2 — Analyse the transcript

Read the full transcript carefully. For each section in the rubric:

1. **List the evidence** — specific timestamps or quotes that support or reduce the score
2. **Award a score** — be critical but fair; deduct points for missed opportunities, not just errors
3. **Write a rationale** — 2–3 sentences max, grounded in what actually happened

Scoring principles:
- A perfect score requires no meaningful gaps — not just "nothing went wrong"
- After-call actions (email, Slack message, escalations) can revise section scores upward
- Be consistent: the same behaviour should score the same across CSMs
- Flag patterns, not just one-off moments — a single vague answer is a note; three in a row is a risk flag

---

## Step 3 — Generate the scorecard

Write a Python script that generates the `.docx` file and run it immediately.

### Output file naming convention
```
{CustomerName}_{CallType}_{CallDate}.docx
```
Examples:
- `Tyson_SuccessCall_Scorecard_Mar17.docx`
- `SAH_Nordic_Onboarding_Scorecard.docx`
- `Drummond_Transition_Scorecard_Mar16.docx`

### Output file location
Save to: `~/Documents/Claude/`

### Required document structure (in order)
1. **Header bar** — INTERCEPTLY in coral (#F27A7D), white text
2. **Title** — "Customer Success Call Scorecard", 20pt Gotham
3. **Subtitle** — `{Customer} x {CSM} | Interceptly`, 13pt Inter, MID colour
4. **Meta table** — Date, Duration, Call Type, Package, CSM, Customer, Segment (2×4 grid)
5. **Overall score banner** — coral background, score + health label, centred
6. **Call Context** — 3–4 sentence summary of the call purpose and customer profile
7. **Sections 1–5** — each with: section heading, score display, evidence table, score rationale
8. **Open Issues** — numbered table with: issue, priority (🔴/🟡), owner, deadline
9. **What Went Well** — bulleted list, 4–6 items
10. **Recommended Next Steps** — table with when + action columns
11. **Footer** — framework credit, health band, package, date generated

### Colour palette
```python
CORAL    = "F27A7D"   # primary brand colour — headers, scores, accents
CORAL_LT = "FDEAEA"   # light coral — alternating row fill
DARK     = "1F2937"   # body text
MID      = "4B5563"   # secondary text, labels
LIGHT_BG = "F9FAFB"   # alternating row fill (neutral)
WHITE    = "FFFFFF"
YELLOW   = "FEF3C7"   # warning rows
RED_LT   = "FEE2E2"   # issue rows
GREEN_LT = "D1FAE5"   # positive rows
```

### Evidence table icons
```
✅  = completed / strong / full marks
⚠️  = partial / weak / needs improvement
❌  = missing / failed / significant gap
```

### Script requirements
- Use Python `zipfile` — no external dependencies
- Build valid OOXML: `[Content_Types].xml`, `_rels/.rels`, `word/document.xml`,
  `word/styles.xml`, `word/numbering.xml`, `word/_rels/document.xml.rels`
- Validate with `xml.etree.ElementTree.fromstring()` before writing (optional but recommended)
- Run the script immediately after writing it
- Open the generated file: `open ~/Documents/Claude/{filename}.docx`

---

## Step 4 — Post-call revision (optional)

If the user provides an after-call email or message, re-read it and revise scores where evidence warrants:

- **Action Items section**: did a post-call message get sent? Was it structured?
- **Issue Resolution**: were bugs escalated? Were questions answered?
- **Platform Expertise**: did the follow-up message demonstrate knowledge deferred from the call?
- **Expansion Signals**: did the CSM follow up on upsell opportunities proactively?

Update the overall score banner to include `(Revised — post-call message reviewed)`.
Update the footer health band to match the new score.
Regenerate the `.docx`.

---

## Step 5 — Push to GitHub (optional)

If the user asks to push the scorecard to GitHub, or after completing a scorecard:

1. Check auth: `gh auth status`
2. If not authenticated: `gh auth login --web` and ask user to complete browser flow
3. Copy the generated `.docx` and Python script to the repo:
   ```bash
   cp ~/Documents/Claude/{filename}.docx ~/Documents/Claude/scorecards/docs/
   cp ~/Documents/Claude/{script}.py ~/Documents/Claude/scorecards/scripts/
   ```
4. Commit and push:
   ```bash
   cd ~/Documents/Claude
   git add scorecards/
   git commit -m "Add scorecard: {Customer} {CallType} {Date}"
   git push origin main
   ```

The GitHub repo is: `https://github.com/cpm-interceptly/cs-scorecards`

---

## Example trigger phrases

- "score this call" + transcript pasted
- "analyse my call and create a scorecard"
- "create a scorecard like the others"
- "update the scorecard" + after-call email pasted
- "push the scorecard to the repo"

---

## Notes

- Each CSM runs this skill locally — scorecards are stored in `~/Documents/Claude/` and pushed to the shared GitHub repo
- The GitHub repo (`cs-scorecards`) is the single source of truth for all team scorecards
- JSON data files (structured scoring data) can optionally be generated alongside `.docx` for training analysis aggregation
- The scoring framework is intentionally critical — the goal is training improvement, not validation

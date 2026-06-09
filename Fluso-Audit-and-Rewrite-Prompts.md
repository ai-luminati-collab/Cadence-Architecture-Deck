# Fluso Prompts — Self-Audit + Rewrite
Use these in order. Step 0 is one-time setup. Steps 1 and 2 run on the content.

---

## STEP 0 — One-time: load the Constitution into Fluso
Paste the entire PremAI Writing Constitution into Fluso and say:

> Save this as a permanent project preference called "PremAI Writing Constitution."
> From now on, every piece of content you draft or review for PremAI must obey it.
> Confirm you have saved it and list the 6 banned sentence structures back to me so I know you parsed them.

(If Fluso lists them back correctly, the guardrail is live. If it paraphrases vaguely, re-paste and ask again.)

---

## STEP 1 — The self-audit ("where it fucked up")
Attach the 75-post calendar and paste:

> Audit every post in this calendar against the PremAI Writing Constitution you saved.
> Do NOT rewrite yet. Produce an audit table only, with these columns:
> | Post # | Platform | Offending line (quote it exactly) | Rule broken (cite the Constitution section) | Severity (Critical / Strategic / Craft) | One-line fix |
>
> Rules for the audit:
> - Quote the actual offending text. Do not summarize it.
> - One row per violation, not one row per post. A single post can have several rows.
> - Flag every "not X / it is Y", every rhetorical-question hook, every fragment-list, every use of "compounds" beyond the first, every naked claim with no proof, and every post that leads with abstraction instead of utility.
> - At the end, give me: (a) total violation count, (b) the 5 worst posts, (c) which of the 15 topics are duplicates of each other.
>
> Be harsh. Assume a senior reviewer who reads AI content all day is grading this.

This is the "explain where it all went wrong" step. Save its output — it becomes your recurring flop-log (Suhan asked for exactly this).

---

## STEP 2 — The rewrite
After the audit, paste:

> Now rewrite all 75 posts to fix every violation in your audit, obeying the Constitution.
> Hard requirements:
> - 50/50 utility-to-trust. Half the posts must lead with what Fluso does for the reader's day, not with architecture.
> - Fluso gets ~60% of product emphasis vs Confidential API.
> - Tag each post with its single target buyer.
> - Vary the opening line type across posts: scene, stat, customer quote, flat fact, short story, definition. Never two posts in a row that open the same way.
> - Every technical claim links to a proof (Reticle repo, a benchmark, a demo, an attestation viewer).
> - One specific CTA per post, varied across the calendar.
> - Consolidate the token/open-source/sovereign/rent-vs-own topics to 3 expressions total. Use the freed slots for: a founder POV from Simo, a "we tried to break our own attestation" teardown, and an influencer-enablement brief.
> - X under 280 or threaded. Meta carousels max 2 lines per slide, no rhetorical-question covers, max 3 real hashtags.
>
> Output as a clean doc, grouped by platform, with each post's target-buyer tag and CTA labelled.
> After the doc, run the Part 7 checklist against your own rewrite and tell me any box you cannot tick.

---

## STEP 3 — Compare
You will have two rewrites: Fluso's and Claude's benchmark doc.
Put them side by side on the 15 LinkedIn posts first (the priority channel).
Keep whichever wins line by line. Feed the winners back to Fluso as new examples so its next batch is better.

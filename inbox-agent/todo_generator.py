"""Uses Claude API to analyse emails and produce a prioritised to-do list."""

import logging
from datetime import datetime
from typing import List, Dict

import anthropic

logger = logging.getLogger(__name__)

# Number of emails to send to Claude — enough for context without blowing the window
MAX_EMAILS_FOR_CLAUDE = 60


def _summarise_emails(emails: List[Dict]) -> str:
    """Render a compact text summary of emails suitable for Claude."""
    lines = []
    for i, e in enumerate(emails[:MAX_EMAILS_FOR_CLAUDE], 1):
        unread_flag = "★ UNREAD" if e.get("is_unread") else "  read"
        importance = ""
        if e.get("importance", "").lower() == "high":
            importance = " [HIGH IMPORTANCE]"
        awaiting_flag = " [AWAITING REPLY]" if e.get("awaiting_reply") else ""
        lines.append(
            f"{i:>3}. [{e.get('source', '?').upper():>7}] {unread_flag}{importance}{awaiting_flag}\n"
            f"       From: {e.get('from', '')}\n"
            f"    Subject: {e.get('subject', '(no subject)')}\n"
            f"       Date: {e.get('date', '')[:25]}\n"
            f"    Preview: {e.get('snippet', '')[:180]}"
        )
    return "\n\n".join(lines)


def generate_todo_list(
    gmail_emails: List[Dict],
    outlook_emails: List[Dict],
    anthropic_api_key: str,
) -> str:
    """
    Send combined inbox contents to Claude and get back a prioritised to-do list.
    Returns formatted Markdown that can be embedded in the briefing email.
    """
    client = anthropic.Anthropic(api_key=anthropic_api_key)

    today = datetime.now().strftime("%A, %d %B %Y")

    # Interleave by recency — unread first, then read
    all_emails = (
        [e for e in gmail_emails + outlook_emails if e.get("is_unread")]
        + [e for e in gmail_emails + outlook_emails if not e.get("is_unread")]
    )

    email_block = _summarise_emails(all_emails)

    prompt = f"""You are the executive assistant for Eddie G, a commercial real estate agent at IB Property in Sydney, Australia.
His email addresses are edwardenag@gmail.com (personal/Gmail) and edward@ibproperty.com.au (work/Outlook).
Today is {today}.

Below is a summary of his combined inbox. Starred items (★ UNREAD) have not been read yet.
Items marked [AWAITING REPLY] have been confirmed (by checking the actual email
thread/conversation) to have NO reply sent from Eddie yet — this is different
from "unread": a read-but-un-actioned email is exactly what this flag is for.

---
{email_block}
---

Your job: produce a clear, actionable, prioritised to-do list based on these emails.

Rules:
1. Only include items that require Eddie's ACTION — replies, calls, documents, follow-ups, approvals.
2. IGNORE promotional emails, newsletters, and anything auto-generated with no required response.
3. Prioritise: urgent client/deal emails first, then landlord/tenant issues, then admin, then low-priority.
4. Be specific and concrete — "Reply to John Smith re: 45 George St lease renewal offer" not "reply to email".
5. Flag deadlines or time-sensitive items prominently.
6. If you spot something in [OUTLOOK] that duplicates [GMAIL], mention it once, note it appears in both.
7. Every item marked [AWAITING REPLY] goes ONLY in the "✉️ Awaiting Your Reply" section below —
   do not also duplicate it under Urgent/Important/etc. These are specifically the
   things Eddie needs to get onto today because he genuinely hasn't responded yet.

Format your response exactly like this (omit any section with no items):

## Prioritised To-Do — {today}

### ✉️ Awaiting Your Reply
1. **[action verb + what]** — *[from name/company]* | [Gmail/Outlook]
   → [one-line context note if helpful]

### 🔴 Urgent — Do Today
1. **[action verb + what]** — *[from name/company]* | [Gmail/Outlook]
   → [one-line context note if helpful]

### 🟡 Important — This Week
1. **[action]** — *[from]* | [source]

### 🟢 When You Have a Moment
1. **[action]** — *[from]* | [source]

### 📋 FYI — No Action Required
- [brief note on anything notable that doesn't need a response]

Keep each item to one or two lines maximum. Do not pad the list."""

    try:
        response = client.messages.create(
            model="claude-opus-4-8",
            max_tokens=2048,
            messages=[{"role": "user", "content": prompt}],
        )
        todo_md = response.content[0].text
        logger.info("Claude to-do list generated (%d chars)", len(todo_md))
        return todo_md

    except anthropic.APIError as exc:
        logger.error("Claude API error: %s", exc)
        return (
            f"## Prioritised To-Do — {today}\n\n"
            f"*(Error generating AI to-do list: {exc})*\n\n"
            f"Please review your inboxes manually."
        )

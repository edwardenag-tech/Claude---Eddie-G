"""Format and send the morning briefing email to both addresses."""

import logging
from datetime import datetime
from typing import Dict, List

import markdown
import pytz

logger = logging.getLogger(__name__)

SYDNEY_TZ = pytz.timezone("Australia/Sydney")


# ─── HTML helpers ────────────────────────────────────────────────────────────

def _email_table(emails: List[Dict], max_rows: int = 25) -> str:
    """Render a list of email dicts as an HTML table."""
    if not emails:
        return '<p style="color:#6b7280;font-style:italic;">No new emails.</p>'

    rows = []
    for e in emails[:max_rows]:
        bold = "font-weight:600;" if e.get("is_unread") else ""
        dot = '<span style="color:#ef4444;font-size:9px;vertical-align:middle;">●</span> ' if e.get("is_unread") else ""
        rows.append(
            f'<tr>'
            f'<td style="{bold}padding:5px 10px 5px 0;max-width:220px;overflow:hidden;'
            f'white-space:nowrap;text-overflow:ellipsis;">'
            f'{dot}{e.get("from", "")[:45]}</td>'
            f'<td style="{bold}padding:5px 10px;max-width:320px;overflow:hidden;'
            f'white-space:nowrap;text-overflow:ellipsis;">'
            f'{e.get("subject", "")[:70]}</td>'
            f'<td style="padding:5px 0;color:#6b7280;white-space:nowrap;font-size:12px;">'
            f'{str(e.get("date",""))[:16]}</td>'
            f'</tr>'
        )

    overflow = ""
    if len(emails) > max_rows:
        overflow = (
            f'<tr><td colspan="3" style="padding:6px 0;color:#6b7280;font-style:italic;font-size:12px;">'
            f'… and {len(emails) - max_rows} more</td></tr>'
        )

    return (
        '<table style="width:100%;border-collapse:collapse;font-size:13px;">'
        '<thead><tr style="border-bottom:2px solid #e5e7eb;">'
        '<th style="text-align:left;padding:6px 10px 6px 0;color:#374151;">From</th>'
        '<th style="text-align:left;padding:6px 10px;color:#374151;">Subject</th>'
        '<th style="text-align:left;padding:6px 0;color:#374151;">Received</th>'
        '</tr></thead>'
        f'<tbody>{"".join(rows)}{overflow}</tbody>'
        "</table>"
    )


def _cleaning_summary(report: Dict) -> str:
    """Render the cleaning report as an HTML summary table + expandable action log."""
    total = sum(
        report.get(k, 0)
        for k in (
            "gmail_archived", "gmail_promoted", "gmail_trashed", "gmail_aggressive_deleted",
            "outlook_archived", "outlook_promoted", "outlook_deleted", "outlook_aggressive_deleted",
        )
    )

    if total == 0:
        return '<p style="color:#6b7280;font-style:italic;">Inboxes already tidy — no actions taken overnight.</p>'

    def row(label, emoji, gmail_key, outlook_key):
        g = report.get(gmail_key, 0)
        o = report.get(outlook_key, 0)
        if g + o == 0:
            return ""
        return (
            f'<tr><td style="padding:4px 20px 4px 0;">{emoji} {label}</td>'
            f'<td style="text-align:right;padding:4px 12px;">{g}</td>'
            f'<td style="text-align:right;padding:4px 0;">{o}</td></tr>'
        )

    gmail_total = (
        report.get("gmail_archived", 0) + report.get("gmail_promoted", 0)
        + report.get("gmail_trashed", 0) + report.get("gmail_aggressive_deleted", 0)
    )
    outlook_total = (
        report.get("outlook_archived", 0) + report.get("outlook_promoted", 0)
        + report.get("outlook_deleted", 0) + report.get("outlook_aggressive_deleted", 0)
    )

    rows = "".join([
        row("Archived", "📥", "gmail_archived", "outlook_archived"),
        row("→ Promotions", "🏷️", "gmail_promoted", "outlook_promoted"),
        row("Trashed / Deleted", "🗑️", "gmail_trashed", "outlook_deleted"),
        row("Deleted (AI-judged: competitor spam / irrelevant CC)", "🤖",
            "gmail_aggressive_deleted", "outlook_aggressive_deleted"),
        f'<tr style="border-top:2px solid #e5e7eb;font-weight:600;">'
        f'<td style="padding:6px 20px 4px 0;">Total</td>'
        f'<td style="text-align:right;padding:6px 12px 4px;">{gmail_total}</td>'
        f'<td style="text-align:right;padding:6px 0 4px;">{outlook_total}</td></tr>',
    ])

    table = (
        '<table style="font-size:13px;border-collapse:collapse;">'
        '<thead><tr style="border-bottom:2px solid #e5e7eb;">'
        '<th style="text-align:left;padding:6px 20px 6px 0;">Action</th>'
        '<th style="text-align:right;padding:6px 12px;">Gmail</th>'
        '<th style="text-align:right;padding:6px 0;">Outlook</th>'
        '</tr></thead>'
        f'<tbody>{rows}</tbody></table>'
    )

    actions = report.get("actions", [])
    log_section = ""
    if actions:
        items = "".join(
            f'<li style="font-size:12px;color:#6b7280;padding:2px 0;">{a}</li>'
            for a in actions
        )
        log_section = (
            f'<details style="margin-top:12px;">'
            f'<summary style="cursor:pointer;color:#6b7280;font-size:13px;">'
            f'View full action log ({len(actions)} items)</summary>'
            f'<ul style="margin:8px 0 0 0;padding-left:20px;">{items}</ul>'
            f"</details>"
        )

    return table + log_section


# ─── Email body builder ───────────────────────────────────────────────────────

def build_briefing_html(
    todo_md: str,
    gmail_new: List[Dict],
    outlook_new: List[Dict],
    cleaning_report: Dict,
    date_str: str,
) -> str:
    """Assemble the full HTML briefing email."""

    todo_html = markdown.markdown(todo_md, extensions=["nl2br"])

    g_count = len(gmail_new)
    o_count = len(outlook_new)
    g_unread = sum(1 for e in gmail_new if e.get("is_unread"))
    o_unread = sum(1 for e in outlook_new if e.get("is_unread"))

    def badge(text):
        return (
            f'<span style="display:inline-block;background:#1d4ed8;color:#fff;'
            f'border-radius:999px;padding:2px 12px;font-size:12px;margin-left:8px;'
            f'vertical-align:middle;">{text}</span>'
        )

    section_style = (
        "background:#fff;margin:0;padding:28px 36px;"
        "border-bottom:1px solid #e5e7eb;"
    )
    h2_style = (
        "margin:0 0 18px 0;font-size:13px;font-weight:700;text-transform:uppercase;"
        "letter-spacing:.08em;color:#1e3a5f;"
    )

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Morning Briefing — {date_str}</title>
<style>
  body {{
    margin: 0; padding: 0;
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Arial, sans-serif;
    background: #f3f4f6; color: #111827;
  }}
  .wrapper {{ max-width: 760px; margin: 24px auto; border-radius: 10px;
              overflow: hidden; box-shadow: 0 4px 24px rgba(0,0,0,.10); }}
  .todo-body h2 {{ font-size: 16px; color: #1e3a5f; margin: 20px 0 6px; }}
  .todo-body h3 {{ font-size: 14px; color: #374151; margin: 16px 0 6px; }}
  .todo-body ol, .todo-body ul {{ padding-left: 20px; line-height: 1.9; }}
  .todo-body li {{ margin-bottom: 4px; }}
  .todo-body strong {{ color: #111827; }}
  .todo-body em {{ color: #4b5563; }}
  a {{ color: #2563eb; }}
</style>
</head>
<body>
<div class="wrapper">

  <!-- Header -->
  <div style="background:linear-gradient(135deg,#1e3a5f 0%,#1d4ed8 100%);
              color:#fff;padding:32px 36px;">
    <div style="font-size:26px;margin-bottom:6px;">🌅 Morning Briefing</div>
    <div style="opacity:.85;font-size:14px;">{date_str} &nbsp;·&nbsp; Eddie G — IB Property Sydney</div>
  </div>

  <!-- Section 1: To-Do -->
  <div style="{section_style}">
    <h2 style="{h2_style}">1 · Prioritised To-Do List</h2>
    <div class="todo-body">{todo_html}</div>
  </div>

  <!-- Section 2: Gmail -->
  <div style="{section_style}">
    <h2 style="{h2_style}">
      2 · Gmail Inbox
      {badge(f"{g_count} emails · {g_unread} unread")}
    </h2>
    {_email_table(gmail_new)}
  </div>

  <!-- Section 3: Outlook -->
  <div style="{section_style}">
    <h2 style="{h2_style}">
      3 · Outlook Inbox
      {badge(f"{o_count} emails · {o_unread} unread")}
    </h2>
    {_email_table(outlook_new)}
  </div>

  <!-- Section 4: Cleaning Report -->
  <div style="{section_style}">
    <h2 style="{h2_style}">4 · Overnight Cleaning Report</h2>
    {_cleaning_summary(cleaning_report)}
  </div>

  <!-- Footer -->
  <div style="background:#f9fafb;padding:16px 36px;text-align:center;
              font-size:12px;color:#9ca3af;border-top:1px solid #e5e7eb;">
    Inbox Agent · {date_str} ·
    <a href="mailto:edward@ibproperty.com.au" style="color:#9ca3af;">edward@ibproperty.com.au</a>
  </div>

</div>
</body>
</html>"""


# ─── Send function ────────────────────────────────────────────────────────────

def send_briefing(
    gmail_client,
    outlook_client,
    todo_list: str,
    gmail_new: List[Dict],
    outlook_new: List[Dict],
    cleaning_report: Dict,
    notify_emails: List[str],
) -> bool:
    """
    Build the briefing HTML and send it.
    Tries Gmail first; falls back to Outlook if Gmail send fails.
    """
    now_sydney = datetime.now(SYDNEY_TZ)
    date_str = now_sydney.strftime("%A, %d %B %Y")
    subject = f"🌅 Morning Briefing — {date_str}"

    html_body = build_briefing_html(
        todo_md=todo_list,
        gmail_new=gmail_new,
        outlook_new=outlook_new,
        cleaning_report=cleaning_report,
        date_str=date_str,
    )

    sent = False

    if gmail_client:
        sent = gmail_client.send_email(notify_emails, subject, html_body)
        if sent:
            logger.info("Briefing dispatched via Gmail to %s", notify_emails)
        else:
            logger.warning("Gmail send failed — will try Outlook")

    if not sent and outlook_client:
        sent = outlook_client.send_email(notify_emails, subject, html_body)
        if sent:
            logger.info("Briefing dispatched via Outlook to %s", notify_emails)
        else:
            logger.error("Outlook send also failed — briefing not delivered")

    return sent

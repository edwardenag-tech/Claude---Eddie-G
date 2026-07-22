"""
Inbox Agent — main orchestrator.

Usage:
    python agent.py           # run the full agent (clean + briefing)
    python agent.py --auth    # authenticate only (useful for first-time setup)
    python agent.py --dry-run # fetch + analyse but don't clean or send

Schedule with cron (7:30am Sydney):
    30 7 * * * cd /path/to/inbox-agent && /usr/bin/python3 agent.py >> logs/cron.log 2>&1

Or with launchd — see README.md for the plist template.
"""

import argparse
import logging
import os
import sys
from datetime import datetime
from pathlib import Path

import pytz
from dotenv import load_dotenv

# Load .env from the same directory as this script
_HERE = Path(__file__).parent
load_dotenv(_HERE / ".env")

# ─── Logging ─────────────────────────────────────────────────────────────────

_LOG_DIR = _HERE / "logs"
_LOG_DIR.mkdir(exist_ok=True)
_today = datetime.now().strftime("%Y-%m-%d")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)-8s] %(name)s: %(message)s",
    handlers=[
        logging.FileHandler(_LOG_DIR / f"inbox_agent_{_today}.log"),
        logging.StreamHandler(sys.stdout),
    ],
)
logger = logging.getLogger("inbox_agent")


# ─── Config ──────────────────────────────────────────────────────────────────

def load_config() -> dict:
    """Read config from environment variables (populated from .env)."""
    cfg = {
        "anthropic_api_key":     os.environ.get("ANTHROPIC_API_KEY", ""),
        "azure_client_id":       os.environ.get("AZURE_CLIENT_ID", "077f47ad-e210-48dd-997f-5ca413ca735c"),
        "azure_tenant_id":       os.environ.get("AZURE_TENANT_ID", "255db1c7-fd44-4604-9aea-032bd2883c0d"),
        "user_gmail":            os.environ.get("USER_GMAIL", "edwardenag@gmail.com"),
        "user_outlook":          os.environ.get("USER_OUTLOOK", "edward@ibproperty.com.au"),
        "notify_emails":         [
            e.strip()
            for e in os.environ.get(
                "NOTIFY_EMAILS", "edwardenag@gmail.com,edward@ibproperty.com.au"
            ).split(",")
            if e.strip()
        ],
        "gmail_credentials_path": str(_HERE / os.environ.get("GMAIL_CREDENTIALS_PATH", "gmail_credentials.json")),
        "gmail_token_path":       str(_HERE / os.environ.get("GMAIL_TOKEN_PATH", "gmail_token.json")),
        "msal_token_cache_path":  str(_HERE / os.environ.get("MSAL_TOKEN_CACHE_PATH", "msal_token_cache.json")),
        "campaign_doc_id":        os.environ.get("CAMPAIGN_DOC_ID", "12gKTGiqwqgEc5iBnypfKFP79bEKThc1ltBOnIvC8Mrk"),
    }

    if not cfg["anthropic_api_key"]:
        logger.warning("ANTHROPIC_API_KEY not set — to-do list generation will be skipped")

    return cfg


# ─── Auth helpers ─────────────────────────────────────────────────────────────

def build_gmail_client(cfg: dict):
    """Authenticate Gmail and return a GmailClient, or None on failure."""
    creds_path = cfg["gmail_credentials_path"]
    if not os.path.exists(creds_path):
        logger.warning(
            "Gmail credentials not found at '%s'. Skipping Gmail. "
            "See README.md to set up OAuth2.",
            creds_path,
        )
        return None

    try:
        from gmail_client import GmailClient
        client = GmailClient(
            credentials_path=creds_path,
            token_path=cfg["gmail_token_path"],
        )
        logger.info("Gmail client ready")
        return client
    except Exception as exc:
        logger.error("Gmail auth failed: %s", exc)
        return None


def build_outlook_client(cfg: dict):
    """Authenticate Outlook via MSAL and return an OutlookClient, or None on failure."""
    try:
        from outlook_client import OutlookClient
        client = OutlookClient(
            client_id=cfg["azure_client_id"],
            tenant_id=cfg["azure_tenant_id"],
            token_cache_path=cfg["msal_token_cache_path"],
        )
        logger.info("Outlook client ready")
        return client
    except Exception as exc:
        logger.error("Outlook auth failed: %s", exc)
        return None


# ─── Main ────────────────────────────────────────────────────────────────────

def run(dry_run: bool = False):
    cfg = load_config()

    logger.info("=" * 65)
    logger.info("Inbox Agent starting (dry_run=%s)", dry_run)
    logger.info("=" * 65)

    # ── Authenticate ──────────────────────────────────────────────────────────
    gmail = build_gmail_client(cfg)
    outlook = build_outlook_client(cfg)

    if not gmail and not outlook:
        logger.error("No email clients available — cannot continue.")
        sys.exit(1)

    # ── Fetch recent emails (for briefing) ────────────────────────────────────
    logger.info("Fetching recent emails (last 24 h)…")

    from gmail_client import GmailClient
    from outlook_client import OutlookClient

    gmail_raw = gmail.get_recent_emails(since_days=1) if gmail else []
    outlook_raw = outlook.get_recent_emails(since_days=1) if outlook else []

    logger.info("Gmail: %d emails | Outlook: %d emails", len(gmail_raw), len(outlook_raw))

    gmail_emails = [GmailClient.extract_email_data(m) for m in gmail_raw]
    outlook_emails = [OutlookClient.extract_email_data(m) for m in outlook_raw]

    # ── Clean inboxes ─────────────────────────────────────────────────────────
    cleaning_report: dict = {
        "gmail_archived": 0, "gmail_promoted": 0, "gmail_trashed": 0, "gmail_aggressive_deleted": 0,
        "outlook_archived": 0, "outlook_promoted": 0, "outlook_deleted": 0, "outlook_aggressive_deleted": 0,
        "actions": [],
    }

    if dry_run:
        logger.info("[DRY RUN] Skipping inbox cleaning")
    else:
        logger.info("Cleaning inboxes…")
        from cleaner import InboxCleaner
        cleaner = InboxCleaner(
            gmail_client=gmail,
            outlook_client=outlook,
            anthropic_api_key=cfg["anthropic_api_key"],
            user_addresses=[cfg["user_gmail"], cfg["user_outlook"]],
        )
        cleaning_report = cleaner.run_all()
        logger.info(
            "Cleaning done — Gmail: %d archived, %d promoted, %d trashed, %d AI-deleted | "
            "Outlook: %d archived, %d promoted, %d deleted, %d AI-deleted",
            cleaning_report["gmail_archived"],
            cleaning_report["gmail_promoted"],
            cleaning_report["gmail_trashed"],
            cleaning_report["gmail_aggressive_deleted"],
            cleaning_report["outlook_archived"],
            cleaning_report["outlook_promoted"],
            cleaning_report["outlook_deleted"],
            cleaning_report["outlook_aggressive_deleted"],
        )

    # ── Draft enquiry replies ─────────────────────────────────────────────────
    if dry_run:
        logger.info("[DRY RUN] Skipping enquiry-reply drafting")
    else:
        logger.info("Drafting enquiry replies…")
        try:
            from draft_agent import main as draft_agent_main
            draft_agent_main()
            logger.info("Enquiry-reply drafting done")
        except Exception as exc:
            logger.error("Enquiry-reply drafting failed (continuing): %s", exc)

    # ── Generate to-do list ───────────────────────────────────────────────────
    todo_list = ""
    if cfg["anthropic_api_key"]:
        logger.info("Generating to-do list with Claude…")
        from todo_generator import generate_todo_list
        todo_list = generate_todo_list(
            gmail_emails=gmail_emails,
            outlook_emails=outlook_emails,
            anthropic_api_key=cfg["anthropic_api_key"],
        )
    else:
        today_str = datetime.now().strftime("%A, %d %B %Y")
        todo_list = (
            f"## Prioritised To-Do — {today_str}\n\n"
            "*(ANTHROPIC_API_KEY not configured — to-do list unavailable)*\n\n"
            "Please check your inboxes manually."
        )

    # ── Fetch today's calendar ────────────────────────────────────────────────
    calendar_events = None
    try:
        if outlook:
            logger.info("Fetching today's calendar events…")
            calendar_events = outlook.get_todays_events()
            logger.info("Calendar: %d event(s) today", len(calendar_events))
        else:
            logger.info("Outlook not connected — skipping calendar")
    except Exception as exc:
        logger.error("Calendar fetch failed (continuing): %s", exc)
        calendar_events = None

    # ── Fetch campaign highlights ─────────────────────────────────────────────
    campaign_highlights = []
    try:
        logger.info("Fetching campaign highlights from live doc…")
        from docs_client import DocsClient
        from campaign_doc_parser import parse_active_campaigns, summarize_campaign_highlights
        docs_client = DocsClient(
            credentials_path=cfg["gmail_credentials_path"],
            token_path=cfg["gmail_token_path"],
        )
        doc_text = docs_client.fetch_doc_text(cfg["campaign_doc_id"])
        if doc_text:
            campaigns = parse_active_campaigns(doc_text)
            campaign_highlights = summarize_campaign_highlights(campaigns)
            logger.info("Campaign highlights: %d active campaign(s)", len(campaign_highlights))
        else:
            logger.warning("Campaign doc returned no content")
    except Exception as exc:
        logger.error("Campaign highlights fetch failed (continuing): %s", exc)
        campaign_highlights = []

    # ── Send morning briefing ─────────────────────────────────────────────────
    if dry_run:
        logger.info("[DRY RUN] Skipping briefing send. To-do list preview:")
        print("\n" + todo_list + "\n")
    else:
        logger.info("Sending morning briefing to %s…", cfg["notify_emails"])
        from briefing import send_briefing
        success = send_briefing(
            gmail_client=gmail,
            outlook_client=outlook,
            todo_list=todo_list,
            gmail_new=gmail_emails,
            outlook_new=outlook_emails,
            cleaning_report=cleaning_report,
            notify_emails=cfg["notify_emails"],
            calendar_events=calendar_events,
            campaign_highlights=campaign_highlights,
        )
        if success:
            logger.info("Morning briefing sent successfully!")
        else:
            logger.error("Failed to deliver morning briefing")
            sys.exit(1)

    logger.info("Inbox Agent finished")


# ─── Entry point ──────────────────────────────────────────────────────────────

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Inbox Agent for Eddie G")
    parser.add_argument(
        "--auth",
        action="store_true",
        help="Authenticate with Gmail and Outlook only (no email processing)",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Fetch and analyse emails but do not clean or send the briefing",
    )
    args = parser.parse_args()

    if args.auth:
        cfg = load_config()
        logger.info("Auth-only mode — testing connections…")
        gmail = build_gmail_client(cfg)
        outlook = build_outlook_client(cfg)
        if gmail:
            logger.info("Gmail: OK")
        if outlook:
            logger.info("Outlook: OK")
        logger.info("Auth complete.")
    else:
        run(dry_run=args.dry_run)

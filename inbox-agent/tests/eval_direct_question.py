"""Live check of draft_agent.claude_classify's "does this ask Eddie something
directly?" judgement, against synthetic labelled emails. Needs a VALID
ANTHROPIC_API_KEY (calls the real model; costs a few cents). Not part of the
unit-test run -- invoke directly:

    python3 tests/eval_direct_question.py

Exit code 0 only if every case matches its expected label. The recipient
filter (is_cc_only) is deterministic and unit-tested, so it's not repeated here.
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import anthropic
from dotenv import load_dotenv

load_dotenv(os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), ".env"))
import draft_agent  # noqa: E402

ME = "edward@ibproperty.com.au"


def case(name, expect, subject, body, to=ME, cc="", sender="Alex Sample", addr="alex@example.com"):
    return dict(name=name, expect=expect, email={
        "from_name": sender, "from": addr, "to": to, "cc": cc, "subject": subject, "body": body,
    })


CASES = [
    # ── should draft ────────────────────────────────────────────────────────
    case("price question on a listing", True, "12 Smith St, Sydney",
         "Hi Edward, what's the asking rent for the ground floor and is parking included?"),
    case("request for the IM", True, "Re: 8 Pitt St sale",
         "Hi Edward, can you send through the Information Memorandum? Thanks, Alex"),
    case("scheduling question from a landlord", True, "Inspection Tuesday",
         "Edward, are you free Tuesday 2pm to walk the tenant through?", sender="Lena Landlord"),
    case("direct question among several To, addressed to him", True, "Fit-out access",
         "Edward - can you confirm whether the contractor can access level 2 on Friday? Sarah FYI.",
         to=f"{ME}, sarah@ibproperty.com.au"),
    case("question below a long greeting and signature", True, "Vendor update",
         "Hi Edward,\n\nHope you had a good weekend. Thanks again for the update last week on the campaign "
         "and for the photos.\n\nOne thing before Thursday - has the offer from the second buyer come in "
         "yet?\n\nRegards\nVic Vendor\nDirector | Vendor Pty Ltd\nThis email is confidential."),
    # ── should NOT draft ────────────────────────────────────────────────────
    case("FYI update", False, "Campaign update", "Hi all, just letting you know the open home went well, 12 groups through."),
    case("thank you", False, "Re: 8 Pitt St", "Thanks Edward, much appreciated. Speak soon."),
    case("interest but no question or request", False, "12 Smith St enquiry",
         "Hi, I'm interested in the property at 12 Smith Street."),
    case("question addressed to someone else", False, "Fit-out",
         "@Sarah can you confirm the contractor's start date? Edward, just keeping you across it.",
         to=f"sarah@ibproperty.com.au, {ME}"),
    case("question to the group in general", False, "Anyone know?",
         "Does anyone know if the Kent St building has been sold? Cheers",
         to=f"team@ibproperty.com.au, {ME}"),
    case("question only in quoted history", False, "Re: Suite 4",
         "Great, thanks for that.\n\nOn Mon, Edward Ghattas wrote:\n> Can you confirm the lease start date?\n> Regards Edward"),
    case("cold sales pitch with rhetorical question", False, "Quick question",
         "Hi Edward, are you tired of high loan rates? Our brokers can save you thousands. Book a call today!",
         sender="Marcus Broker", addr="marcus@loans.example.com"),
    case("portal stats report", False, "Your listing performance", "Your listing received 214 views this week. Are you getting enough leads? Upgrade now.",
         addr="noreply@portal.example.com"),
    case("LEASED notice", False, "LEASED - 4 King St", "Congratulations, 4 King St has been leased. No action required."),
]


def main() -> int:
    ai = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY", ""))
    wrong = 0
    for c in CASES:
        try:
            category, asks, question = draft_agent.claude_classify(ai, c["email"])
        except anthropic.AuthenticationError:
            print("ANTHROPIC_API_KEY is invalid -- cannot run the evaluation.")
            return 2
        ok = asks == c["expect"]
        wrong += not ok
        print(f"{'PASS' if ok else 'FAIL'}  expect={str(c['expect']):5} got={str(asks):5} [{category}]  {c['name']}"
              + (f"\n        question: {question}" if question else ""))
    print(f"\n{len(CASES) - wrong}/{len(CASES)} correct")
    return 1 if wrong else 0


if __name__ == "__main__":
    sys.exit(main())

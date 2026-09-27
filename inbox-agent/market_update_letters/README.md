# Market update letters

One PDF per **suburb** — not per region. A region like Lower North Shore
covers ~34 suburbs; each week Eddie writes (or has designed) a letter for
whichever of those suburbs he actually has an update for, and only those
suburbs get emailed. `market_update_agent.py` reads recipients from the
workbook tab for each region (see `SEGMENTS` in `market_update_agent.py`),
groups them by suburb, and for each suburb checks this folder for a matching
PDF. No PDF here this week for a suburb → that suburb is skipped entirely,
nothing drafted for it. It never invents a letter, never sends the last
suburb's letter to a different suburb, and never writes to this folder.

## File format

```
market_update_letters/<suburb>.pdf     (required -- this IS the letter)
market_update_letters/<suburb>.html    (optional -- overrides subject/greeting)
```

- **`<suburb>.pdf`** — Eddie's actual designed letter for that suburb: the
  polished, branded PDF with photos and FOR SALE / FOR LEASE / LEASED
  sections, exactly as sent to a normal recipient. Every page of this PDF is
  embedded as an image directly in the email body (not attached as a
  separate file) — recipients see the real designed pages when they open
  the email, no click-through needed.
- **`<suburb>` (the filename)** — the suburb name, lowercased, with spaces
  and punctuation collapsed to underscores. `Northbridge` → `northbridge.pdf`,
  `St Ives` → `st_ives.pdf`, `Kirribilli` → `kirribilli.pdf`. Matching is
  case/punctuation-insensitive against the workbook's Suburb column, so the
  exact capitalization doesn't matter.
- **`<suburb>.html`** — optional. If present, overrides the default subject
  line and the short greeting shown above the embedded PDF pages. Same
  format as before:

  ```
  Subject: Market update for {{suburb}}

  <p>Hi {{first_name}},</p>
  <p>... short intro, in HTML ...</p>
  ```

  Line 1 must be `Subject: ...`; everything after the blank line is the
  greeting HTML, with merge fields substituted. If this file is missing,
  the default subject is `Market update for {{suburb}}` and the default
  greeting is `Hi {{first_name}}, here's this week's market update for
  {{suburb}}.`

## Merge fields

Available in the subject line and greeting (not inside the PDF itself,
which is a fixed image):

| Field                | Example                          | Notes |
|-----------------------|-----------------------------------|-------|
| `{{first_name}}`      | `Lyndon`                         | Falls back to `there` if blank |
| `{{last_name}}`        | `Catzel`                         | Blank if not on file |
| `{{full_name}}`        | `Lyndon Catzel`                  | Falls back to the company name, then `there` |
| `{{suburb}}`           | `Beecroft`                       | Title-cased from the workbook's ALL CAPS |
| `{{street_address}}`   | `14A Hannah Street`              | Title-cased |
| `{{state}}`            | `NSW`                            | |
| `{{postcode}}`         | `2119`                           | |
| `{{property_type}}`    | `Commercial`                     | |

## Dedupe

An owner gets one draft per suburb they own property in — several
properties in the *same* suburb still collapse to one draft, but properties
in different suburbs each get their own draft, addressed with that suburb's
own merge fields. See the status report for real counts of how often owners
repeat within a suburb.

## Weekly workflow

1. Drop this week's suburb PDF(s) in this folder, named `<suburb>.pdf`.
2. Run `python market_update_agent.py --dry-run` and check the preview —
   it lists exactly which suburbs have a letter and will draft, and which
   are being skipped for having none this week.
3. Once satisfied, real drafting is gated separately (see the top of
   `market_update_agent.py`) — nothing sends automatically either way.
4. Leave last week's PDFs in place or remove them; a suburb with no PDF is
   simply skipped, so there's no need to clean up between runs.

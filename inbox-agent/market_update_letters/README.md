# Market update letters

One file per segment. `market_update_agent.py` loads the file matching the
segment key, fills in the merge fields below for each recipient, and creates
one draft per owner. It never sends anything and never writes to this folder.

Segments currently wired up (see `SEGMENTS` in `market_update_agent.py`):

| Segment key         | File                          | Workbook tab       |
|----------------------|-------------------------------|---------------------|
| `lower_north_shore`  | `lower_north_shore.html`      | Master Database    |
| `upper_north_shore`  | `upper_north_shore.html`      | Upper North        |

Northern Beaches and City Fringe aren't here yet — there's no tab or data for
either in `DATA BASE.xlsx` right now. See the status report for what's needed.

## File format

```
Subject: Market update for {{suburb}}

<p>Hi {{first_name}},</p>

<p>... your letter, in HTML ...</p>
```

- **Line 1 must be `Subject: ...`** — the email subject, merge fields allowed.
- A blank line, then the body as HTML (`<p>`, `<br>`, `<strong>`, etc. — this
  becomes the draft's HTML body directly).
- Until this file exists and has real content, the agent skips that segment
  entirely and drafts nothing for it — it will never invent a letter on your
  behalf.

## Merge fields

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

An owner gets one draft per suburb they own property in, within a segment —
several properties in the *same* suburb still collapse to one draft, but
properties in different suburbs each get their own draft (addressed with
that suburb's own merge fields). See the status report for real counts of
how often owners repeat within a suburb.

## Placeholder

Until you write the real letter, leave (or restore) this exact line
somewhere in the file — the agent treats it as "not written yet" and skips
the segment rather than drafting a placeholder letter:

```
[EDDIE: WRITE THIS LETTER]
```

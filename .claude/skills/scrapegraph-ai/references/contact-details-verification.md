# Verifying company contact details

Contact details are the highest-stakes field a scraper produces. A wrong sofa dimension
annoys someone; a wrong complaints number sends a person with a financial dispute to a
stranger, or to a scammer who bought the abandoned number. Scam call-centres actively
seed fake "customer service" numbers into search results and directory sites, so a
naive scrape doesn't just risk being stale — it risks being poisoned.

An LLM extracting from a web page cannot promise correctness. What you *can* build is a
pipeline where every published detail was seen on the company's own domain, corroborated,
machine-validated, and stamped with the date it was last confirmed — and where anything
short of that is withheld rather than shown. That is the achievable version of "100%
correct", and it is worth being explicit with stakeholders that this is what is on offer.

## 1. Source hierarchy — never treat sources as equal

Rank every observation. The tier decides whether it can be published at all.

| Tier | Source | Use |
|---|---|---|
| 1 | The company's own website, on its registered domain — `/contact`, `/complaints`, footer, legal/imprint page | Publishable. The only tier that can stand alone. |
| 1 | A statutory or regulator register (in ZA: FSCA FSP register, CMS medical-scheme register, NCR register, CIPC) | Publishable, and authoritative for the legal entity name and registration number. |
| 2 | Ombudsman / industry-body scheme pages listing the member's escalation contact | Publishable for escalation contacts; corroborates tier 1. |
| 3 | The company's verified social profile or app-store listing | Corroboration only. |
| 4 | Directories, aggregators, review sites, search snippets, AI summaries | Never publish from these. Use only to *flag* that a tier-1 value may have changed. |

The rule that does the most work: **a contact detail is only publishable if it was seen on
a tier-1 source.** Directory data is a tripwire, not a source.

## 2. Fetch the tier-1 page in a way you can audit

```python
class ContactDetail(BaseModel):
    value: str = Field(description="Exactly as printed on the page, no reformatting")
    label: str = Field(description="The page's own wording, e.g. 'Client Care', 'Complaints'")
    context: str = Field(description="The surrounding sentence, verbatim, for audit")

class CompanyContacts(BaseModel):
    legal_name: Optional[str]
    phones: List[ContactDetail] = Field(default_factory=list)
    emails: List[ContactDetail] = Field(default_factory=list)
    complaints_url: Optional[str]
    physical_address: Optional[str]
    hours: Optional[str]
```

Asking for `context` verbatim is the single most useful trick here: it lets you (or a
reviewer) confirm the number really was labelled "complaints" and not "sales", and it makes
hallucination obvious, because invented values rarely come with a coherent surrounding
sentence that also appears in the raw HTML. Keep the raw HTML — you will want to re-check
against it later without re-fetching.

Add to the prompt: *"Copy values exactly as printed. If the page does not state a value,
leave it null — never infer, complete, or reformat a number."* Models are strongly inclined
to 'helpfully' complete a partial phone number.

## 3. Machine validation before anything else

Run `scripts/validate_contacts.py` in this skill. It is deterministic and cheap, and it
catches the failure modes an LLM produces:

- **Format** — phone normalises to E.164; email matches a strict pattern; URLs are https.
- **Domain binding** — the email's domain must match the company's own registrable domain
  (or a declared alias). `support@capitecbank.co.za` on capitecbank.co.za is coherent;
  a gmail address on a bank's contact page is a red flag, not a contact.
- **Cross-source agreement** — the same normalised value seen on ≥2 independent sources.
- **Provenance completeness** — every value has a source URL, tier, and timestamp, or it
  cannot be scored.

Formatting checks catch transcription errors; the domain and agreement checks are what
catch *poisoning*.

## 4. Confidence tiers, and publish only the top one

| Status | Condition | App behaviour |
|---|---|---|
| `verified` | Tier-1 source, format valid, domain-coherent, corroborated by a second independent source, confirmed within the refresh window | Show it. |
| `needs_review` | Tier-1 but uncorroborated, or sources disagree, or stale | Hold in a review queue. Don't show. |
| `rejected` | Fails format/domain checks, or only tier-3/4 sources | Never show. |

For anything `needs_review` or `rejected`, the app should link to the company's contact page
rather than print a number. "We link you to their official contact page" is honest and safe;
a stale number presented as fact is neither.

## 5. A human confirms before first publication

Automated checks establish *coherence*, not *truth*. For a first publication — and for any
change to an existing verified value — a person should open the tier-1 source and confirm.
This is bounded work: a few dozen companies is an afternoon, and it only recurs when a
value actually changes. Record who confirmed it and when; that record is what lets you
answer a complaint about a wrong number.

Where a company relationship exists, the strongest version is to send the record to the
company and ask them to confirm it. Many will, and a confirmed-by-the-company flag is worth
more than any amount of scraping.

## 6. Keep it fresh, and show the date

Contact details decay — call centres get renumbered, complaints addresses get retired.

- Re-verify tier-1 sources monthly; escalation/ombudsman contacts quarterly.
- If re-verification fails (page moved, value gone), demote to `needs_review` immediately
  rather than keeping the last known value.
- Store `last_verified_at` and **display it** next to the detail. A visible "verified 3 days
  ago" both earns trust and correctly transfers a little of the risk to a dated claim.
- Keep full history. When a value changes, you want to know what it was, when it changed,
  and what page said so.

## 7. Give users a correction path

Even a good pipeline will be wrong occasionally. A "report incorrect details" control that
files into the same review queue closes the loop, and users find stale numbers faster than
any crawler. Treat a user report as a tier-4 signal: it triggers re-verification, it doesn't
overwrite the value.

## Suggested storage shape

```sql
create table company_contacts (
  id uuid primary key default gen_random_uuid(),
  company_id uuid not null references providers(id) on delete cascade,
  kind text not null,               -- phone_general | phone_complaints | email_complaints | contact_url | address
  value text not null,              -- normalised (E.164 for phones)
  value_as_printed text,            -- what the page actually showed
  label text,                       -- the page's own wording
  status text not null default 'needs_review',   -- verified | needs_review | rejected
  source_url text not null,
  source_tier int not null,
  corroborating_sources jsonb not null default '[]'::jsonb,
  first_seen_at timestamptz not null default now(),
  last_verified_at timestamptz,
  verified_by text,                 -- 'human:<name>' | 'company_confirmed' | 'auto'
  raw_context text,                 -- verbatim surrounding sentence
  unique (company_id, kind, value)
);
create index on company_contacts (company_id, status);
```

The app reads only `status = 'verified'`. Everything else exists so a human can adjudicate
it, and so you can prove later why you published what you published.

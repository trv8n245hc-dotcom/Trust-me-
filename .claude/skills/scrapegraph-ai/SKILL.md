---
name: scrapegraph-ai
description: 'Extract structured data from websites with ScrapeGraphAI, the LLM-powered Python scraping library — describe the fields you want in a Pydantic schema, get clean JSON, no CSS selectors. Use whenever the task involves pulling product, catalogue, listing, pricing, spec or image data off retailer or supplier sites; gathering company details such as contact numbers, complaints addresses or regulator records; building an ingestion pipeline for a product or company database; scraping sites with no API; or requests phrased as "scrape", "crawl", "import this catalogue" or "find the contact details for these companies". Especially relevant for interior design and home-goods apps populating a catalogue from many retailers, and for directory or comparison apps publishing company contact details — it carries a verification playbook and validator script for those, since wrong contact details are a safety problem. Reach for it even when ScrapeGraphAI is not named: turning web pages into database rows is the trigger.'
---

# ScrapeGraphAI

ScrapeGraphAI turns "here is a URL and the fields I want" into JSON. It fetches the
page (Playwright), converts it to text/markdown, chunks it, and asks an LLM to fill in
your schema. You never write a CSS selector, which is the whole point: retailer sites
redesign constantly, and a prompt survives a redesign where `div.product-grid > span.price`
does not.

The trade-off is that every page costs LLM tokens and a few seconds. Everything below is
about getting reliable data out of it without burning money or getting blocked.

## Setup

```bash
pip install scrapegraphai
playwright install chromium      # required — the fetcher is a real browser
export SCRAPEGRAPHAI_TELEMETRY_ENABLED=false   # the library phones home by default
```

**Python ≥3.12 is required** (`requires-python = ">=3.12,<4.0"`). It pulls in LangChain 1.x
and Playwright, so give it its own virtualenv rather than sharing one with an app runtime.

ScrapeGraphAI is Python-only. In a JS/TS app (React, Vite, Next, Supabase Edge Functions)
it cannot run in the app process — it belongs in a separate scheduled Python job that
writes to your database, or behind a small internal HTTP endpoint. Plan for that boundary
early; it changes the architecture.

## Pick the right graph

Each "graph" is a prebuilt pipeline. Choosing the cheapest one that does the job is the
single biggest lever on cost and latency.

| Situation | Graph | Notes |
|---|---|---|
| One known page, want fields from it | `SmartScraperGraph` | The default. Start here. |
| A list of known pages, same fields | `SmartScraperMultiGraph` | Parallel LLM calls; `SmartScraperMultiBatchGraph` uses the OpenAI Batch API (much cheaper, slower). |
| Page is mostly clean already | `SmartScraperLiteGraph` | Skips the RAG/chunking stage. Cheaper, but loses detail on long pages. |
| Don't know the URLs yet | `SearchGraph` | Runs a web search (ddgs by default), then scrapes the top `max_results`. |
| Need to follow links | `DepthSearchGraph` | Set `depth`, and `only_inside_links: True` to stay on-domain. |
| Need what's in the images | `OmniScraperGraph` | Vision model reads up to `max_images` per page. Expensive — use deliberately. |
| Just want clean markdown | `MarkdownifyGraph` | No extraction, no schema. |
| Source is a local CSV/JSON/XML/PDF/docx | `CsvScraperGraph`, `JsonScraperGraph`, `XmlScraperGraph`, `DocumentScraperGraph` | Supplier price lists arrive as files more often than you'd think. |
| You'll scrape this same layout thousands of times | `ScriptCreatorGraph` / `CodeGeneratorGraph` | Emits a plain Python scraper. Run the LLM once, then run free forever. |

That last row deserves emphasis. If you're ingesting 5,000 product pages from one retailer,
paying an LLM 5,000 times is wasteful — use `SmartScraperGraph` on three sample pages to
confirm the fields, then `ScriptCreatorGraph` to generate a selector-based scraper for the
rest, and fall back to the LLM only when the generated script starts returning nulls.

The full catalogue of graphs, nodes and LLM providers is in
`references/graphs-and-providers.md` — read it when the table above doesn't cover your case.

## The core recipe: always pass a schema

Without a schema you get an arbitrary dict whose keys drift between runs — fine for
exploration, useless for writing to a database. A Pydantic schema pins the shape, and the
`Field(description=...)` text is read by the model, so it doubles as per-field instructions.

```python
from typing import List, Optional
from pydantic import BaseModel, Field
from scrapegraphai.graphs import SmartScraperGraph

class Product(BaseModel):
    name: str = Field(description="Product name as shown on the page")
    price: Optional[float] = Field(description="Current price as a number, no currency symbol")
    currency: Optional[str] = Field(description="ISO 4217 code, e.g. ZAR, USD")
    material: Optional[str] = Field(description="Primary material, e.g. solid oak, boucle, brass")
    dimensions_cm: Optional[str] = Field(description="W x D x H in centimetres")
    colours: List[str] = Field(default_factory=list, description="Colourways offered")
    image_urls: List[str] = Field(default_factory=list, description="Absolute URLs of product photos")
    in_stock: Optional[bool] = None

class Catalogue(BaseModel):
    products: List[Product]

config = {
    "llm": {"api_key": os.environ["OPENAI_API_KEY"], "model": "openai/gpt-4o-mini"},
    "headless": True,
    "verbose": False,
    "timeout": 480,
}

graph = SmartScraperGraph(
    prompt="Extract every product in the listing with its price, material, dimensions and photos.",
    source="https://example-furniture.com/collections/sofas",
    schema=Catalogue,
    config=config,
)
result = graph.run()          # dict matching Catalogue
print(graph.get_execution_info())   # token counts and per-node cost
```

Three habits that matter:

- **Make every uncertain field `Optional`.** A missing price is normal; a schema that
  forbids it pushes the model to invent one. Nulls you can detect and retry; hallucinated
  prices silently poison the catalogue.
- **Say what "not present" means in the prompt** — "leave a field null if the page does not
  state it" measurably reduces confabulation.
- **Log `get_execution_info()`** on every run. It's how you find out that one retailer's
  pages cost 40× the others because they inline their entire CSS.

## Model and provider config

`model` is `"provider/model-name"`. Supported providers: `openai`, `anthropic`,
`azure_openai`, `google_genai`, `google_vertexai`, `bedrock`, `groq`, `mistralai`,
`ollama`, `deepseek`, `fireworks`, `togetherai`, `xai`, `nvidia`, `hugging_face`,
`minimax`, `ernie`, `oneapi`, `clod`.

```python
"llm": {"api_key": "...", "model": "openai/gpt-4o-mini"}                  # cheap workhorse
"llm": {"api_key": "...", "model": "anthropic/claude-sonnet-4-5"}         # harder pages
"llm": {"model": "ollama/llama3.2", "model_tokens": 8192, "format": "json"}  # local, free
```

Use a small model for bulk extraction and reserve a stronger one for pages that fail
validation — a two-tier pass is usually 5–10× cheaper than running everything on the big
model. If a model isn't in the library's token table, pass `model_tokens` yourself or
chunking will be sized wrong.

## Other config keys worth knowing

| Key | Default | Why you'd change it |
|---|---|---|
| `headless` | `True` | `False` to watch the browser while debugging a blank page. |
| `verbose` | `False` | `True` prints each node as it runs — the fastest way to see where a run stalls. |
| `timeout` | `480` | Per-fetch seconds. Lower it in a nightly job so one hanging site can't eat the window. |
| `cache_path` | `False` | Path enables caching — set it while iterating on prompts so you re-fetch nothing. |
| `loader_kwargs` | `{}` | Playwright options: `proxy`, `slow_mo`, extra headers. |
| `storage_state` | — | Playwright storage state file for pages behind a login (trade/dealer pricing). |
| `max_results` | `3` | `SearchGraph`/multi-graphs: how many results to scrape. |
| `max_images` | `5` | `OmniScraperGraph`: images sent to the vision model. Each one costs. |
| `depth`, `only_inside_links` | `1`, `False` | `DepthSearchGraph` crawl bounds. Always set `only_inside_links: True` unless you truly want the open web. |

`source` also accepts an HTML string, not just a URL — so you can fetch pages yourself
(with your own retry/proxy logic) and hand the HTML to ScrapeGraphAI purely as an
extractor. For a production pipeline this separation is usually the right call: fetching
and extraction fail for different reasons and want different retry policies.

## Patterns for a product-catalogue pipeline

**Two-stage ingestion.** Listing pages give you URLs; detail pages give you fields. Run
`SmartScraperGraph` over the collection page for links, then `SmartScraperMultiGraph` over
the detail URLs. Trying to get full specs off a listing page yields shallow data.

**Normalise after extraction, not in the prompt.** Ask for the price as written; convert
currency, parse dimensions, and map colour names to your palette in ordinary Python you can
unit-test. Asking the LLM to also normalise makes failures untraceable.

**Store provenance with every row.** Source URL, scrape timestamp, model used, and the raw
LLM output. When a designer asks "why does this say the sofa is 240cm when the site says
220", you need to see what the page said that day. This also lets you re-run extraction on
cached raw HTML when you improve a prompt, without re-fetching.

**Diff, don't overwrite.** On re-scrape, compare against the stored row and write a change
record for price/stock moves. That history is often more valuable to the app than the
current value — it powers "price dropped" and "back in stock" features.

**Validate before it reaches the app.** A cheap gate catches most of what goes wrong:
price within an order of magnitude of the last known value, image URLs absolute and
resolving, required fields non-null, dimensions plausible. Quarantine failures for a second
pass with a stronger model instead of writing them through.

**Images.** Product photos are usually in the HTML, so plain `SmartScraperGraph` gets the
URLs cheaply — reach for `OmniScraperGraph` only when the information is *inside* the image
(a dimensions diagram, a spec sheet render, a fabric swatch you need described). Re-host
images you rely on; retailer CDNs rotate URLs.

## Contact details are a different problem

Company contact details — call-centre numbers, complaints addresses, escalation contacts —
need a higher standard than product data, and no prompt gets you there. A wrong sofa
dimension annoys someone; a wrong complaints number sends a person with a live dispute to a
stranger, or to a scammer who bought the abandoned number. Scam operations deliberately seed
fake "customer service" numbers into directories and search results, so the risk is not only
staleness — it is poisoning.

The honest position to take with stakeholders: an LLM scraper cannot promise correctness,
but a pipeline can promise that every published detail was seen on the company's own domain,
corroborated by a second source, machine-validated, and stamped with the date it was last
confirmed — and that anything failing those tests is withheld rather than shown.

Two rules carry most of the weight:

- **Only publish what a tier-1 source said** — the company's own registered domain, or a
  statutory register. Directories, aggregators and search snippets are tripwires that tell
  you to go re-check the tier-1 source; they are never the source itself.
- **Ask the model for the surrounding sentence verbatim** alongside each value. It proves
  the number was labelled "complaints" and not "sales", and invented values rarely arrive
  with coherent surrounding text that also appears in the raw HTML.

Then run `scripts/validate_contacts.py` on the extraction. It is deterministic and
network-free: it normalises phone numbers to E.164, rejects free-mail addresses and emails
whose domain doesn't belong to the company, counts independent corroborating sources,
enforces a freshness window, and returns `verified` / `needs_review` / `rejected` per value
with the reasons attached. Publish only `verified`; for everything else, link the user to
the company's contact page rather than printing a number you can't stand behind. Run
`--self-test` after changing the rules.

`references/contact-details-verification.md` has the full playbook: the source hierarchy,
the extraction schema, confidence tiers, the human confirmation step, refresh cadence, and a
storage shape that keeps provenance so you can prove later why you published what you did.

## Before you scrape someone's site

Check `robots.txt` and the site's terms — ScrapeGraphAI ships a `RobotsNode` and the
library's own README states it's intended for research and data exploration. Prefer an
official API, affiliate feed, or product data feed when one exists; retailers often provide
one and it's cheaper and more reliable than scraping. Rate-limit yourself, identify your
crawler honestly, and cache aggressively so you fetch each page once. If a supplier
relationship is involved, ask them for a feed before scraping them — it usually works.

The OSS library gives you no proxy rotation or anti-bot handling; that's `loader_kwargs`
and your own infrastructure. If you find yourself building anti-bot evasion, that is a
signal the site doesn't want to be scraped — reconsider rather than escalate.

## Troubleshooting

| Symptom | Cause and fix |
|---|---|
| Every field null / "NA" | Page is client-rendered and content hadn't loaded. Add `loader_kwargs: {"slow_mo": 2000}`, or check with `headless: False`. |
| `Provider X is not supported` | `model` needs the `provider/model` form, or the provider isn't in the supported list. |
| `model_tokens not specified` | Model unknown to the token table — pass `model_tokens` in the llm config. |
| Truncated results on long pages | Chunking dropped content. Use a larger-context model, or scrape detail pages individually instead of one giant listing. |
| Output shape changes between runs | You're not passing `schema=`. Pass one. |
| Slow, expensive runs | Set `cache_path` while developing; move stable layouts to `ScriptCreatorGraph`; drop to `SmartScraperLiteGraph` where pages are simple. |
| HTTP errors silently swallowed | Fixed in 2.2.2 — upgrade; older versions answered "NA" instead of surfacing the error. |

## Managed API alternative

The same team sells a hosted API (`scrapegraph-py` / `scrapegraph-js`, key `SGAI_API_KEY`)
that handles browsers, proxies, anti-bot, crawling and scheduled monitoring, billed per
credit. It has a JS SDK, so unlike the OSS library it can be called from a TypeScript
backend. Worth raising with the user when the pipeline needs anti-bot handling or scheduled
crawls and they'd rather not run infrastructure — but the OSS library is MIT and runs on
local models, so it stays the better default for cost control and data that shouldn't leave
your infrastructure.

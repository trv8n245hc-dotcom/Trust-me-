# ScrapeGraphAI reference — graphs, nodes, providers, config

Read this when the decision table in SKILL.md doesn't cover your case. Verified against
scrapegraphai **v2.2.2**.

## Contents
- [Graph catalogue](#graph-catalogue)
- [Nodes](#nodes)
- [LLM providers](#llm-providers)
- [Config keys](#config-keys)
- [Document loaders](#document-loaders)
- [Runtime notes](#runtime-notes)

## Graph catalogue

All graphs take `(prompt, source, config, schema=None)` and expose `.run()` plus
`.get_execution_info()`.

**Single page**
| Graph | Behaviour |
|---|---|
| `SmartScraperGraph` | Fetch → parse → (RAG) → generate answer. The default. Has a conditional retry node that regenerates with additional info when the first answer looks thin. |
| `SmartScraperLiteGraph` | Fetch → parse → answer. No RAG stage; cheaper, weaker on long pages. |
| `OmniScraperGraph` | Adds image-to-text over up to `max_images` images. Needs a vision-capable model. |
| `ScreenshotScraperGraph` | Screenshots the page and extracts from the image. For canvas/anti-text-extraction pages. |
| `MarkdownifyGraph` | Returns clean markdown, no LLM extraction step. |
| `SpeechGraph` | Extraction plus text-to-speech output. |

**Multiple pages**
| Graph | Behaviour |
|---|---|
| `SmartScraperMultiGraph` | Runs `SmartScraperGraph` over a list of sources, then merges answers. `max_results` bounds it. |
| `SmartScraperMultiLiteGraph` | Same, on the lite pipeline. |
| `SmartScraperMultiConcatGraph` | Concatenates rather than LLM-merging results — preserves per-source rows instead of blending them. Usually what you want for catalogue ingestion. |
| `SmartScraperMultiBatchGraph` | Uses the OpenAI Batch API. Much cheaper, results arrive later. Good for nightly jobs. |
| `SearchGraph` | Search (ddgs by default; `search_engine` and `serper_api_key` configurable) → scrape top `max_results`. |
| `OmniSearchGraph` | SearchGraph with vision. |
| `SearchLinkGraph` | Returns relevant links from a page rather than extracted fields. |
| `DepthSearchGraph` | Crawls to `depth`, honouring `only_inside_links`; `cut` and `force` tune link filtering. |

**Local files**
`CsvScraperGraph`, `JsonScraperGraph`, `XmlScraperGraph`, `DocumentScraperGraph` (PDF/docx),
plus `*MultiGraph` variants of each.

**Code generation**
| Graph | Behaviour |
|---|---|
| `ScriptCreatorGraph` | Emits a Python scraping script for the page. |
| `ScriptCreatorMultiGraph` | Merges scripts across several sources. |
| `CodeGeneratorGraph` | Generates and validates extraction code against your schema — the durable option for high-volume, stable layouts. |

## Nodes

Compose your own pipeline with `BaseGraph` when no prebuilt graph fits:

`FetchNode`, `FetchNodeLevelK`, `FetchScreenNode`, `ParseNode`, `ParseNodeDepthK`,
`RagNode`, `GenerateAnswerNode` (+ `Csv`/`Omni`/`FromImage`/`KLevel` variants),
`BatchGenerateAnswerNode`, `MergeAnswersNode`, `ConcatAnswersNode`,
`MergeGeneratedScriptsNode`, `GenerateScraperNode`, `GenerateCodeNode`,
`SearchInternetNode`, `SearchLinkNode`, `SearchNodeWithContext`, `GraphIteratorNode`,
`ConditionalNode`, `RobotsNode`, `HtmlAnalyzerNode`, `PromptRefinerNode`, `ReasoningNode`,
`DescriptionNode`, `GetProbableTagsNode`, `ImageToTextNode`, `TextToSpeechNode`,
`MarkdownifyNode`.

`RobotsNode` checks robots.txt before fetching — include it in any pipeline that touches
sites you don't own.

## LLM providers

`model` is `"provider/model-name"`. Recognised providers:

`openai`, `azure_openai`, `anthropic`, `google_genai`, `google_vertexai`, `bedrock`,
`groq`, `mistralai`, `ollama`, `hugging_face`, `deepseek`, `fireworks`, `togetherai`,
`xai`, `nvidia`, `minimax`, `ernie`, `oneapi`, `clod`.

Notes:
- Omitting the `provider/` prefix makes the library guess from its token table and pick the
  first match — ambiguous and silent. Always write the prefix.
- Unknown models fall back to an 8192-token assumption unless you pass `model_tokens`.
  Wrong token limits mean wrong chunking, which shows up as truncated extractions.
- You can bypass the factory entirely with `"llm": {"model_instance": <LangChain model>,
  "model_tokens": 128000}` — the escape hatch for a provider or gateway the library
  doesn't know.

## Config keys

| Key | Default | Meaning |
|---|---|---|
| `llm` | required | Model config dict. |
| `verbose` | `False` | Per-node logging. |
| `headless` | `True` | Playwright headless mode. |
| `timeout` | `480` | Per-fetch timeout, seconds. |
| `cache_path` | `False` | Path enables fetch caching. |
| `loader_kwargs` | `{}` | Passed to the Chromium loader: `proxy`, `slow_mo`, headers. |
| `storage_state` | — | Playwright storage-state file for authenticated pages. |
| `browser_base` / `scrape_do` | — | Third-party managed-browser backends. |
| `max_results` | `3` | Search and multi graphs. |
| `max_images` | `5` | `OmniScraperGraph`. |
| `depth`, `only_inside_links`, `cut`, `force` | `1`, `False`, `True`, `False` | `DepthSearchGraph` crawl controls. |
| `search_engine`, `serper_api_key` | ddgs | `SearchGraph` backend. |
| `embedder_model` | — | Embeddings for the RAG stage. |
| `additional_info` | — | Extra text prepended to the extraction prompt. |
| `burr_kwargs` | — | Burr tracing integration. |

## Document loaders

`scrapegraphai/docloaders/`: `chromium.py` (default Playwright fetcher, with
`undetected-playwright` support), `browser_base.py`, `scrape_do.py`, `plasmate.py`.

`source` accepts a URL, a local file path, or a raw HTML string. Passing HTML you fetched
yourself is the cleanest way to separate fetch retries from extraction retries.

## Runtime notes

- Python ≥3.12, <4.0. `playwright install chromium` is required after pip install.
- Telemetry is on by default: `SCRAPEGRAPHAI_TELEMETRY_ENABLED=false`.
- 2.2.2 fixed HTTP errors being reported as "NA" answers — on older versions a 403 looks
  like an empty page, which is a silent data-quality bug. Pin ≥2.2.2.
- MIT licensed. The library's own README states it is intended for data exploration and
  research; respect robots.txt and site terms.

# Slinky Precedent Research Plugin

Slinky is a structured knowledge base of 100,000+ multi-year corporate episodes across active and delisted public companies. Instead of sifting through disclosures, you can natively search by strategic precedent—finding companies that adopted similar strategic moves, capital-structure shifts, or M&A bets across industries, peeking into their financial position before and after a strategic event, and getting the source filings behind any precedent. Defaults to crisp chat synthesis with optional 4-tab `.xlsx` Precedent Matrix workbooks in Claude (`claude.ai`, Claude Desktop, Claude in Excel, Claude Code).

## Requirements & Data Handling

Connects to Slinky's hosted remote MCP server (`https://mcp.slinky.so/mcp`), an OAuth 2.1 Protected Resource Server.

- **Authentication:** Requires an approved Slinky account. Contains no credentials, API keys, or local scripts. Claude discovers OAuth metadata via `/.well-known/oauth-protected-resource`, registers with PKCE (`S256`), and opens `https://slinky.so/oauth/consent`.
- **Data Sent:** Structured search parameters (tickers, strategic filters, research prompts) and OAuth bearer token only. No local files or environment variables are transmitted.
- **Links & Support:** [Privacy Policy](https://slinky.so/privacy) · [Terms of Service](https://slinky.so/terms) · [Setup & Docs](https://slinky.so/mcp/setup) · [aly@anlgs.com](mailto:aly@anlgs.com)

Maintained by Analogous ([slinky.so](https://slinky.so)).

## What Slinky Covers

Instead of sifting through raw filings, query Slinky by **strategic precedent**: find companies across industries that attempted similar moves, capital-structure shifts, or M&A bets, inspect their directional pre/post financial position (`FY` and standalone `Q` reported separately), and follow source EDGAR links. Slinky does not provide live market quotes, consensus estimates, or price targets.

---

## Components

### Remote MCP Server (12 Read-Only Tools & 3 Built-in Prompts)

Bundled via `.mcp.json` (`type: "http"`, `https://mcp.slinky.so/mcp`). All 12 tools are read-only (`readOnlyHint=True`).

| Tool | Purpose |
|------|---------|
| `query_arcs_by_predicates` | Filter precedents by strategic mechanism, execution vehicle, catalyst, intent, or ticker (`output_mode`: `chat` / `workbook` / `deep`, optional matrix export) |
| `query_situations_by_predicates` | Evidence-level SEC and news situations across 9 situation kinds and tickers |
| `query_canonicals_by_predicates` | Structural precedent cohorts ranked by matching case count and outcome mix |
| `list_predicate_vocabulary` | Enumerate active `v3` controlled filter vocabularies and aliases |
| `arc_finder` | Natural-language search for raw strategic situations (`action` + `motivation`) |
| `get_arc_details` | Full episode record for 1 ID or a batch of up to 6 `arc_ids`: chronological filings, quotes, EDGAR links, and unpooled `FY` vs. standalone `Q` financials |
| `search_canonical_patterns` | Natural-language search over recurring cross-industry strategic patterns |
| `find_pattern_exemplars` | Exemplar companies and documented outcomes for a structural pattern |
| `get_situations_by_pattern` | Filing and news situations belonging to a pattern |
| `get_company_history` | Unified multi-year strategic history for a company name or ticker |
| `get_recent_news_situations` | Recent `T+1` SEC filing disclosures with procedural noise filtering and optional watchlist export |
| `search_company_pages` | Semantic search over archived company foundational pages |

**Built-in MCP prompts:** `stress_test_company`, `precedent_cohort_sweep`, `watchlist_triage` (each supports `output_mode`: `chat` | `workbook` | `deep`).

### `/slinky:research` Skill

[`skills/research/SKILL.md`](skills/research/SKILL.md) routes three analyst workflows—**`T+1` watchlist triage**, **company/deal stress tests**, and **cross-industry pattern sweeps**—with depth routing:

- **`chat` (default):** Fast inline prose and compact tables (no `.xlsx` or `.docx`).
- **`deep` (opt-in):** Unpacks 2–4 contrasting precedents for a Condition Ledger, Negative Mappings, and Counterfactuals.
- **`workbook` (opt-in):** Pairs an executive cover note with a 4-tab `<slug>_precedent_comps.xlsx` workbook (`Cohort Summary`, `Precedent Matrix`, Strategy Context tab, and `SEC Filing Trail`) in Claude's spreadsheet sandbox or directly inside `Claude in Excel`.

---

## Examples

```text
/slinky:research this week
/slinky:research SNOW, CRWV, WYFI last 14 days
/slinky:research Stress-test CRWV funding GPU capex with long-term leases under high leverage
/slinky:research How have infrastructure operators fared externalizing internal compute capacity?
```

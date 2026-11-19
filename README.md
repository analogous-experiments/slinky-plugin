# Slinky Precedent Research Plugin

Connects Slinky's knowledge base of 105,000+ historical SEC corporate episodes, structural cohorts, and filing-backed execution trajectories to analyst workflows in Claude (`claude.ai`, Claude Desktop, Claude in Excel, Claude Code). Pairs concise institutional precedent analysis with 4-tab `.xlsx` Precedent Matrix workbooks.

## Requirements & Data Handling

The plugin connects to Slinky's hosted remote MCP server (`https://slinky-mcp-production.up.railway.app/mcp`), an authenticated OAuth 2.1 Protected Resource Server.

- **Authentication:** Requires an approved Slinky account (design partner or paid plan). This plugin contains no credentials, no API keys, and no local executable scripts. On first use, Claude discovers Slinky's OAuth 2.1 server via `/.well-known/oauth-protected-resource`, registers dynamically with PKCE (`S256`), and opens `https://slinky.so/oauth/consent` for approval with rotating refresh tokens.
- **Data Sent:** Only your structured search/query parameters (tickers, strategic move filters, natural-language research prompts) and OAuth bearer token are sent to `https://slinky-mcp-production.up.railway.app/mcp` to retrieve read-only public SEC/KB records. No local machine files or environment variables are transmitted.
- **Policies & Support:** [Privacy Policy](https://slinky.so/privacy) · [Terms of Service](https://slinky.so/terms) · [Setup & Docs](https://slinky.so/mcp/setup) · [aly@anlgs.com](mailto:aly@anlgs.com)

Maintained by Analogous ([slinky.so](https://slinky.so)).

## What Slinky Does (and Doesn't)

Slinky indexes **structural strategic precedents and filing-backed execution histories** across active and delisted SEC issuers: how companies expanded platforms, diversified products, navigated debt/capex cycles, executed M&A, or pivoted across industries — complete with verbatim SEC/news evidence quotes, unpooled historical financial windows (`annual_fy` vs. `quarterly_standalone` `pre_arc` / `during_arc` / `post_arc` snapshots), and one-shot Precedent Matrix / Disclosure Wire JSON exports for `.xlsx` workbook compilation. It does not cover live market quotes, consensus earnings estimates, or price targets.

---

## Plugin Components

### Remote MCP Server (12 Read-Only KB Tools)

The plugin bundles `.mcp.json` (`type: "http"`) pointing to `https://slinky-mcp-production.up.railway.app/mcp`. Every tool exposes a human-readable `title` and read-only `ToolAnnotations` (`readOnlyHint=True`, `destructiveHint=False`, `idempotentHint=True`, `openWorldHint=False`).

| Tool | Purpose |
|------|---------|
| `query_arcs_by_predicates` | Predicate-first precedent retrieval across active `v3` mechanisms, condition cases (`mechanism_is_null=True` + `execution_vehicle`), catalysts, strategic intents, and tickers — with full-cohort base rates and optional `include_matrix_export=True` (`25–50` period-locked rows + EDGAR trail) |
| `query_situations_by_predicates` | Evidence-level SEC and news situation retrieval across all 9 `situation_kind` values, execution vehicles, and tickers |
| `query_canonicals_by_predicates` | Retrieve structural precedent cohorts ranked by matching case count and cohort base rates |
| `list_predicate_vocabulary` | Enumerate active `v3` controlled vocabularies and legacy alias mappings |
| `arc_finder` | Natural-language precedent discovery for raw business situation fragments |
| `get_arc_details` | Full record for a single `arc_id` or concurrent batch `arc_ids=[...]` (up to 6 IDs): chronological situations, verbatim evidence quotes, resolved EDGAR URLs, SIC metadata, and unpooled `FY` vs. standalone `Q` financial windows |
| `search_canonical_patterns` | Natural-language structural cohort search |
| `find_pattern_exemplars` | Exemplar companies and documented outcomes within a structural pattern cohort |
| `get_situations_by_pattern` | Filing and news situations belonging to a pattern cohort |
| `get_company_history` | Unified strategic history for a company name or ticker symbol |
| `get_recent_news_situations` | Recent `T+1` SEC disclosures ordered by filing date, with procedural noise filtering and optional `include_wire_export=True` watchlist export |
| `search_company_pages` | Semantic search over crawled company foundational pages |

### Skill & Paired 4-Tab `.xlsx` Workbook Contract (`/slinky:research`)

[`skills/research/SKILL.md`](skills/research/SKILL.md) unifies three institutional workflows:
- **Mode A:** `T+1` SEC filing catalysts & multi-ticker watchlist triage
- **Mode B:** Company history, M&A / capital-structure stress tests & IC red-teams
- **Mode C:** Top-down cross-industry structural pattern & base-rate sweeps

When spreadsheet execution is available (Claude Desktop / `claude.ai` built-in `.xlsx` sandbox or native `Claude in Excel`), `/slinky:research` pairs a concise executive cover note with a 4-tab Precedent Workbook (`<slug>_precedent_comps.xlsx`):
1. **`Cohort Base Rates`** — Full-cohort server denominator + live cross-sheet `=COUNTIF`, `=AVERAGEIFS`, and `MEDIAN` formulas strictly separating `Annual Basis (FY)` from `Standalone Quarterly Basis (Q)`.
2. **`Precedent Matrix`** — `25–50` period-locked rows (including explicit `DELISTED` anti-survivorship tags, USD-million flows, decimal margins, live `bps`-delta formulas, blank-not-zero `None` cells, and clickable `=HYPERLINK` EDGAR links).
3. **Strategy Context Tab** — Polymorphically rendered as `Anchor Scorecard` (`Mode B`), `Sector Breakdown` (`Mode C`), or `Disclosure Wire` (`Mode A`).
4. **`SEC Filing Trail`** — Chronological non-procedural milestone filings with clickable EDGAR accession hyperlinks.

---

## Quick-Start Examples

```text
/slinky:research this week
/slinky:research SNOW, CRWV, WYFI last 14 days
/slinky:research Stress-test CRWV funding GPU capex with long-term leases under high leverage
/slinky:research How have infrastructure operators fared externalizing internal compute capacity?
```

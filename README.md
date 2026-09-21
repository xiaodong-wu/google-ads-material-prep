# Google Ads Material Prep

Prepare Google Ads campaign materials from a target website and multiple seed keywords. The skill produces **eight Excel workbooks and one HTML report**, using English ad copy and Chinese research commentary by default. It prepares local files; it does not create or publish ads.

## Installation

Install this skill and its required keyword dependency, [Google Ads Keywords to Sheets or CSV](https://github.com/xiaodong-wu/google-ads-keyword-to-sheets), in your Codex skills directory. For a fresh installation:

```bash
git clone https://github.com/xiaodong-wu/google-ads-material-prep.git "${CODEX_HOME:-$HOME/.codex}/skills/google-ads-material-prep"
git clone https://github.com/xiaodong-wu/google-ads-keyword-to-sheets.git "${CODEX_HOME:-$HOME/.codex}/skills/google-ads-keyword-to-sheets"
```

Restart Codex if the skills are not discovered immediately.

### Requirements

- A supported tool that can control a signed-in Chrome session, such as Codex's `mcp__cua_repl`. The older `chrome:control-chrome` skill is optional, not a mandatory dependency.
- A Google Ads account that can use Keyword Planner and download keyword ideas as CSV.
- Python 3.9 or later with `openpyxl` available in the environment used to run the workbook helper.

## Usage

In Codex, enter:

```text
Use $google-ads-material-prep to prepare all Google Ads materials for https://example.com/.
Seed keywords: product one, product two.
Keyword language: English. Location: All locations.
Create a separate worksheet for each seed keyword, then consolidate the campaign plan.
```

Step 01 queries real Google Ads data for each seed separately, exports all source metric columns, and then consolidates, deduplicates, and groups the results. Each seed retains its own worksheet, including a clearly identified empty or blocked result when applicable.

The target website is used only for business analysis and landing-page matching. It is never entered into Keyword Planner's website filter. The standard material-preparation workflow uses local CSV exports and does not require Google Sheets.

## Deliverables

By default, files are saved in a dedicated output directory within the current workspace, with separate subdirectories for unrelated brands. Filenames and audit-folder names follow the skill's existing naming convention.

| Step | File type | Contents |
| --- | --- | --- |
| 01 | Excel | Keyword plan: one worksheet per seed plus a consolidated plan. |
| 02 | Excel | Audience analysis across five dimensions, role profiles, a 3-4 sentence summary, and 2-3 recommendations. |
| 03 | Excel | Site hierarchy and five English descriptions per target section or important page. Each description is at most 90 characters, including spaces; approximately one third include a strong CTA. Each section requires at least two descriptions with evidence-backed numerical selling points. |
| 04 | Excel | Google Ads sitelink assets. |
| 05 | Excel | Google Ads callout assets. |
| 06 | Excel | Lead form settings and copy. |
| 07 | Excel | Structured snippet assets. |
| 08 | HTML | A styled, responsive, self-contained overseas market and trade report covering exactly five core dimensions. |
| 09 | Excel | Recommended advertising countries and regions, supported by the market report. |

The HTML report covers global market size and CAGR, Google search trends and geographic interest, import demand and purchasing markets, trade policies and compliance requirements, and competition and market concentration. Country tiers and bidding recommendations belong in the separate country-recommendation workbook.

## Evidence and validation

Sources, original keyword CSVs, per-seed execution status, and validation results are retained in the output's audit subdirectory. Missing data is marked as pending verification or blocked; the skill does not invent figures to fill gaps. Static validation does not establish Google Ads approval.

Description counts depend on the site's actual sections. Step 03 no longer uses a fixed total of 20 descriptions or a 70-character minimum. When numerical claims, certifications, or delivery commitments lack evidence, the skill reports the shortfall instead of fabricating promises.

See [SKILL.md](SKILL.md) for the complete workflow and [Export instructions](references/export.md) for input schemas and validation commands.

## Updating existing materials

Older step 02 and step 03 workbooks must be regenerated with the supplementary worksheets required by the current data contract. Updating the skill alone does not rewrite existing client files.

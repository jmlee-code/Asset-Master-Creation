# Confluence IT Purchase Scraper

**Version:** `v0.1.1`  
**Team:** KRAFTON EMEA Finance — IT Asset Management  
**Platform:** Windows (Python 3.12+, Tkinter GUI)

---

## Overview

An internal desktop tool that automatically scrapes the IT Purchase table from Confluence, enriches each row using Snipe-IT, SAP Vendor Master, and IT Team Members data, and generates a SAP-ready upload file (`.xls`) — all in a single click.

---

## Pipeline

```
[1/4] Scrape Confluence page  →  IT_Purchase_RAW_YYYYMMDD_HHMM.xlsx
         ↓  Filter by selected Year / Month (Purchase Date)
[2/4] Save filtered raw data
         ↓
[3/4] Read SAP template headers  +  Load Vendor Master  +  Load Team Members
         ↓
[4/4] Per-row processing:
        · Claude AI       → Asset Class (4-digit code)
        · Snipe-IT API    → Acquisition Value (purchase_cost) + Department
        · Vendor Master   → Vendor code  (Col C keyword → Col B code)
        · Team Members    → Employee No. (name → E-code, strip "E" prefix)
        · Rules           → Cost Center, Acquisition Date, Currency
         ↓
      IT_Purchase_SAP_Upload_YYYYMMDD_HHMM.xls
```

---

## Output Files

| File | Description |
|---|---|
| `IT_Purchase_RAW_*.xlsx` | Full raw scraped table (all rows, all columns) |
| `IT_Purchase_SAP_Upload_*.xls` | SAP upload file — filtered + enriched, template headers |

---

## SAP Field Mapping

| SAP Column | Source | Logic |
|---|---|---|
| **Asset Class** | Claude AI | 4-digit code — see classification rules below |
| **Asset Tag** | Confluence: `Snipe-IT Asset Tag` | Strip surrounding `__` underscores |
| **Description** | Confluence: `Item Description` | Direct mapping |
| **Acquisition Date** | Confluence: `Purchase Date` | First day of the month → `YYYY.MM.DD` |
| **Vendor** | Confluence: `Vendor` → Vendor Master | Keyword match on Col C → Col B code |
| **Acquisition Value** | Snipe-IT API | `purchase_cost` field via `/api/v1/hardware/bytag/{tag}` |
| **Currency** | — | Default: `EUR` |
| **Employee No.** | Confluence: `Kissflow Raised By` → IT Team Members | Name match → E-code, strip `E` prefix |
| **Asset No.** | — | Blank |
| **Sub-No** | — | Blank |
| **Cost Center** | Snipe-IT: `department` | `PBB` in dept name → `140069`, else `140006` |

---

## Asset Class Definitions (Hardcoded)

| Code | Label | Classification Criteria |
|---|---|---|
| `1011` | Office Equipment_PC/Monitor | Laptops, desktops, workstations, peripherals ≥ €1,000 (e.g. MacBook, Dell, RTX 5070) |
| `1012` | Office Equipment_Server&Network | Servers, switches, routers, firewalls, storage, rack (e.g. Cisco, Arista) |
| `1010` | Office Equipment_IT | IT devices excluding PC / Server / Network |
| `1020` | Office Equipment_General | Office furniture, appliances, consumables (e.g. chairs, desks, air purifiers) |
| `2101` | Software_Others | Software licenses, perpetual licenses, installed programs |
| `2900` | Assets under construction | Prepayments, ongoing projects, blank Snipe-IT Asset Tag |

---

## Input Files Required

| File | Description | Key Columns |
|---|---|---|
| `Asset registration master_1400_template.xls` | SAP upload template — header row defines output columns | Row 1 = column headers |
| `SAP Vendor Master_*.xlsx` | Vendor / place-of-purchase master | Col B = Code, Col C = Name |
| `IT_Team_Members.xlsx` | EMEA IT team employee list | Col B = Employee code (E-prefixed), Col C = Full name |

---

## GUI Layout

```
┌─ CONFLUENCE ──────────────────────────────────────────────┐
│  Page URL    [ https://krafton.atlassian.net/wiki/... ]   │
│  Period      [ 2026 ] 년  [ 01 - Jan ] 월  기준으로 필터링  │
└───────────────────────────────────────────────────────────┘
┌─ INPUT ────────────────────────────────────────────────────┐
│  SAP Template    [ path ]                      [ Browse ]  │
│  Vendor Master   [ path ]                      [ Browse ]  │
│  Team Members    [ path ]                      [ Browse ]  │
└───────────────────────────────────────────────────────────┘
┌─ OUTPUT ───────────────────────────────────────────────────┐
│  Output Folder   [ path ]                      [ Browse ]  │
└───────────────────────────────────────────────────────────┘
[ ▶  RUN ]  [ Clear Log ]                    [  progress  ]
┌─ LOG ──────────────────────────────────────────────────────┐
│  Live pipeline log with colour-coded status                │
└───────────────────────────────────────────────────────────┘
```

All input paths are **persisted across sessions** in `scraper_session.json` (auto-created in the same folder as the script).

---

## Installation

```bash
pip install requests pandas beautifulsoup4 openpyxl xlrd xlwt
```

Run:
```bash
python confluence_scraper_gui.py
```

---

## Configuration (Hardcoded in Script)

The following values are set at the top of the script and are not shown in the GUI:

| Constant | Description |
|---|---|
| `DEFAULT_EMAIL` | Atlassian account email for Confluence API auth |
| `DEFAULT_API_TOKEN` | Atlassian API token (generate at id.atlassian.com) |
| `CLAUDE_API_KEY` | Anthropic Claude API key |
| `SNIPEIT_BASE_URL` | Snipe-IT server URL (e.g. `http://100.68.20.67`) |
| `SNIPEIT_API_TOKEN` | Snipe-IT API token (generate in Profile → API Keys) |

---

## Confluence API Token

If you receive a `401 Unauthorized` error, your API token has expired.  
Generate a new one at: https://id.atlassian.com/manage-profile/security/api-tokens

Update `DEFAULT_API_TOKEN` in the script with the new token.

---

## Session Cache

On exit, the tool saves all current input paths to `scraper_session.json` in the working directory. On the next launch, these paths are restored automatically as default values.

To reset to factory defaults, delete `scraper_session.json`.

---

## Version History

| Version | Date | Changes |
|---|---|---|
| `v0.1.0` | 2026-03-17 | Year dropdown extended to 2099; version rename |
| `v0.1.1` | 2026-03-17 | Initial release — Confluence scraper + Tkinter GUI |

---

## Author

KRAFTON Europe B.V. — EMEA Finance  
Contact: jmlee@krafton.com

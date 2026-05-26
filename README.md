# 🔍 FUZZY — Intelligent Text Matching UDF for Excel (VBA)

> A custom Excel VBA User-Defined Function that automatically corrects misspelled, abbreviated, and alternate-form geographical names using a **6-layer matching engine**.

---

## 📌 Table of Contents

- [Overview](#-overview)
- [Demo](#-demo)
- [Features](#-features)
- [Getting Started](#-getting-started)
- [MASTER Sheet Setup](#️-master-sheet-setup)
- [Syntax & Parameters](#-syntax--parameters)
- [How It Works — 6-Layer Engine](#-how-it-works--6-layer-engine)
- [Examples](#-examples)
- [Helper Functions](#️-helper-functions)
- [MatchScore Algorithm](#-matchscore-algorithm)
- [Configuration](#-configuration)
- [File Structure](#-file-structure)
- [Contributing](#-contributing)
- [License](#-license)

---

## 🧩 Overview

Data collected from surveys, forms, and legacy systems rarely has clean, consistent geographic names. `FUZZY()` solves this in a single Excel formula — no VBA knowledge required to use it.

```excel
=FUZZY(A2, "S")   → corrects misspelled state name in cell A2
=FUZZY(B2, "D")   → corrects misspelled district name in cell B2
```

It handles:
- ✅ Typos and transpositions (`KERLA` → `Kerala`)
- ✅ Missing characters (`Rajastan` → `Rajasthan`)
- ✅ Abbreviations (`UP` → `Uttar Pradesh`)
- ✅ Legacy / alternate names (`Bombay` → `Maharashtra`)
- ✅ Prefix / partial inputs (`Maha` → `Maharashtra`)
- ✅ Mixed case, extra spaces, punctuation, special characters

---

## 🎬 Demo

| Input (raw)       | Formula             | Output           | Layer Matched     |
|-------------------|---------------------|------------------|-------------------|
| `maharashtra`     | `=FUZZY(A2,"S")`    | `Maharashtra`    | Exact Match       |
| `Bombay`          | `=FUZZY(A3,"S")`    | `Maharashtra`    | Wrong-Form Match  |
| `UP`              | `=FUZZY(A4,"S")`    | `Uttar Pradesh`  | Short-Form Match  |
| `Maha`            | `=FUZZY(A5,"S")`    | `Maharashtra`    | Prefix Match      |
| `Uttarpradsh`     | `=FUZZY(A6,"S")`    | `Uttar Pradesh`  | Fuzzy Score Match |
| `KERLA`           | `=FUZZY(A7,"S")`    | `Kerala`         | Fuzzy Score Match |
| `New Dlhi`        | `=FUZZY(B2,"D")`    | `New Delhi`      | Fuzzy Score Match |
| `XYZABC`          | `=FUZZY(A9,"S")`    | `XYZABC`         | No match (safe fallback) |

---

## ✨ Features

- **6-layer matching engine** — Exact → Wrong-Form → Short-Form → Prefix → Contains → Fuzzy Score
- **4 configurable data types** — `S` (States), `D` (Districts), `C` (Custom C), `T` (Custom T)
- **Auto-abbreviation expansion** — no hardcoded mappings; derives initials dynamically
- **Prefix / partial matching** — catches truncated inputs without needing the full name
- **Upgraded MatchScore algorithm** — position-weighted scoring with word-start and contains bonuses
- **Strict threshold (0.72)** — reduces false positives while handling real typos
- **Safe fallback** — returns original text when no confident match found; never silently corrupts data
- **Empty input guard** — returns `""` immediately for blank cells
- **Single formula** — just `=FUZZY(A2,"S")` and drag down; no helper columns needed
- **Extensible MASTER sheet** — add rows or wrong-form entries anytime; formulas update automatically

---

## 🚀 Getting Started

### Prerequisites

- Microsoft Excel (any version supporting VBA)
- Workbook saved as **`.xlsm`** (macro-enabled workbook)

### Installation

**Step 1 — Download the files**

```
git clone https://github.com/your-username/fuzzy-excel-udf.git
```

Or download `FUZZY.bas` directly from the [`/src`](./src/) folder.

**Step 2 — Open your workbook in Excel**

Save it as `.xlsm` if it isn't already (`File → Save As → Excel Macro-Enabled Workbook`).

**Step 3 — Import the VBA module**

1. Press `Alt + F11` to open the VBA Editor
2. In the Project Explorer, right-click your workbook name
3. Select `Import File...`
4. Choose `FUZZY.bas`
5. Click `OK` and close the VBA Editor

**Step 4 — Set up the MASTER sheet**

Create a sheet named exactly **`MASTER`** with 8 columns. See [MASTER Sheet Setup](#️-master-sheet-setup) for the full layout.

**Step 5 — Use the formula**

```excel
=FUZZY(A2, "S")    ← correct state name in A2
=FUZZY(B2, "D")    ← correct district name in B2
```

---

## 🗂️ MASTER Sheet Setup

The `MASTER` sheet is the reference table FUZZY reads. It must be named exactly **`MASTER`** and follow this 8-column structure:

| Column | Label               | Used When | Description |
|--------|---------------------|-----------|-------------|
| A (1)  | `State_Correct`     | `Typ="S"` | Official / canonical state names |
| B (2)  | `District_Correct`  | `Typ="D"` | Official / canonical district names |
| C (3)  | `CustomC_Correct`   | `Typ="C"` | Correct names for custom category C |
| D (4)  | `CustomT_Correct`   | `Typ="T"` | Correct names for custom category T |
| E (5)  | `State_Wrong`       | `Typ="S"` | Alternate / wrong state spellings |
| F (6)  | `District_Wrong`    | `Typ="D"` | Alternate / wrong district spellings |
| G (7)  | `CustomC_Wrong`     | `Typ="C"` | Wrong forms for custom category C |
| H (8)  | `CustomT_Wrong`     | `Typ="T"` | Wrong forms for custom category T |

### Sample MASTER Data

| A (State)       | B (District)  | C (Custom C)  | D (Custom T)  | E (State Wrong)              | F (District Wrong)       | G (C Wrong)     | H (T Wrong)            |
|-----------------|---------------|---------------|---------------|------------------------------|--------------------------|-----------------|------------------------|
| Maharashtra     | Mumbai        | North Zone    | Retail        | `Bombay,MH,Maharastra`       | `Bombay City,BOM`        | `N Zone,NZ`     | `Retl,Retail Segment`  |
| Uttar Pradesh   | Lucknow       | South Zone    | Wholesale     | `UP,U.P.,UttarPradesh`       | `Lknow,Lakhnaoo`         | `SZ,South`      | `WHOLSALE,WS`          |
| Tamil Nadu      | Chennai       | East Zone     | Manufacturing | `TN,Tamilnadu,Madras`        | `Madras,Chenai`          | `EZ,East`       | `Mfg,Manufact`         |
| Karnataka       | Bangalore     | West Zone     | Services      | `KA,Karanataka`              | `Bangalor,Bengaluru`     | `WZ,West`       | `Svc,Service Dept`     |

### MASTER Sheet Rules

- Row 1 must be a header row — data starts at Row 2
- Column A must be fully populated (used to detect last row)
- Columns A–D: correct/canonical spellings only
- Columns E–H: wrong/alternate forms (comma-separated values or one per cell)
- Columns C–D (and G–H) can be left empty if only `S` and `D` types are used
- The more wrong-form entries in E–H, the higher the accuracy of Layer 2 matching

> **Tip:** Adding known wrong forms to Columns E–H is the fastest way to improve match accuracy for your specific dataset.

---

## 📐 Syntax & Parameters

```vba
Public Function FUZZY(txt As String, Typ As String) As String
```

```excel
=FUZZY( txt , Typ )
```

### Parameters

| Parameter | Type   | Required | Description |
|-----------|--------|----------|-------------|
| `txt`     | String | ✅ Yes   | The raw input text to correct. Can contain typos, abbreviations, alternate spellings, or mixed case. Leading/trailing spaces are trimmed automatically. |
| `Typ`     | String | ✅ Yes   | The data type to match against. See values below. Case-insensitive (`"s"` works the same as `"S"`). |

### Typ Values

| Typ   | Searches         | Correct Col | Wrong-Form Col | Use For                     |
|-------|------------------|-------------|----------------|-----------------------------|
| `"S"` | States           | Column A    | Column E       | State / province names      |
| `"D"` | Districts        | Column B    | Column F       | District / sub-region names |
| `"C"` | Custom Category C| Column C    | Column G       | Zones, regions, departments |
| `"T"` | Custom Category T| Column D    | Column H       | Segments, types, categories |
| Other | —                | —           | —              | Returns original text immediately |

### Return Value

Returns a **String** — the best-matching correct name from the MASTER sheet.

- If `bestScore >= 0.72` → returns the matched correct name
- If `bestScore < 0.72` → returns `originalText` (the original input, untouched)
- If `txt` is empty after trimming → returns `""`

---

## 🔄 How It Works — 6-Layer Engine

Before any matching begins, FUZZY pre-processes the input:

```
originalText = Trim(txt)
txt = CleanText(txt)        ← normalise for comparison
if txt = "" → return ""
Set columns based on Typ
bestScore = 0
```

Then for every row in MASTER, it checks 6 layers **in order**, exiting immediately when a high-confidence match is found:

---

### Layer 1 — Exact Match

```
CleanText(input) == CleanText(correct name)
```

The fastest path. Case, spaces, and punctuation are ignored via `CleanText()`. Exits immediately on first hit.

```
"maharashtra"  →  Maharashtra   ✓ exits here
"MAHARASHTRA"  →  Maharashtra   ✓ exits here
" Maharashtra" →  Maharashtra   ✓ exits here
```

---

### Layer 2 — Wrong-Form Match

```
input is a substring of wrong-forms cell (Col E/F/G/H)
```

Checks the curated list of known alternate spellings. Only evaluated if the wrong-forms cell is not empty.

```
"Bombay"  →  Maharashtra   (BOMBAY found in Col E)
"Madras"  →  Tamil Nadu    (MADRAS found in Col E)
"U.P."    →  Uttar Pradesh (UP found after CleanText)
```

---

### Layer 3 — Short-Form Match

```
CleanText(input) == GetShortForm(correct name)
```

Derives initials dynamically from the correct name — no hardcoded abbreviation table needed.

```
"UP"  →  Uttar Pradesh    (U+P = UP)
"MP"  →  Madhya Pradesh   (M+P = MP)
"TN"  →  Tamil Nadu       (T+N = TN)
"HP"  →  Himachal Pradesh (H+P = HP)
```

---

### Layer 4 — Prefix / First-Word Match *(added in this version)*

```
Left(correctText, Len(input)) == input   [only if Len(input) ≤ Len(correct)]
```

Catches partial / truncated inputs efficiently without fuzzy scoring.

```
"MAHA"   →  Maharashtra
"UTTAR"  →  Uttar Pradesh
"RAJA"   →  Rajasthan
```

---

### Layer 5 — Contains Match *(added in this version)*

```
InStr(correctText, input) > 0  →  score = 0.71
```

If the input appears anywhere inside the correct name, assigns score `0.71`. Does not exit — continues to find the best containment match. Since `0.71 < 0.72` threshold, this result only wins if no better fuzzy score is found.

---

### Layer 6 — Fuzzy Score Match

```
MatchScore(input, correctText) → score stored if > bestScore
```

Runs the full position-weighted character matching algorithm across all rows. After the loop completes:

```
if bestScore >= 0.72  →  return bestMatch
else                  →  return originalText
```

```
"Uttarpradsh"  →  Uttar Pradesh  (~0.80)
"RAJSTAN"      →  Rajasthan      (~0.84)
"KERLA"        →  Kerala         (~0.78)
"XYZABC"       →  XYZABC         (score < 0.72, fallback)
```

---

## 📊 Examples

### State Matching (`Typ = "S"`)

```excel
=FUZZY(A2, "S")
```

| Input           | Output          | Layer         |
|-----------------|-----------------|---------------|
| `maharashtra`   | `Maharashtra`   | Exact         |
| `Bombay`        | `Maharashtra`   | Wrong-Form    |
| `UP`            | `Uttar Pradesh` | Short-Form    |
| `Maha`          | `Maharashtra`   | Prefix        |
| `Uttarpradsh`   | `Uttar Pradesh` | Fuzzy Score   |
| `KERLA`         | `Kerala`        | Fuzzy Score   |
| `Rajastan`      | `Rajasthan`     | Fuzzy Score   |
| `TN`            | `Tamil Nadu`    | Short-Form    |
| `XYZABC`        | `XYZABC`        | No match      |

### District Matching (`Typ = "D"`)

```excel
=FUZZY(B2, "D")
```

| Input         | Output        | Layer       |
|---------------|---------------|-------------|
| `mumbai`      | `Mumbai`      | Exact       |
| `Bangalor`    | `Bangalore`   | Fuzzy Score |
| `New Dlhi`    | `New Delhi`   | Fuzzy Score |
| `Hydrabad`    | `Hyderabad`   | Fuzzy Score |
| `Chenn`       | `Chennai`     | Prefix      |

### Edge Cases

```excel
=FUZZY("", "S")     → ""           ← empty input guard
=FUZZY(A2, "X")     → originalText ← invalid Typ, safe exit
=FUZZY(A2, "s")     → works fine   ← Typ is case-insensitive
```

### Bulk Correction (drag formula down)

```excel
' Column A: raw state names from a survey form
' Column B: corrected state names

B2 = =FUZZY(A2, "S")
B3 = =FUZZY(A3, "S")
...drag to B1000
```

---

## ⚙️ Helper Functions

The module contains three helper functions used internally by `FUZZY()`. They are also callable directly from VBA if needed.

### `CleanText(txt)`

Normalises raw text before comparison. Called on both input and MASTER data.

| Transformation      | Example                        |
|---------------------|--------------------------------|
| `UCase + Trim`      | `"Tamil Nadu"` → `"TAMILNADU"` |
| Remove spaces       | `"NEW DELHI"` → `"NEWDELHI"`   |
| Remove `&`          | `"J&K"` → `"JK"`              |
| Remove `. - _ / \` | `"U.P.-State"` → `"UPSTATE"`  |
| Remove `, ( )`      | `"Delhi(NCR)"` → `"DELHINCR"` |
| `@` → `A`           | `"M@harastra"` → `"MAHARASTRA"` |
| `0` → `O`           | `"Maharasht0a"` → `"MAHARASHTOA"` |
| `1` → `I`           | `"Del1i"` → `"DELII"`         |

### `GetShortForm(txt)`

Generates initials of a multi-word string.

```
"Uttar Pradesh"    →  "UP"
"Madhya Pradesh"   →  "MP"
"Tamil Nadu"       →  "TN"
"Himachal Pradesh" →  "HP"
```

Works by splitting on spaces and taking the first character of each word — no hardcoded abbreviation list needed.

### `MatchScore(s1, s2)`

Returns a similarity score between `0.0` and `1.0`. See [MatchScore Algorithm](#-matchscore-algorithm) for full details.

---

## 🧮 MatchScore Algorithm

`MatchScore` uses a **position-weighted character matching** approach with bonus scoring for common patterns.

### Scoring Weights

| Condition                        | Points      | Rationale                         |
|----------------------------------|-------------|-----------------------------------|
| Same character, same position    | `+3`        | Strong positional match           |
| Same character, within 2 places  | `+2`        | Near-position hit                 |
| Same character, far away         | `+1`        | Weak / anagram hit                |
| First 3 characters match (bonus) | `+5`        | Words sharing a prefix are likely the same |
| Input contained inside correct   | `+5`        | Containment is a strong signal    |
| Length difference penalty        | `-0.5 × diff` | Penalises strings of very different length |

### Normalisation

```
maxScore = Len(s2) × 3
MatchScore = score / maxScore
```

Score is clamped to `0` if negative.

### Acceptance Threshold

```
bestScore >= 0.72  →  match accepted
bestScore <  0.72  →  original text returned
```

The threshold is set at **0.72** to balance correction accuracy against false positives. Lowering it increases corrections but risks wrong matches; raising it makes the function more conservative.

### Score Examples

| Input         | Correct        | Score (approx) | Accepted? |
|---------------|----------------|----------------|-----------|
| `KERALA`      | `KERLA`        | ~0.78          | ✅ Yes     |
| `RAJSTAN`     | `RAJASTHAN`    | ~0.84          | ✅ Yes     |
| `MUMBAI`      | `MUMBAI`       | 1.00           | ✅ Yes     |
| `DLEHI`       | `DELHI`        | ~0.75          | ✅ Yes     |
| `XYZABC`      | `DELHI`        | ~0.03          | ❌ No      |

---

## ⚙️ Configuration

### Changing the Threshold

In `FUZZY.bas`, find the final result block and adjust the value:

```vba
If bestScore >= 0.72 Then     ' ← change this value
    FUZZY = bestMatch
Else
    FUZZY = originalText
End If
```

| Threshold | Effect |
|-----------|--------|
| `0.55`    | More corrections, higher false-positive risk |
| `0.72`    | Default — balanced accuracy |
| `0.85`    | Very conservative; only near-exact matches accepted |

### Adding New Data Types

To add a 5th type (e.g., `"Z"` for Zones using columns 9 and 10):

```vba
Case "Z"
    returnCol = 9
    wrongCol  = 10
```

Then add the corresponding columns to the MASTER sheet.

---

## 📁 File Structure

```
fuzzy-excel-udf/
│
├── README.md               ← This file
│
├── src/
│   └── FUZZY.bas           ← VBA module (import this into Excel)
│
├── docs/
│   ├── MASTER_SETUP.md     ← Detailed MASTER sheet guide
│   ├── HOW_IT_WORKS.md     ← Deep-dive into the matching engine
│   └── EXAMPLES.md         ← Extended examples and use cases
│
├── examples/
│   └── FUZZY_Demo.xlsm     ← Sample workbook with MASTER sheet pre-loaded
│
└── CHANGELOG.md            ← Version history
```

---

## 🤝 Contributing

Contributions are welcome! Here are some ways to help:

- **Add wrong-form data** — Submit PRs with more alternate spellings for the MASTER sheet
- **Report false matches** — Open an issue with the input, expected output, and actual output
- **Suggest new layers** — Jaro-Winkler, phonetic matching, etc.
- **Extend to other domains** — Company names, product names, country names

### How to Contribute

1. Fork the repository
2. Create a branch: `git checkout -b feature/your-feature-name`
3. Make your changes
4. Test with the demo workbook in `/examples`
5. Submit a Pull Request

---

## 📋 Changelog

See [CHANGELOG.md](./CHANGELOG.md) for full version history.

---

## 📄 License

This project is licensed under the **MIT License** — see [LICENSE](./LICENSE) for details.

You are free to use, modify, and distribute this code in personal or commercial projects.

---

## 🙋 FAQ

**Q: Does it work on Excel for Mac?**
A: Yes — VBA is supported on Excel for Mac. The function behaviour is identical.

**Q: Can I use it with non-geographic data?**
A: Yes. The MASTER sheet can contain any reference list — company names, product SKUs, department names, etc.

**Q: What happens if MASTER sheet is missing?**
A: Excel will throw a runtime error. Ensure the sheet is named exactly `MASTER` (all caps).

**Q: Can I change the column layout in MASTER?**
A: Yes — update the `returnCol` and `wrongCol` values in the `Select Case` block inside `FUZZY.bas`.

**Q: Why does FUZZY return the original text for a word I expect it to correct?**
A: The fuzzy score is below 0.72. Try adding the specific wrong-form to Column E or F of MASTER for a guaranteed match via Layer 2.

---

<div align="center">
  Made with ❤️ for cleaner data
</div>

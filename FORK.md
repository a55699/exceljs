# Fork Status

This is a **temporary fork** of [exceljs/exceljs](https://github.com/exceljs/exceljs) maintained by Protobi.

## Why This Fork Exists

Upstream exceljs has 100+ open PRs, some over a year old. We needed these fixes for production use, so we:
1. Adopted well-tested upstream PRs
2. Added critical pivot table features
3. Published to npm for community benefit
4. **Submitted all changes back to upstream**

** Goal:** Sunset this fork once upstream merges our contributions.

## Development Workflow

This repository uses [Claude Code](https://claude.com/claude-code) to accelerate development and maintenance:

- **Issue Triage & Analysis**: Review and investigate bug reports and feature requests
- **Pull Request Evaluation**: Assess upstream PRs for adoption into this fork
- **Test Development**: Create and maintain unit tests, integration tests, and end-to-end tests
- **Code Implementation**: Plan and execute bug fixes, feature additions, and refactoring
- **Release Management**: Version bumping, changelog updates, and npm publishing

This AI-assisted workflow enables rapid response to community issues while maintaining code quality through comprehensive test coverage.

---

## Fork Release History

### Unreleased

**Fixes**

- Defined names of a loaded file that are not only cells are kept: names defined by a formula or a constant were dropped (or written as a broken range, like `'SUM(A'!$A$4:$B$4`), and a name of one worksheet became a name of the workbook, so two sheets with their own `Rate` were written with one `Rate` and the formulas of one sheet used the cells of the other. Hidden names stay hidden
- The print area, print titles and names of a worksheet were given to the wrong worksheet when the file had a chartsheet before it

### 4.4.0-protobi.11.a55699.2 (2026-10-09)

Tagged `v4.4.0-protobi.11.a55699.2`; package [`@a55699/exceljs`](https://www.npmjs.com/package/@a55699/exceljs). Adds charts and fixes drawings with charts and images. Every change was made on its own branch with tests; the written files were checked in Excel 16.

**Charts**

- `worksheet.addChart()` adds bar (vertical or horizontal; clustered, stacked or percent stacked), line, area, pie, doughnut and scatter charts drawn from cell ranges, placed over a range like images; see README "Charts". Series come from `{sheet, ref}` references, formula strings or a data range; titles, axis titles, min/max, number format, gridlines and legend position can be set
- The file keeps the series values next to the references, as Excel does, taken from the cells when the workbook is written
- Charts are drawn in the worksheet's drawing with images and with the charts, shapes and pictures of a loaded file

**Fixes**

- Adding an image to a loaded workbook with a chart wrote a file Excel could not open (on the same worksheet or on another one). Charts, shapes and other drawing anchors of a loaded file are now kept one by one, so images and new charts can be added next to them
- A loaded workbook with a text box drawn over a chart (chart shapes) could not be opened after it was written again
- The same image added to two worksheets in a row lost its relationship on the second worksheet (protobi/exceljs#24)

### 4.4.0-protobi.11.a55699.1 (2026-10-08)

First release of the [a55699/exceljs](https://github.com/a55699/exceljs) fork, tagged `v4.4.0-protobi.11.a55699.1`, with the package renamed to `@a55699/exceljs`. It is 4.4.0-protobi.11, which Protobi has not released yet, plus the changes below. Every change was made on its own branch with tests; changes to the written XML were checked by opening the files in Excel 16.

**Package**

- Renamed from `@protobi/exceljs` to `@a55699/exceljs` (Protobi publishes `@protobi/exceljs`); `repository` points to a55699/exceljs
- `prepublishOnly` runs `npm run build`, so the published package includes `dist/` (it is not in git); `publishConfig.access` is `public`
- README: installation and import examples use `@a55699/exceljs`; the migration note says that imports change when switching back to `exceljs`

**Included from 4.4.0-protobi.11**

- **Number format limit guard** (Protobi [#83](https://github.com/protobi/exceljs/issues/83))
  - New workbook option `numFmtLimit` (default 206) and `Workbook#addNumberFormat(formatCode)`, which returns the numFmtId and throws `NumberFormatLimitError` once the workbook holds `numFmtLimit` custom formats
  - Opt-in: setting `cell.numFmt` directly does not go through this check

**Pivot tables**

- Escape `&`, `<`, `>` and quotes in field names and text values; reject empty or duplicate (case-insensitive) header names, which Excel cannot open
- Write cell values as Excel does: formulas by their result, rich text and hyperlinks by their text, dates as date items, numbers and booleans by type, with the `sharedItems` attributes Excel expects for mixed columns
- Write correct record, item and field counts, including empty columns and page fields without items
- Much faster on large sources (200,000 rows: from not finishing to about 5 seconds)
- Adding a pivot table to a loaded workbook that already has pivot tables: number the new parts and cache ids after the loaded ones and keep the loaded relationships
- Rebuild the pivot cache from the source sheet when writing, so data changed after `addPivotTable()` is used instead of failing

**Form controls and notes**

- Form checkboxes are read when a workbook is loaded (`getFormCheckboxes()`), so a load + write keeps them
- Other form controls (buttons, drop-downs, list boxes, spinners, option buttons, group boxes...) are kept as read when a workbook is loaded and written
- A checkbox linked to an empty cell writes its checked state to that cell, since Excel shows the state of the linked cell; setting `checked` also sets the linked cell. Cells that already hold a value are not changed
- Notes and checkboxes on the same worksheet get different shape ids
- Notes on cells without a value are loaded (they were dropped)
- Notes loaded without margins are written with the default margins (were `NaNmm`)

**Other fixes**

- `deepMerge` ignores `__proto__`, `constructor` and `prototype` keys (prototype pollution through crafted note data)
- `spliceRows` moves merged cells and data validations with their rows

**TypeScript, tests and tooling**

- Typings for pivot tables, form checkboxes, `numFmtLimit`, `addNumberFormat` and `NumberFormatLimitError`
- The TypeScript specs run in CI; the failing stream spec is fixed
- CI runs on Node 22, 24 and 26 with `npm ci`; actions pinned by commit; `engines` set to `>=14.14.0`
- prettier and eslint no longer disagree in the pre-commit hook; `npm ci` works without `--legacy-peer-deps`
- Lockfile refreshed for `tmp`, `minimatch` and `brace-expansion`
- `.gitattributes` normalizes line endings to LF; `test-output-dir/` is no longer tracked
- New tests for the pivot count metric (#2885), pivot grouping, multiple pivots and page fields, and form checkbox output

**Known limitations**

- `spliceColumns` still does not move merged cells
- A note whose text is a single run without a font loads as a plain string, so custom margins on it are not kept
- Button macros are not stored in .xlsx files, so a kept button's macro reference does not resolve

### 4.4.0-protobi.10 (2026-05-06)

**Bug Fix: richText shared-string deduplication in streaming writer** ([#66](https://github.com/protobi/exceljs/pull/66), cherry-picked from [#50](https://github.com/protobi/exceljs/pull/50) by @gwkline)

When using `WorkbookWriter` with `useSharedStrings` enabled, every cell containing richText collapsed into a single shared-string entry, corrupting the strings table. Root cause: `SharedStrings.add()` used the value directly as a hash key; richText objects coerced to the string `"[object Object]"`, deduping all of them into one entry.

**Fix:** Hash richText values by their rendered XML representation (via `SharedStringXform.toXml()`). Plain strings unchanged.

**Tests:** Adds two regression tests (`spec/unit/utils/shared-strings.spec.js`):
1. Equal richText values share one entry; different ones get separate entries
2. richText values that differ only in formatting are NOT deduped

**Upstream context:** This bug was first reported as [exceljs/exceljs#2267](https://github.com/exceljs/exceljs/issues/2267) (May 2023). An alternative fix was opened upstream as [exceljs/exceljs#2588](https://github.com/exceljs/exceljs/pull/2588) (Nov 2023) but has been stale since Feb 2024.

**Build Fix: Pin uuid to ^9.0.1** ([#64](https://github.com/protobi/exceljs/pull/64))

`npm install` had drifted to `uuid@14.0.0`, which broke `npm run build`:
- uuid@14 dropped the `main` field in favor of `exports`-only — browserify@16 doesn't honor `exports`
- uuid@11 (npm's recommended CJS target) ships ES2021 syntax (`??=`) browserify@16 can't parse
- uuid@9.0.1 is the highest version that keeps both `main` and pre-ES2021 CJS output

**Documentation: AGENTS.md for AI-generated PRs** ([#65](https://github.com/protobi/exceljs/pull/65))

Adds `AGENTS.md` at repo root with hard rules for AI coding agents (Claude Code, Cursor, Codex, etc.): surgical scope, no formatter sweeps, complete PR descriptions, real-fixture testing, cross-check open PRs, Excel-verification for serialization changes. Auto-discovered by AI tools; referenced from README.md, CONTRIBUTING.md, and the PR template.

**Tests:** All 886 unit tests passing (4 in `SharedStrings`, up from 2).

---

### 4.4.0-protobi.9 (2026-02-01)

**New Feature: Pivot Table & Chart Round-Trip Preservation** ([#41](https://github.com/cjnoname/excelts/issues/41))

Implements complete round-trip preservation for Excel files containing pivot tables, charts, and drawings. Previously, reading an Excel file with pivot tables/charts and writing it back would result in Excel corruption warnings and loss of pivot table/chart data.

**Hybrid Preservation Approach:**
- Stores raw XML for pivot tables, charts, and drawings
- Extracts minimal metadata (cacheId, relationships) for structural integrity
- Doesn't attempt to fully model complex Excel structures
- Enables round-trip without data loss or corruption

**What's Preserved:**

1. **Pivot Tables:**
   - Pivot table definitions (`xl/pivotTables/pivotTable*.xml`)
   - Cache definitions (`xl/pivotCache/pivotCacheDefinition*.xml`)
   - Cache records (`xl/pivotCache/pivotCacheRecords*.xml`)
   - Correct cacheId-to-filename mapping via relationship traversal

2. **Charts:**
   - Chart definitions (`xl/charts/chart*.xml`)
   - Chart styles (`xl/charts/style*.xml`)
   - Chart colors (`xl/charts/colors*.xml`)
   - All chart relationships

3. **Drawings:**
   - Detects drawings with chart references
   - Preserves as raw XML instead of parsing
   - Regular drawings without charts still parsed normally (for image support)
   - Recreates drawing relationships on write

4. **Row Heights:**
   - Preserves `x14ac:dyDescent` attribute for accurate row height
   - Only outputs `customHeight` if present in original file
   - Fixes row height discrepancies during round-trip

**Implementation Details:**
- 16 files modified, 593 insertions, 40 deletions
- Added `test/test-roundtrip-pivot.js` for verification
- Updated test expectations for dyDescent preservation
- All required content type declarations added
- Relationship mappings preserved correctly

**Testing:**
- ✅ 1091/1091 tests passing (unit + integration + end-to-end)
- ✅ Round-trip test with real-world pivot table/chart files
- ✅ Excel opens files without corruption warnings
- ✅ Pivot tables remain functional after round-trip
- ✅ Charts display correctly after round-trip

**Known Limitations:**
- Write-only pivot table creation still supported (existing functionality)
- Reading/modifying existing pivot table definitions not yet implemented
- This implementation enables preservation, not programmatic access
- Hybrid approach: pivot tables can be created OR preserved, not both simultaneously

---

**Bug Fix: Cell Comment/Note Protection Properties** (2026-02-02)

Fixes a critical bug where comment protection properties (`locked`, `lockText`) and positioning (`editAs`) were not being parsed correctly during file reading, causing them to return `null` instead of their actual values.

**Root Cause:**
- SAX XML parser strips `x:` namespace prefix (Excel namespace) but preserves `v:` prefix (VML namespace)
- VML xform classes were using prefixed tag names (`x:Locked`, `x:ClientData`) in their maps
- Parser delivered unprefixed tag names (`Locked`, `ClientData`), causing lookup failures

**The Fix:**
- Updated 5 VML xform classes to use unprefixed tag names for parsing
- Render methods still write prefixed XML for Excel compatibility
- Added comments documenting the SAX parser behavior

**Files Modified:**
- `lib/xlsx/xform/comment/vml-client-data-xform.js`
- `lib/xlsx/xform/comment/vml-shape-xform.js`
- `lib/xlsx/xform/comment/vml-anchor-xform.js`
- `lib/xlsx/xform/comment/style/vml-protection-xform.js`
- `lib/xlsx/xform/comment/style/vml-position-xform.js`

**Testing:**
- ✅ 1091/1091 tests passing (vs 1089 passing before fix)
- ✅ Fixed 2 previously failing integration tests:
  - "writes notes" - comment protection properties now round-trip correctly
  - "Cell annotation supports setting margins and protection properties"
- ✅ Verified this bug also exists in upstream ExcelJS (not fixed there yet)

**Impact:**
- Comment protection properties (`note.protection.locked`, `note.protection.lockText`) now parse correctly
- Comment margins (`note.margins`) now preserve correctly during round-trip
- Comment positioning (`note.editAs`) now reads back correctly
- Excel files with protected comments can now be read and written without losing properties

**Related Issues:**
- Inspired by [ExcelTS Issue #41](https://github.com/cjnoname/excelts/issues/41)
- Addresses Excel corruption: "Removed Part: /xl/pivotTables/pivotTable1.xml"
- Fixes missing content type declarations
- Resolves row height preservation issues

**Commit:** 853aa94

---

### 4.4.0-protobi.8 (2026-01-31)

**New Features**

- **Form Control Checkbox support** ([#31](https://github.com/protobi/exceljs/issues/31))
  - Implement legacy Form Control checkboxes compatible with Excel 2007+, WPS Office, and LibreOffice
  - Add `addFormCheckbox()` and `getFormCheckboxes()` worksheet API
  - Support checked/unchecked state, linked cells, and custom text labels
  - Write-only support (reading not yet implemented)
  - Example: `ws.addFormCheckbox('B2:D3', { checked: true, link: 'E2', text: 'Accept terms' })`

- **Page fields support for pivot tables** ([#3021](https://github.com/protobi/exceljs/issues/3021))
  - Add `pages` and `pageDefaults` parameters to pivot table API
  - Enables report filtering in Excel pivot tables
  - Example: `ws.addPivotTable({ rows: ['Region'], pages: ['Year'], pageDefaults: { Year: 2025 } })`

**Security & Dependency Updates**

- **Upgrade archiver to 7.x** ([#11](https://github.com/protobi/exceljs/issues/11))
  - Implement Duplex interface in StreamBuf (`_read()` and `_write()` methods)
  - Upgrade archiver: ^5.3.2 → ^7.0.1
  - Removes deprecated glob 7.x from production dependencies (now uses glob 10.5.0)
  - Addresses CVE-2025-64756 (glob command injection vulnerability)
  - Addresses CVE-772 (inflight vulnerability via glob→inflight chain)
  - glob 7.x remains only in devDependencies (acceptable)

**Testing:**
- ✅ 884/884 unit tests passing
- ✅ Integration tests passing
- ✅ End-to-end tests passing
- ✅ Checkbox functionality verified with archiver 7.x

**Pull Requests:** Merged to master via issue-31 and issue-11 branches

---

### 4.4.0-protobi.7 (2026-01-26)

**Bug Fixes**

- Added support for reading Excel files created by HAN CELL (Korean spreadsheet software) ([exceljs#3014](https://github.com/exceljs/exceljs/issues/3014))
  - Fixed crash when reading files with `x:` namespace prefix
  - Added null checks for workbook, appProperties, and coreProperties
  - Made parsers lenient with unknown XML tags
- Fixed data bar conditional formatting crash when using minimal API options ([exceljs#3015](https://github.com/exceljs/exceljs/issues/3015))
  - Provide default cfvo array: `[{ type: "min" }, { type: "max" }]`
  - Provide default color: Excel blue `#638EC6`

**Testing:**
- Added integration test with real HAN CELL file
- Added unit tests for data bar default values
- All 198 tests passing (2 pre-existing unrelated failures)

**Pull Requests:** [#33](https://github.com/protobi/exceljs/pull/33), [#34](https://github.com/protobi/exceljs/pull/34), [#35](https://github.com/protobi/exceljs/pull/35), [#36](https://github.com/protobi/exceljs/pull/36)

---

### 4.4.0-protobi.6 (2025-12-14)

**Security Fix**

- Fixed Content Security Policy (CSP) violation in Vite builds ([#28](https://github.com/protobi/exceljs/issues/28))
- Added npm overrides to force `asn1.js@5.4.1` (removes eval usage)
- Dependency chain: `browserify` → `crypto-browserify` → `browserify-sign` → `parse-asn1` → `asn1.js`
- Verified zero `eval()` calls in all browser bundles
- Browser builds now CSP-compliant for modern build tools (Vite, Webpack 5+)

**Testing:**
- All existing tests passing
- Verified dist bundles contain no eval/runInThisContext calls

**Commits:** 2f6f8b6, 8d106d5

---

### 4.4.0-protobi.5 (2025-12-06)

**Security & Dependency Updates**

**Production Dependencies:**
- `archiver`: ^5.0.0 → ^5.3.2
- `unzipper`: listed as 0.10.11 → 0.12.3, but this upgrade is not in the code: package.json has `^0.10.11` (0.10.14 locked)

**Dev Dependencies (Major Updates):**
- `mocha`: ^7.2.0 → ^11.7.5 (fixes ReDoS in debug, js-yaml, minimatch)
- `chai-xml`: ^0.3.2 → ^0.4.1 (fixes xml2js prototype pollution)
- `got`: ^9.0.0 → ^11.8.6 (downgraded from 14 for test compatibility)
- `eslint`: ^6.5.1 → ^8.57.1 (previously listed as ^9.39.1; package.json has ^8.57.1)
- `grunt-contrib-jasmine`: ^2.2.0 → ^4.0.0
- `prettier-eslint`: ^11.0.0 → ^16.4.2
- `prettier-eslint-cli`: ^5.0.0 → ^8.0.1

**Security Improvements:**
- Reduced npm audit vulnerabilities from 38 → 15 
- All remaining vulnerabilities are in dev/test dependencies only
- No production code security issues

**Testing:**
- ✅ Unit tests: 883 passing, 1 pending
- ✅ Integration tests: 198 passing
- ✅ End-to-end tests: 1 passing
- ✅ Build: Successful
- ⚠️ Browser tests: Disabled (puppeteer@19 compatibility issues)

**Known Limitations:**
- `glob@7.x` remains in dependency tree (deprecated but no active CVEs)
  - Only affects glob 10.x-11.x CLI (CVE-2025-64756), not 7.x library API
  - Upgrading to archiver 7.x requires StreamBuf refactoring (_read/_write methods)
- Browser tests disabled via `.disable-test-browser` (puppeteer hangs)
- `got` kept at v11 instead of v14 due to breaking API changes in end-to-end tests

**Related Issues:**
- [#11](https://github.com/protobi/exceljs/issues/11) - npm audit vulnerabilities (protobi fork)
- [#2984](https://github.com/exceljs/exceljs/issues/2984) - Security vulnerabilities in transitive dependencies (upstream)
- [#3006](https://github.com/exceljs/exceljs/issues/3006) - glob 7.x dependency discussion (upstream)

**Commits:** f8e3e47, 9c0f259, 0d57c7b, 639922f

---

### 4.4.0-protobi.4 (2025-11)

Multiple pivot tables support, pivot table count metric, streaming limitations documented.

### 4.4.0-protobi.2 (2025-11-07)

**Added:**
- Pivot table count metric (upstream PR #2885)
- 5 bug fixes from upstream PRs

**Fixed:**
- Boolean XML attribute parsing (#2851)
- ExcelToDate validation (#2956)
- DynamicFilter parsing (#2973)
- WorkbookReader sharedString resolution (#2915)
- Autofilter undefined guard (#2978)

**Security:**
- Added package-lock.json
- Ran npm audit fix (reduced vulnerabilities)

### 4.4.0-protobi.1 (2025-11-06)

**Initial fork release:**
- Multiple pivot tables from same source
- XML special character escaping fixes
- Column width control for pivot tables

---

## Status Tracking

### ✅ Features Already Submitted to Upstream (Waiting for Merge)

| Feature/Fix | Our Issue | Upstream PR | Status | Date |
|-------------|-----------|-------------|--------|------|
| Pivot table count metric | [#12](https://github.com/protobi/exceljs/issues/12) | [#2885](https://github.com/exceljs/exceljs/pull/2885) | ⏳ Open | Feb 2025 |
| Boolean XML parsing fix | [#18](https://github.com/protobi/exceljs/issues/18) | [#2851](https://github.com/exceljs/exceljs/pull/2851) | ⏳ Open | Nov 2024 |
| ExcelToDate validation | [#19](https://github.com/protobi/exceljs/issues/19) | [#2956](https://github.com/exceljs/exceljs/pull/2956) | ⏳ Open | Aug 2025 |
| DynamicFilter parsing | [#20](https://github.com/protobi/exceljs/issues/20) | [#2973](https://github.com/exceljs/exceljs/pull/2973) | ⏳ Open | Sep 2025 |
| SharedString fix | [#21](https://github.com/protobi/exceljs/issues/21) | [#2915](https://github.com/exceljs/exceljs/pull/2915) | ⏳ Open | Apr 2025 |
| Autofilter undefined guard | [#22](https://github.com/protobi/exceljs/issues/22) | [#2978](https://github.com/exceljs/exceljs/pull/2978) | ⏳ Open (partial) | Sep 2025 |
| Image reuse fix | [#24](https://github.com/protobi/exceljs/issues/24) | [#2876](https://github.com/exceljs/exceljs/pull/2876) | ⏳ Open | Feb 2024 |
| Conditional formatting + hyperlinks corruption fix | [#25](https://github.com/protobi/exceljs/issues/25) | [#2803](https://github.com/exceljs/exceljs/pull/2803) | ⏳ Open | Sep 2024 |
| Conditional formatting stopIfTrue & operators | [#26](https://github.com/protobi/exceljs/issues/26) | [#2736](https://github.com/exceljs/exceljs/pull/2736) | ⏳ Open | Jul 2024 |

**Total:** 9 PRs adopted from upstream, waiting for official merge

### 📝 Fork-Specific Features (Submitted to Upstream)

| Feature/Fix | Our Issue | Upstream PR | Status | Date |
|-------------|-----------|-------------|--------|------|
| Multiple pivot tables support | [#5](https://github.com/protobi/exceljs/issues/5) | [#2995](https://github.com/exceljs/exceljs/pull/2995) | ⏳ Open | Nov 2025 |
| XML special character escaping | [#3](https://github.com/protobi/exceljs/issues/3) | [#2996](https://github.com/exceljs/exceljs/pull/2996) | ⏳ Open | Nov 2025 |
| Pivot table column width control | [#8](https://github.com/protobi/exceljs/issues/8) | [#2997](https://github.com/exceljs/exceljs/pull/2997) | ⏳ Open | Nov 2025 |
| HAN CELL file support | - | [#3020](https://github.com/exceljs/exceljs/pull/3020) (replaces closed #3017) | ⏳ Open | Jan 2026 |
| Data bar conditional formatting defaults | - | [#3019](https://github.com/exceljs/exceljs/pull/3019) (replaces closed #3018) | ⏳ Open | Jan 2026 |

**Status:** All original contributions submitted, waiting for upstream review

###  Community Fork Contributions (Adopted & Submitted)

| Feature/Fix | Our Issue | Source | Upstream PR | Status | Date |
|-------------|-----------|--------|-------------|--------|------|
| Table addRow() fix | [#23](https://github.com/protobi/exceljs/issues/23) | [rmartin93/exceljs-fork](https://github.com/rmartin93/exceljs-fork) | [#2998](https://github.com/exceljs/exceljs/pull/2998) | ⏳ Open | Nov 2025 |

**Status:** Adopted from community fork, submitted to upstream

###  Security & Maintenance

| Feature/Fix | Our Issue | Upstream PR | Status |
|-------------|-----------|-------------|--------|
| Add package-lock.json | [#10](https://github.com/protobi/exceljs/issues/10) | - | Fork-specific |
| Run npm audit fix | [#11](https://github.com/protobi/exceljs/issues/11) | - | Fork-specific |

**Status:** Security improvements for our fork deployment

---

## Installation

### This Fork (Current)

```bash
npm install @protobi/exceljs
```

### Official Package (Future)

Once upstream merges our changes:

```bash
npm install exceljs
```

---

## Migration Plan

### When to Switch Back to Official

Monitor these conditions:
1. Upstream releases version with our features
2. All critical PRs merged (#2885, #2915, etc.)
3. Our unique features submitted and merged

### How to Switch Back

```bash
# 1. Check if official version has our features
npm view exceljs version
# Compare against our version: 4.4.0-protobi.2

# 2. Review official changelog
npm view exceljs --json | jq .versions

# 3. Switch packages
npm uninstall @protobi/exceljs
npm install exceljs

# 4. No code changes needed - API compatible!
```

### Deprecation Process

When upstream catches up, we will:

1. **Add deprecation notice** to npm package
   ```bash
   npm deprecate @protobi/exceljs "Use official 'exceljs' package - our changes have been merged upstream"
   ```

2. **Update README** with migration instructions

3. **Archive repository** on GitHub

4. **Final release** pointing to official package

---

## Features Exclusive to This Fork

### 1. Multiple Pivot Tables from Same Source

**Why it matters:** Upstream only supports one pivot table per workbook. We support multiple pivot tables from the same source data with unique cache IDs.

**Code:**
```javascript
// ✅ Works in @protobi/exceljs
// ❌ Crashes in official exceljs

const worksheet1 = workbook.addWorksheet('Data');
worksheet1.addRows([/* data */]);

const worksheet2 = workbook.addWorksheet('Pivot1');
worksheet2.addPivotTable({ sourceSheet: worksheet1, /* ... */ });

const worksheet3 = workbook.addWorksheet('Pivot2');
worksheet3.addPivotTable({ sourceSheet: worksheet1, /* ... */ }); // Works!
```

**Files changed:**
- `lib/doc/pivot-table.js` - Unique cache IDs per pivot table
- `lib/xlsx/xform/book/workbook-xform.js` - Support multiple cache definitions

**Upstream status:** Submitted as [#2995](https://github.com/exceljs/exceljs/pull/2995) (open)

### 2. Pivot Table Count Metric

**Why it matters:** Count is a fundamental Excel pivot table aggregation. Upstream only supports sum.

**Code:**
```javascript
// ✅ Works in @protobi/exceljs
// ❌ Not available in official exceljs (yet)

worksheet.addPivotTable({
  sourceSheet: worksheet1,
  rows: ['Category'],
  columns: ['Region'],
  values: ['Sales'],
  metric: 'count', //  NEW!
});
```

**Files changed:**
- `lib/doc/pivot-table.js` - Accept count metric
- `lib/xlsx/xform/pivot-table/pivot-table-xform.js` - Generate XML with subtotal="count"

**Upstream status:** PR exists ([#2885](https://github.com/exceljs/exceljs/pull/2885)), adopted here

### 3. Table addRow() Fix for Templates

**Why it matters:** Loading Excel files with tables and adding rows to them crashes in official exceljs. This is critical for template-based workflows.

**Code:**
```javascript
// ✅ Works in @protobi/exceljs
// ❌ Crashes in official exceljs with "Cannot read properties of undefined (reading 'length')"

const workbook = new Excel.Workbook();
await workbook.xlsx.readFile('template.xlsx');

const worksheet = workbook.getWorksheet('Data');
const table = worksheet.getTable('MyTable');

// Add rows to the loaded table
table.addRow(['New', 'Data', 'Here']); // Works!
table.addRow(['More', 'Data', 'Here']); // Works!

await workbook.xlsx.writeFile('output.xlsx');
```

**What it fixes:**
- "Cannot read properties of undefined (reading 'length')" error
- Missing worksheet references in loaded tables
- Table references not expanding dynamically when rows are added
- Excel filter buttons disappearing after save

**Files changed:**
- `lib/doc/table.js` - Dynamic table reference updates, autoFilterRef handling
- `lib/doc/worksheet.js` - Table loading compatibility fixes

**Upstream status:** Adopted from [rmartin93/exceljs-fork](https://github.com/rmartin93/exceljs-fork), submitted as [#2998](https://github.com/exceljs/exceljs/pull/2998) (open)

### 4. Enhanced Bug Fixes

See "Status Tracking" section above for 6 bug fixes adopted from upstream PRs.

---

## API Compatibility

This fork maintains **100% API compatibility** with official exceljs.

**What this means:**
- ✅ Drop-in replacement: `require('exceljs')` works identically
- ✅ All official features work exactly the same
- ✅ Additional features are opt-in (won't break existing code)
- ✅ Switching back to official requires zero code changes

**Version mapping:**
```
Official exceljs:     4.4.0
This fork:            4.4.0-protobi.2
                      └─┬─┘ └──┬───┘
                        │      └─── Fork version (increments with our changes)
                        └────────── Matches upstream base version
```

---

## Contributing

### For Bug Fixes & Features

Please contribute to **upstream exceljs**, not this fork!

- Upstream repo: https://github.com/exceljs/exceljs
- Upstream issues: https://github.com/exceljs/exceljs/issues
- Upstream discussions: https://github.com/exceljs/exceljs/discussions

### For Fork-Specific Issues

Only use our repo for:
- Issues with our specific features (#5, #8)
- Questions about migration
- Fork maintenance

Create issue: https://github.com/protobi/exceljs/issues

---

## Testing

This fork passes all upstream tests plus additional tests for our features.

```bash
# Run all tests
npm test

# Run pivot table tests specifically
npm run test:integration -- --grep "Pivot Tables"
```

### Test Updates vs Upstream

**New Test Files Added:**
- `spec/integration/workbook/pivot-tables-with-count.spec.js` (78 lines)
  - Tests pivot table with `metric: 'count'` feature
  - Validates XML generation for count aggregation

- `spec/integration/issues/issue-1804-add-image.spec.js` (53 lines)
  - Tests image reuse fix (PR #2876)
  - Validates same image added multiple times maintains correct references

**Test Infrastructure Changes:**
- Browser tests disabled (puppeteer compatibility issues with updated dependencies)
- Dev test dependencies updated: mocha 7→11, chai-xml 0.3→0.4
- End-to-end tests: got API updated (v9→v11)

**Test Results:**
- Unit tests: 883 passing, 1 pending (unchanged from upstream)
- Integration tests: 198 passing (includes 2 new tests above)
- End-to-end tests: 1 passing
- Browser tests: Skipped via `.disable-test-browser`

**Note:** Fork maintains all upstream tests + adds specific tests for adopted PR features. No upstream tests were removed or modified.

---

## Monitoring Upstream

We actively monitor upstream for:
1. **Our PRs being merged** - Track at https://github.com/exceljs/exceljs/pulls
2. **New releases** - Watch https://github.com/exceljs/exceljs/releases
3. **Breaking changes** - Review changelogs for compatibility

**Current watch list:**
- 9 adopted PRs awaiting merge (#2851, #2876, #2915, #2956, #2973, #2978, #2885, #2803, #2736)
- 3 original PRs awaiting merge (#2995, #2996, #2997)
- 1 community fork contribution awaiting merge (#2998)
- **Total: 13 PRs** watching upstream

**Update frequency:** Monthly check for upstream progress

---

## Support & Contact

- **Bug reports:** [GitHub Issues](https://github.com/protobi/exceljs/issues)
- **Questions:** [GitHub Discussions](https://github.com/protobi/exceljs/discussions)

For official exceljs support:
- Upstream discussions: https://github.com/exceljs/exceljs/discussions

---

## License

This fork maintains the same license as upstream exceljs: **MIT**

See [LICENSE](LICENSE)

---

**Last Updated:** 2026-10-08
**Watching:** 15 upstream PRs awaiting review/merge
**Status:** Active maintenance until upstream merge - Continuing to adopt community contributions!

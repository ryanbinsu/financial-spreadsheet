# Financial spreadsheet

Builds a GBP financial tracker with exactly three connected worksheets: **Overview**, **Income**, and **Outgoings**. The workbook includes monthly summaries, expandable transaction tables, category dropdowns, and a separate debt table.

## Requirements

- Windows with PowerShell and desktop Microsoft Excel installed. The verification scripts automate Excel through COM and need a signed-in desktop session.
- Node.js with the `@oai/artifact-tool` package available. This project currently uses the package and Node executable bundled with the Codex primary runtime; it does not include a `package.json` or dependency installer.

Run all commands below from the repository root. Close any open copies of the generated workbooks first.

**Rebuilding overwrites the generated workbooks and resets transactions to blank. Back up any workbook containing your own financial records before running these scripts.**

## Set up the local runtime

The following PowerShell commands use the runtime location on the machine where this project was created. Adjust `$runtime` if your installation differs.

```powershell
$runtime = Join-Path $env:USERPROFILE '.cache/codex-runtimes/codex-primary-runtime/dependencies/node'
$node = Join-Path $runtime 'bin/node.exe'
$packages = Join-Path $runtime 'node_modules'

if (-not (Test-Path -LiteralPath $node)) {
    throw "Node executable not found: $node"
}
if (-not (Test-Path -LiteralPath (Join-Path $packages '@oai/artifact-tool'))) {
    throw "The Codex artifact-tool dependency is missing: $packages"
}
if (-not (Test-Path -LiteralPath './node_modules')) {
    New-Item -ItemType Junction -Path './node_modules' -Target $packages
}
```

The junction lets `build.mjs` resolve the bundled package without copying it into the repository. If `node_modules` already exists, ensure it provides `@oai/artifact-tool` from the intended runtime. With a different supported Node installation, set `$node` to that executable instead.

## Build and verify

Run these steps in order, stopping if any step fails.

1. Generate the initial workbook and previews:

   ```powershell
   & $node ./build.mjs
   if ($LASTEXITCODE -ne 0) { throw 'Workbook generation failed.' }
   ```

2. Repair the initial export in Excel, apply workbook settings, and verify calculations:

   ```powershell
   & ./verify.ps1
   ```

   This opens Excel in the background and tests added table rows, category edits, transfer exclusions, debt separation, month/year changes, and opening/closing balance carry-forward. It removes its temporary test data, checks zero summaries, and saves `outputs/tracker/Financial Tracker Ready.xlsx`.

3. Reopen the verified workbook, check native Excel features, and save the final file:

   ```powershell
   & ./finish.ps1
   ```

   This checks the worksheet count, category dropdowns, filters, frozen headers, and debt table. It also finishes navigation links and print settings, then saves the final deliverable:

   **`outputs/tracker/Financial Tracker.xlsx`**

All three steps are required. The JavaScript calculation engine has limitations with this workbook's structured references and hyperlinks, so its `Behaviour checks passed` message alone does not establish correctness. `verify.ps1` performs the calculation assertions in Excel; `finish.ps1` produces the final file. Use the final workbook only after both PowerShell scripts complete successfully.

## Generated files

| File | Purpose |
| --- | --- |
| `outputs/tracker/Financial Tracker.xlsx` | Final workbook after all three steps complete |
| `outputs/tracker/Financial Tracker Ready.xlsx` | Intermediate workbook saved by Excel verification |
| `outputs/tracker/Overview.png`, `Income.png`, `Outgoings.png` | Initial builder previews; these precede Excel repairs and final formatting |
| `outputs/tracker/*.inspect.ndjson` | Diagnostic inspection output, when emitted by the exporter |

## Using the workbook

1. On **Overview**, set the starting month in `A31` to the first day of your first tracked month and enter its opening cash balance in `A33`.
2. Select a reporting month in `B4`, on or after the starting month.
3. Enter positive amounts and real dates in **Income** and **Outgoings**. Choose a category for each transaction; press Tab in the last table cell to add a row.
4. Use **Transfer between own accounts** for internal transfers. These entries are excluded from cash-flow totals.
5. Record repayments as **Debt Repayments** and update outstanding debt balances manually in the separate table on **Outgoings**. Outstanding balances are current snapshots and are not deducted again from monthly cash flow.

Editable inputs are highlighted in amber. Monthly summaries update automatically, and opening balances carry forward from earlier recorded cash flows.

## Troubleshooting

- **Cannot resolve `@oai/artifact-tool`:** Check the runtime path and the `node_modules` junction.
- **PowerShell blocks a script:** Follow your organisation's policy for allowing local scripts before running `verify.ps1` or `finish.ps1`.
- **Excel COM reports a logon/session or access error:** Run the PowerShell steps from a signed-in Windows desktop session with Excel available. A restricted automation environment may require permission to run Excel outside its sandbox.
- **Excel cannot open or save a workbook:** Close open copies and check that `outputs/tracker` is writable. `verify.ps1` intentionally opens the initial export in repair mode and saves a clean intermediate file.
- **Initial previews show formula or hyperlink errors:** Complete the Excel steps. The previews are generated before Excel repairs and recalculates the workbook; they do not represent the final saved file.

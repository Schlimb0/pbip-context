# pbip-context

A single `.cmd` file that turns a Power BI Project (PBIP / PBIR / TMDL) folder into one dense, logic-only text file, built to be **screenshotted** and sent to an AI assistant alongside screenshots of the report itself.

## Usage

1. Download `PBIP_Context.cmd`.
2. Put it in your Power BI project folder (the one containing `YourReport.pbip`).
3. Double-click it (or run `PBIP_Context.cmd /nopause` from a terminal).
4. Open the generated `PBIP_Context.txt` with **Word Wrap on**, a small font and a maximised window, and screenshot it.

## What it produces

Only what a report screenshot cannot show:

- **Tables**: columns with data types on one line, calculated columns, measures (full DAX), calculation items, hierarchies, and the Power Query (M) or calculated-table source.
- **Relationships**: `Sales[DateKey] *:1 Date[Date] (inactive)`.
- **Shared queries / parameters, DAX functions and row-level security roles.**
- **Report**: per page, each data visual as `type "title": role: fields; sort; filter; formatting driven by`, plus report and page filters and report-level measures.

Left out on purpose: layout, position, colours, fonts and all other formatting (the screenshots show those), visuals with no data bound (text boxes, shapes, images; only counted), lineage tags, annotations, format strings, display folders, cultures, auto date/time tables, and raw file dumps.

## Built for screenshots

- Every item is exactly one line and there are no blank lines or banners. Multi-line DAX and M are collapsed onto one line (string literals are untouched, `//` and `--` comments become `/* */`).
- Short items share a line, separated by `  |  `, up to `WIDTH` characters (default 200). Set `WIDTH` at the top of the script to roughly what fits across your window.
- Long items are not hard-wrapped, so your editor's word wrap fills the full window width whatever its size.

**Data is summarised, not dumped.** "Enter Data" tables are decoded and described per column inline: columns with few distinct values (default 15 or fewer) list each value with its count; others show the distinct count, numeric range and a few samples. Change `MAXD` at the top of the script to adjust.

## Notes

- A single .cmd (it calls the built-in Windows PowerShell internally); no installs needed. Works on Windows 10/11.
- Reads TMDL or `model.bim` semantic models, and PBIR or legacy `report.json` reports.
- Review the output before sharing: it may contain server names, queries or other details from your model.

## License

MIT

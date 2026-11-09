# pbip-context

A single `.cmd` file that turns a Power BI Project (PBIP / PBIR / TMDL) folder into one plain-text context file, so an AI assistant can read the whole project at once.

## Usage

1. Download `PBIP_Context.cmd`.
2. Put it in your Power BI project folder (the one containing `YourReport.pbip`).
3. Double-click it (or run `PBIP_Context.cmd /nopause` from a terminal).
4. Give the generated `PBIP_Context.txt` to your AI of choice.

## What it produces

`PBIP_Context.txt` contains:

1. A folder tree of the project.
2. The full contents of every text file, each preceded by a `FILE: relative\path` header.

Included file types: `.pbip .pbir .pbism .pbidataset .pbivis .json .tmdl .bim .platform .dax .m .pq .md .txt .xml .yaml .yml`

Skipped: the `.pbi` cache folder, `.git`, `node_modules`, binary files (images, `.abf`), the output file and the script itself.

To include other file types, edit the `EXTS` line near the top of the script.

## Notes

- Pure batch, no dependencies. Works on Windows 10/11.
- Large legacy `report.json` files can be very big; consider PBIR format (enhanced report format) for smaller, more AI-friendly output.
- Review the output before sharing: it may contain server names, queries or other details from your model.

## License

MIT

# Jabalpur Retail Region Sales Dashboard

Open [index.html](./index.html) in a browser to use the dashboard.

What is included:
- Dark-themed sales dashboard for daily register analysis
- Master-page Excel/CSV upload so future daily sales row files can refresh the dashboard instantly
- Home page with KPIs, grouped MS/HSD/Power Hist. charts, and compact target gauge charts
- Separate register page with export actions and a Home button
- Register page filters and nil-sale outlet list for generating the daily register view
- Separate monthly register page with present, Hist., growth, and Power conversion columns
- Separate master page to add new outlets, manual Hist. reference data, and sales-area targets
- Hist. Excel backup upload for `MS`, `HSD`, and `Power` on the master page
- CSV export for the generated register and outlet summary
- Print-friendly layout for PDF output

Master data workflow:
- Open [master.html](./master.html)
- Use `Daily import` to upload your daily raw Excel export or CSV file
- Use `Outlet master` to add a newly appointed retail outlet with SAP code, region, type, and plant details
- Use `Hist. backup import` to upload `MS HIST.xlsx`, `HSD HIST.xlsx`, or raw export files like `Apr-Mar 26 EXPORT.XLSX`
- Use `Hist. data` to save manual `MS`, `HSD`, or `Power` KL values for a chosen outlet and date
- Use `Targets` to save sales-area target KL for `MS`, `HSD`, `Power`, `Lube`, and `DEF`
- Saved master records are used by the register page for outlet filters and growth comparison
- Current browser import support is for `.xlsx` and `.xlsm` Hist. files

Register workflow:
- Open [register.html](./register.html) for daily outlet/date rows, including nil-sale outlets
- Open [monthly-register.html](./monthly-register.html) for month-wise outlet rows with present, Hist., growth, Power conversion, and nil-sale outlets

Master backup workflow:
- Use `Backup` on [master.html](./master.html) to download a JSON backup of custom outlets, manual Hist. entries, imported Hist. backups, sales-area targets, and the active dashboard dataset
- Use `Restore backup` on [master.html](./master.html) to upload that JSON backup into the browser later
- Use `Restore` only when you want to reset the Master page back to the built-in state

Current workbook coverage:
- Transaction data is available for `MS`, `HSD`, and `Power`
- `Lube` and `DEF` are shown in the dashboard, but the current workbook does not contain populated transaction rows for them

For future uploads inside the page:
- Open [master.html](./master.html)
- Use the `Daily import` tab and choose `Daily sales file`
- The page will automatically rebuild from the uploaded Excel or CSV file
- The single-sheet raw export format used by `try.XLSX` is supported
- Raw export `Volume` is treated as litres and converted to KL
- A sample upload layout is available in [daily-sales-upload-template.csv](./daily-sales-upload-template.csv)

Recommended Excel/CSV columns:
- `Date` or `Billing Date`
- `Sales Area` or `Sales Group Desc`
- `Outlet Name` or `Ship-to-Party Name`
- `SAP Code` or `Ship-to Party`
- `Product` or `Material Description`
- `Sales KL` or raw export `Volume`
- `Sales Unit` when using raw export `Volume`
- `Net Volume` or `Quantity`
- Optional: `Billing Document`, `Plant Description`, `Documents`, `Outlet Type`

Raw export product mapping:
- `EBMS P95-HSN-27101242` is treated as `Power`
- `EBMS-HSN-27101242` is treated as `MS`
- `EBMS-HSN-27101243` is treated as `MS`
- `HSD - BS VI` is treated as `HSD`

Built-in data now uses:
- Present sales source: `new EXPORT Apr 26.XLSX` + `EXPORT MAY 26.xlsx`
- Hist. source: `Apr-Mar 26 EXPORT.XLSX`
- Project Abhuyaday outlet/target source: `abhuvaday.xlsx`
- Abhuyaday FY 26-27 MS/HSD targets are loaded as all-month monthly targets by dividing the FY target by 12.

If you still want to refresh the built-in snapshot from an updated workbook, run:

```powershell
& 'C:\Users\Rohit\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' `
  '.\generate_sales_data.py' `
  --input 'C:\path\to\your\updated-sales-file.xlsx' `
  --historical-input 'C:\path\to\your\Apr-Mar 26 EXPORT.XLSX' `
  --abhuyaday-input 'C:\path\to\your\abhuvaday.xlsx' `
  --output '.\sales-data.js'
```

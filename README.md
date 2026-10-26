# Customer Tally Check

A simple web page that checks your team's input sheet against the master files and shows which customers don't match.

**Your files never leave your computer.** The page reads them inside your web browser. Nothing is uploaded or stored anywhere.

## How to use it

1. Open the Customer Tally Check link in Chrome, Edge, Safari or Firefox.
2. Add your files by clicking **Choose file** or by dragging a file onto a box. Each box takes `.xlsx`, `.xls` or `.csv` files.
   - **Box 1, Transaction master:** every transaction (Customer ID, Customer Name, Date, Transaction Type, Amount, Current Balance).
   - **Box 2, Customer summary:** one row per customer (Opening Balance, Total Deposits, Total Withdrawals, Current Balance and so on).
   - **Box 3, Input to check:** the sheet you want to check. It needs a **Customer ID** column plus any of the summary columns, in any order. A title row above the headers is fine.
3. You need box 3 and at least one of boxes 1 and 2. If both masters are added, the page also checks the two masters against each other.
4. Optional: set **Allowed difference ₹** if small differences (for example ₹0.05) should count as a match. The default is ₹0.00.
5. Click **Run tally**.

## Reading the results

- **Matched:** every figure agrees with the master.
- **Mismatch:** at least one figure is different. The wrong value is crossed out in red, with the correct master value and the difference underneath.
- **Not in master:** the Customer ID isn't in the master files.
- **Needs a look:** the figures agree, but the Customer ID appears twice in your sheet, or the two master files disagree for this customer.

Use the filter buttons and the search box to find customers. The screen shows up to 500 rows. Download the Excel file to see them all.

## Downloads

- **Download marked Excel:** your sheet with a Tally Status and Remarks column. Wrong cells are red, and you can hover over one to see the master value. There's also a Mismatches sheet, and a Master Check sheet if the masters disagree.
- **Download mismatch list (CSV):** every customer that isn't Matched, with details.

## Good to know

- Column names are matched loosely. For example, "Closing Balance" counts as Current Balance, "Total Credits" as Total Deposits, "Total Debits" as Total Withdrawals and "Cust ID" as Customer ID. Capital letters, "(Rs)" and "₹" don't matter.
- Amounts are compared to the paisa. Commas, ₹ signs and numbers stored as text are fine.
- Dates can be DD-MM-YYYY, DD/MM/YYYY, YYYY-MM-DD or real Excel dates.
- If you see "couldn't load", check your internet connection and reload. The page needs to reach cdnjs.cloudflare.com and cdn.jsdelivr.net.

## For whoever maintains this

It's a plain static site (`index.html`, `styles.css`, `app.js`) with no build step. It's hosted on Vercel, which redeploys automatically when changes are pushed to GitHub. Never commit customer data. `.gitignore` blocks `.csv`, `.xlsx` and `.xls` files and the `samples` folder.

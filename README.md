# 📱 Inventory Management System (Excel)

An Excel-based inventory tracker for a smartphone stockist. Record stock coming in and going out, and the live balance and low-stock alerts update automatically.

## What's inside
- **Product sheet:** master list of 15 phones with unique product codes (M001–M015)
- **Stock-in log:** date, product, auto-filled code and quantity received
- **Stock-out log:** date, product, auto-filled code and quantity sold
- **Balance-in-stock table:** total purchases, total sales and current stock per product
- **Alerts column:** flags any product with fewer than 5 units left (with conditional formatting and data bars)
- **Dropdown lists** on the product columns so entries stay consistent

## Tools
Microsoft Excel (INDEX/MATCH, IFERROR, SUMIF, IF, Data Validation, Conditional Formatting)

## File
`INVENTORY_MANAGEMENT_SYSTEM.xlsx`

## Formulas used
(all on the `inventory system` sheet, row 7 shown, drag down)

Product code lookup (stock-in and stock-out)
=IFERROR(INDEX($P$7:$P$21, MATCH(D7, $O$7:$O$21, 0)), "")

Total purchases per product
=SUMIF($D$7:$D$21, O7, $F$7:$F$21)

Total sales per product
=SUMIF($J$7:$J$21, O7, $L$7:$L$21)

Stock balance
=Q7-R7

Low-stock alert
=IF(S7<5, "LOW STOCK", "")

# excel-vba-automation
A simple but powerful VBA tool that searches an Order ID across multiple sheets (Orders, Returns, Managers) and generates an automatic summary report. Demonstrates multi-sheet lookup, dictionary mapping, and clean reporting using VBA.

This project showcases a simple but powerful Excel VBA automation workflow used to:
- Search an Order ID across multiple sheets  
- Retrieve shipment, profit, quantity, sales, and region information  
- Assign the corresponding regional manager  
- Check whether the order was returned  
- Generate a clean report with a dynamic chart  
- Trigger everything with a single button  

The repository includes:
- `excel-vba-automation.xlsm` — the complete Excel file with all macros and UI elements  

---

#  Features

# Automated Order Lookup
- Searches an Order ID entered in `Report!A1`
- Finds the matching entry in the **Orders** sheet (column *Order ID*)
- Reads:
  - Ship Date  
  - Profit  
  - Quantity  
  - Sales  
  - Region  

###  Manager Assignment
- Uses the **Users** sheet to match each region with the correct manager

### Return Status Check
- Checks the **Returns** sheet to determine whether the order was returned

### Auto-Generated Dashboard
- Clears previous results  
- Populates new metrics  
- Creates a fresh chart summarizing Profit, Quantity, and Sales  

###  Single-Click Execution
A button (“Generate Report”) runs the entire macro instantly.

---

## VBA Macro Used

The main logic is implemented in a single VBA procedure called:

Sub GenerarReporte()
It performs:

Input validation

Dictionary building for managers

Order matching

Return lookup

Report population

Chart generation
---

# How to Use
Download or clone this repository.

Open excel-vba-automation.xlsm in Excel.

Enable macros when prompted.

Go to the Report sheet.

Enter an Order ID in cell A1.

Click Generate Report.

Your dashboard will update automatically.

# Requirements
Microsoft Excel (with macro support)

Macros enabled (.xlsm file)

No external dependencies

# License
This project is provided for learning and demonstration purposes.

# Author

Luisa Gabriela Hernández

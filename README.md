
# Automated Quality Check – Accounts Payable

This project implements an automated, rule-based Quality Check solution for Accounts Payable invoices.
It applies amount thresholds (adjustable by currency for different markets), distinguishes PO and Non‑PO invoices,
and selects a controlled random sample for further manual review.

The solution was rolled out globally to standardize quality controls across markets.

## Key Logic
- Exclude invoices processed via automated booking
- Remove all invoices below 1,000 EUR
- Automatically include all invoices above 100,000 EUR
- Generate a random key for each remaining invoice
- Select a 5% random sample using Excel `xlTop10Percent` logic

## Example – Core Logic (Simplified VBA)

```vb
Sub Random_sample_QC_EUR()

    ' Filter out automated invoice
    Range("A:U").AutoFilter Field:=2, Criteria1:=Array( _
        "Autosend to SAP", "Autosend to SAP (AutoComplete)", _
        "Autostart Workflow (AutoComplete)", "Touchless (AutoComplete)"),
Operator:= _ xlFilterValues

    ' Remove low-risk invoices (<= 1,000 EUR)
    Range("A:T").AutoFilter Field:=13, Criteria1:="<=1000"
    Range("A:T").Offset(1, 0).Rows.SpecialCells(xlCellTypeVisible).Delete

    ' Flag high-risk invoices (>= 100,000 EUR)
    Range("A:T").AutoFilter Field:=13, Criteria1:=">=100000"

    ' Generate random sampling key
    Dim i As Long
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
        Cells(i, 1).Value = Rnd()
    Next i

    ' Select 5% random sample for Quality Check
    Range("A:T").AutoFilter Field:=1, Criteria1:="5", Operator:=xlTop10Percent

End Sub


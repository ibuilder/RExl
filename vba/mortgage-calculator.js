' Real Estate Mortgage Calculator
' Created with multiple sheets for different functions

' ========== MAIN CALCULATOR SHEET ==========
' Sheet 1: Calculator
' This sheet will be the main interface for inputs and summary outputs

' Set up header
Worksheets("Sheet1").Name = "Calculator"
Range("A1:G1").Merge
Range("A1").Value = "REAL ESTATE MORTGAGE CALCULATOR"
Range("A1").Font.Size = 16
Range("A1").Font.Bold = True
Range("A1").HorizontalAlignment = xlCenter

' Create sections
Range("A3").Value = "LOAN INFORMATION"
Range("A3").Font.Bold = True
Range("A13").Value = "PROPERTY TAXES & INSURANCE"
Range("A13").Font.Bold = True
Range("A19").Value = "CLOSING COSTS & FEES"
Range("A19").Font.Bold = True
Range("A25").Value = "PAYMENT SUMMARY"
Range("A25").Font.Bold = True

' Loan Information Section
Range("A4").Value = "Purchase Price ($):"
Range("A5").Value = "Down Payment ($):"
Range("A6").Value = "Down Payment (%):"
Range("A7").Value = "Loan Amount ($):"
Range("A8").Value = "Interest Rate (%):"
Range("A9").Value = "Loan Term (years):"
Range("A10").Value = "Loan Start Date:"
Range("A11").Value = "Payment Type:"

' Set up input fields
Range("C4").Value = 300000
Range("C5").Value = 60000
Range("C6").Formula = "=C5/C4"
Range("C7").Formula = "=C4-C5"
Range("C8").Value = 5.75
Range("C9").Value = 30
Range("C10").Formula = "=TODAY()"
Range("C11").Value = "Standard"

' Format cells
Range("C4:C5,C7").NumberFormat = "$#,##0.00"
Range("C6").NumberFormat = "0.00%"
Range("C8").NumberFormat = "0.000%"
Range("C9").NumberFormat = "0"
Range("C10").NumberFormat = "mm/dd/yyyy"

' Create loan type dropdown
Range("C11").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Standard,Balloon,Interest Only"

' Property Taxes & Insurance Section
Range("A14").Value = "Annual Property Tax ($):"
Range("A15").Value = "Annual Property Tax Rate (%):"
Range("A16").Value = "Annual Homeowners Insurance ($):"
Range("A17").Value = "Monthly PMI (%):"

Range("C14").Value = 3000
Range("C15").Formula = "=C14/C4"
Range("C16").Value = 1200
Range("C17").Value = 0.5

Range("C14,C16").NumberFormat = "$#,##0.00"
Range("C15,C17").NumberFormat = "0.000%"

' Conditional PMI calculation
Range("E17").Formula = "=IF(C5/C4<0.2,C7*C17/12,0)"
Range("E17").NumberFormat = "$#,##0.00"
Range("D17").Value = "Monthly PMI Amount:"

' Closing Costs & Fees Section
Range("A20").Value = "Loan Origination Fee (%):"
Range("A21").Value = "Other Closing Costs ($):"
Range("A22").Value = "Total Closing Costs ($):"

Range("C20").Value = 1
Range("C21").Value = 2500
Range("C22").Formula = "=(C7*C20/100)+C21"

Range("C20").NumberFormat = "0.000%"
Range("C21:C22").NumberFormat = "$#,##0.00"

' Payment Summary Section
Range("A26").Value = "Principal & Interest Payment:"
Range("A27").Value = "Monthly Property Tax:"
Range("A28").Value = "Monthly Insurance:"
Range("A29").Value = "Monthly PMI:"
Range("A30").Value = "Total Monthly Payment:"
Range("A31").Value = "Balloon Payment (if applicable):"

' Calculate P&I payment based on loan type
Range("C26").Formula = "=IF(C11=""Standard"",PMT(C8/12,C9*12,-C7),IF(C11=""Interest Only"",C7*C8/12,PMT(C8/12,C9*12,-C7,-(C7*0.7))))"
Range("C27").Formula = "=C14/12"
Range("C28").Formula = "=C16/12"
Range("C29").Formula = "=E17"
Range("C30").Formula = "=SUM(C26:C29)"
Range("C31").Formula = "=IF(C11=""Balloon"",C7*0.7,0)"

Range("C26:C31").NumberFormat = "$#,##0.00"

' Add balloon payment explanation
Range("E31").Formula = "=IF(C11=""Balloon"",""(Due at end of term)"","""")"

' Create a "View Amortization Schedule" button
ActiveSheet.Buttons.Add(300, 350, 150, 30).Select
Selection.OnAction = "ViewAmortizationSchedule"
Selection.Characters.Text = "View Amortization Schedule"

' ========== AMORTIZATION SCHEDULE SHEET ==========
' Sheet 2: Amortization
Sheets.Add After:=Sheets(Sheets.Count)
Worksheets("Sheet2").Name = "Amortization"

Range("A1:G1").Merge
Range("A1").Value = "AMORTIZATION SCHEDULE"
Range("A1").Font.Size = 16
Range("A1").Font.Bold = True
Range("A1").HorizontalAlignment = xlCenter

' Create column headers
Range("A3").Value = "Payment #"
Range("B3").Value = "Payment Date"
Range("C3").Value = "Beginning Balance"
Range("D3").Value = "Payment Amount"
Range("E3").Value = "Principal"
Range("F3").Value = "Interest"
Range("G3").Value = "Ending Balance"
Range("H3").Value = "Cumulative Interest"

Range("A3:H3").Font.Bold = True

' Format columns
Range("B:B").NumberFormat = "mm/dd/yyyy"
Range("C:H").NumberFormat = "$#,##0.00"

' Set up formulas for amortization calculation
Range("A4").Formula = "1"
Range("B4").Formula = "=Calculator!C10+30"
Range("C4").Formula = "=Calculator!C7"
Range("D4").Formula = "=IF(Calculator!C11=""Standard"",ABS(Calculator!C26),IF(Calculator!C11=""Interest Only"",C4*Calculator!C8/12,ABS(Calculator!C26)))"
Range("E4").Formula = "=IF(Calculator!C11=""Interest Only"",0,D4-F4)"
Range("F4").Formula = "=C4*Calculator!C8/12"
Range("G4").Formula = "=C4-E4"
Range("H4").Formula = "=F4"

' Add rows for each payment
Range("A5").Formula = "=A4+1"
Range("B5").Formula = "=EDATE(B4,1)"
Range("C5").Formula = "=G4"
Range("D5").Formula = "=IF(Calculator!C11=""Standard"",ABS(Calculator!C26),IF(Calculator!C11=""Interest Only"",C5*Calculator!C8/12,ABS(Calculator!C26)))"
Range("E5").Formula = "=IF(Calculator!C11=""Interest Only"",0,D5-F5)"
Range("F5").Formula = "=C5*Calculator!C8/12"
Range("G5").Formula = "=C5-E5"
Range("H5").Formula = "=H4+F5"

' Fill down formulas for 360 payments (30 years)
Range("A5:H5").Copy
Range("A6:H365").PasteSpecial xlPasteFormulas

' Add balloon payment at the end if applicable
Range("I3").Value = "Balloon Payment"
Range("I3").Font.Bold = True
Range("I4").Formula = "=IF(Calculator!C11=""Balloon"",IF(A4=Calculator!C9*12,Calculator!C31,0),0)"
Range("I4").NumberFormat = "$#,##0.00"
Range("I4").Copy
Range("I5:I365").PasteSpecial xlPasteFormulas

' Add a "Return to Calculator" button
ActiveSheet.Buttons.Add(300, 25, 150, 30).Select
Selection.OnAction = "ReturnToCalculator"
Selection.Characters.Text = "Return to Calculator"

' ========== LOAN COMPARISON SHEET ==========
' Sheet 3: Loan Comparison
Sheets.Add After:=Sheets(Sheets.Count)
Worksheets("Sheet3").Name = "Loan Comparison"

Range("A1:G1").Merge
Range("A1").Value = "LOAN COMPARISON CALCULATOR"
Range("A1").Font.Size = 16
Range("A1").Font.Bold = True
Range("A1").HorizontalAlignment = xlCenter

' Create comparison table
Range("A3:G3").Merge
Range("A3").Value = "COMPARE DIFFERENT LOAN OPTIONS"
Range("A3").Font.Bold = True

Range("A5").Value = "Loan Option"
Range("B5").Value = "Rate (%)"
Range("C5").Value = "Term (Years)"
Range("D5").Value = "Monthly P&I"
Range("E5").Value = "Total Monthly Payment"
Range("F5").Value = "Total Interest Paid"
Range("G5").Value = "Total Cost"

Range("A5:G5").Font.Bold = True

' Set up comparison rows
Range("A6").Value = "Option 1 (Current)"
Range("A7").Value = "Option 2"
Range("A8").Value = "Option 3"

' Link first option to main calculator
Range("B6").Formula = "=Calculator!C8"
Range("C6").Formula = "=Calculator!C9"
Range("D6").Formula = "=ABS(Calculator!C26)"
Range("E6").Formula = "=Calculator!C30"
Range("F6").Formula = "=IF(Calculator!C11=""Standard"",Amortization!H365,Calculator!C8/12*Calculator!C7*Calculator!C9*12)"
Range("G6").Formula = "=F6+Calculator!C7+Calculator!C22"

' Set up option 2 with example values
Range("B7").Value = 6
Range("C7").Value = 15
Range("D7").Formula = "=PMT(B7/100/12,C7*12,-Calculator!C7)"
Range("E7").Formula = "=D7+Calculator!C27+Calculator!C28+Calculator!C29"
Range("F7").Formula = "=D7*C7*12-Calculator!C7"
Range("G7").Formula = "=F7+Calculator!C7+Calculator!C22"

' Set up option 3 with example values
Range("B8").Value = 5.5
Range("C8").Value = 30
Range("D8").Formula = "=PMT(B8/100/12,C8*12,-Calculator!C7)"
Range("E8").Formula = "=D8+Calculator!C27+Calculator!C28+Calculator!C29"
Range("F8").Formula = "=D8*C8*12-Calculator!C7"
Range("G8").Formula = "=F8+Calculator!C7+Calculator!C22"

' Format cells
Range("B6:B8").NumberFormat = "0.000%"
Range("C6:C8").NumberFormat = "0"
Range("D6:G8").NumberFormat = "$#,##0.00"

' Add a "Return to Calculator" button
ActiveSheet.Buttons.Add(300, 25, 150, 30).Select
Selection.OnAction = "ReturnToCalculator"
Selection.Characters.Text = "Return to Calculator"

' ========== AFFORDABILITY CALCULATOR SHEET ==========
' Sheet 4: Affordability
Sheets.Add After:=Sheets(Sheets.Count)
Worksheets("Sheet4").Name = "Affordability"

Range("A1:G1").Merge
Range("A1").Value = "AFFORDABILITY CALCULATOR"
Range("A1").Font.Size = 16
Range("A1").Font.Bold = True
Range("A1").HorizontalAlignment = xlCenter

' Create income and expense inputs
Range("A3").Value = "INCOME & EXPENSE INFORMATION"
Range("A3").Font.Bold = True

Range("A5").Value = "Gross Annual Income ($):"
Range("A6").Value = "Monthly Debt Payments ($):"
Range("A7").Value = "Desired Monthly Payment ($):"
Range("A8").Value = "Maximum DTI Ratio (%):"

Range("C5").Value = 85000
Range("C6").Value = 500
Range("C7").Formula = "=C5/12*0.28"
Range("C8").Value = 0.36

Range("C5").NumberFormat = "$#,##0.00"
Range("C6:C7").NumberFormat = "$#,##0.00"
Range("C8").NumberFormat = "0.00%"

' Create affordability calculator
Range("A10").Value = "AFFORDABILITY RESULTS"
Range("A10").Font.Bold = True

Range("A12").Value = "Interest Rate (%):"
Range("A13").Value = "Loan Term (years):"
Range("A14").Value = "Property Tax Rate (%):"
Range("A15").Value = "Annual Insurance Rate (%):"
Range("A16").Value = "Down Payment (%):"

Range("C12").Formula = "=Calculator!C8"
Range("C13").Formula = "=Calculator!C9"
Range("C14").Formula = "=Calculator!C15"
Range("C15").Formula = "=Calculator!C16/Calculator!C4"
Range("C16").Formula = "=Calculator!C6"

Range("C12").NumberFormat = "0.000%"
Range("C13").NumberFormat = "0"
Range("C14:C16").NumberFormat = "0.000%"

Range("A18").Value = "Maximum Affordable Loan:"
Range("A19").Value = "Maximum Home Price:"
Range("A20").Value = "Down Payment Amount:"
Range("A21").Value = "Monthly Principal & Interest:"
Range("A22").Value = "Monthly Property Taxes:"
Range("A23").Value = "Monthly Insurance:"
Range("A24").Value = "Total Monthly Payment:"

' Calculate maximum affordable home price based on income and expenses
Range("C18").Formula = "=PV(C12/12,C13*12,-MAX(0,(C5/12*C8)-C6))"
Range("C19").Formula = "=C18/(1-C16)"
Range("C20").Formula = "=C19*C16"
Range("C21").Formula = "=PMT(C12/12,C13*12,-C18)"
Range("C22").Formula = "=C19*C14/12"
Range("C23").Formula = "=C19*C15/12"
Range("C24").Formula = "=SUM(ABS(C21),C22,C23)"

Range("C18:C24").NumberFormat = "$#,##0.00"

' Add a "Return to Calculator" button
ActiveSheet.Buttons.Add(300, 25, 150, 30).Select
Selection.OnAction = "ReturnToCalculator"
Selection.Characters.Text = "Return to Calculator"

' ========== VBA CODE FOR BUTTONS ==========
' This would need to be added to the VBA module:
'
' Sub ViewAmortizationSchedule()
'     Sheets("Amortization").Activate
' End Sub
'
' Sub ReturnToCalculator()
'     Sheets("Calculator").Activate
' End Sub

' Return to the main Calculator sheet
Sheets("Calculator").Activate

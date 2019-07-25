Attribute VB_Name = "Module8"
Private Sub Workbook_Open()

'Get Sheets
orderChecklist = "Order Checklist"
accountSheet = "Account Info-DO NOT DELETE"
equipmentInfoSheet = "Equip. Info-DO NOT DELETE"
financialInfoSheet = "Financial Info-DO NOT DELETE"
InstructionsSheet = "Instructions"
lpmSheet = "Lease Price Model 2.0"
bosSheet = "BoS 2.0"
bosTc = "BoS - T & C"
LASheet = "Lease Agreement 2.0"
LeaseTc = "Lease - T & C"


'Get Account, Customer, and Rep
repName = Sheets(accountSheet).Cells(12, 2).Value
accountNumber = Sheets(accountSheet).Cells(17, 2).Value
customerName = Sheets(accountSheet).Cells(21, 2).Value

'insert values
Sheets(orderChecklist).Cells(1, 10).Value = repName
Sheets(orderChecklist).Cells(2, 10).Value = Date
Sheets(orderChecklist).Cells(4, 3).Value = customerName
Sheets(orderChecklist).Cells(5, 3).Value = accountNumber

'******************************
'Hide forms based on deal type'
'******************************
leasePayment = Sheets(financialInfoSheet).Cells(19, 8).Value

'Hide Data Dump
ActiveWorkbook.Sheets(accountSheet).Visible = xlSheetVeryHidden
ActiveWorkbook.Sheets(financialInfoSheet).Visible = xlSheetVeryHidden
ActiveWorkbook.Sheets(equipmentInfoSheet).Visible = xlSheetVeryHidden

If leasePayment > 0 Then
    'hide bill of sale sheets
    ActiveWorkbook.Sheets(InstructionsSheet).Visible = xlSheetVeryHidden
    ActiveWorkbook.Sheets(bosSheet).Visible = xlSheetVeryHidden
    ActiveWorkbook.Sheets(bosTc).Visible = xlSheetVeryHidden
Else
    'hide lease sheets
    ActiveWorkbook.Sheets(lpmSheet).Visible = xlSheetVeryHidden
    ActiveWorkbook.Sheets(LASheet).Visible = xlSheetVeryHidden
    ActiveWorkbook.Sheets(LeaseTc).Visible = xlSheetVeryHidden
End If

End Sub


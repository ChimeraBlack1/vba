Attribute VB_Name = "Module8"
Sub OrderChecklist_Button1_Click()

'******************************
'Lease Type'
'******************************
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
lpmSheet = "Lease Price Model 2.0"

'Hide Data Dump
ActiveWorkbook.Sheets(accountSheet).Visible = xlSheetVeryHidden
ActiveWorkbook.Sheets(financialInfoSheet).Visible = xlSheetHidden
ActiveWorkbook.Sheets(equipmentInfoSheet).Visible = xlSheetHidden

If Sheets(bosSheet).Visible = xlSheetHidden Then
    'Show lease sheets
    ActiveWorkbook.Sheets(lpmSheet).Visible = True
    ActiveWorkbook.Sheets(LASheet).Visible = True
    ActiveWorkbook.Sheets(LeaseTc).Visible = True

    'Hide show BoS Sheets
    ActiveWorkbook.Sheets(InstructionsSheet).Visible = xlSheetHidden
    ActiveWorkbook.Sheets(bosSheet).Visible = xlSheetHidden
    ActiveWorkbook.Sheets(bosTc).Visible = xlSheetHidden
Else

    'Show show BoS Sheets
    ActiveWorkbook.Sheets(InstructionsSheet).Visible = True
    ActiveWorkbook.Sheets(bosSheet).Visible = True
    ActiveWorkbook.Sheets(bosTc).Visible = True
    

End If



End Sub
Sub OrderChecklist_Button2_Click()


'******************************
'BoS Type'
'******************************

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
lpmSheet = "Lease Price Model 2.0"


If Sheets(lpmSheet).Visible = True Then
    'Show BoS Sheets
    ActiveWorkbook.Sheets(InstructionsSheet).Visible = True
    ActiveWorkbook.Sheets(bosSheet).Visible = True
    ActiveWorkbook.Sheets(bosTc).Visible = True
    
    'Hide Lease Machines
    ActiveWorkbook.Sheets(lpmSheet).Visible = xlSheetHidden
    ActiveWorkbook.Sheets(LASheet).Visible = xlSheetHidden
    ActiveWorkbook.Sheets(LeaseTc).Visible = xlSheetHidden
    
    Sheets(orderChecklist).Activate
    
ElseIf Sheets(lpmSheet).Visible = False Then

    'Show Lease Machines
    ActiveWorkbook.Sheets(lpmSheet).Visible = True
    ActiveWorkbook.Sheets(LASheet).Visible = True
    ActiveWorkbook.Sheets(LeaseTc).Visible = True
    
    Sheets(orderChecklist).Activate

End If





End Sub
Sub Button3_Click()

'******************************
'<-- Click to Fill'
'******************************

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

End Sub

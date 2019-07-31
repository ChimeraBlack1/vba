Attribute VB_Name = "Module9"
Sub OrderChecklist_Button1_Click()

End Sub
Sub Button6_Click()
'******************************
'Lease Button'
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

If Sheets(bosSheet).Visible = True Then
    'Show lease sheets
    ActiveWorkbook.Sheets(lpmSheet).Visible = True
    ActiveWorkbook.Sheets(LASheet).Visible = True
    ActiveWorkbook.Sheets(LeaseTc).Visible = True

    'Hide show BoS Sheets
    'ActiveWorkbook.Sheets(InstructionsSheet).Visible = xlSheetHidden
    ActiveWorkbook.Sheets(bosSheet).Visible = xlSheetHidden
    ActiveWorkbook.Sheets(bosTc).Visible = xlSheetHidden
    
    'Hide Data Dump
    ActiveWorkbook.Sheets(accountSheet).Visible = xlSheetVeryHidden
    ActiveWorkbook.Sheets(financialInfoSheet).Visible = xlSheetHidden
    ActiveWorkbook.Sheets(equipmentInfoSheet).Visible = xlSheetHidden
    
    Sheets(orderChecklist).Activate
Else

    'Show show BoS Sheets
    'ActiveWorkbook.Sheets(InstructionsSheet).Visible = True
    ActiveWorkbook.Sheets(bosSheet).Visible = True
    ActiveWorkbook.Sheets(bosTc).Visible = True
    
    'Show Data Dump
    ActiveWorkbook.Sheets(accountSheet).Visible = True
    ActiveWorkbook.Sheets(financialInfoSheet).Visible = True
    ActiveWorkbook.Sheets(equipmentInfoSheet).Visible = True
    
    Sheets(orderChecklist).Activate
    
End If


'todo - Call the scripts that will fill out the sheets

'Click to fill
Call Button3_Click

'Lease - LPM, Lease Agreement
'LPM
Call LeasePriceModel20_Button1_Click
'LA
Call Button1_Click

'Fires every time - Service Agreement, INQ
Call Button2_Click
Call IQ20_Button1_Click

'Reactivate checkllist
Sheets(orderChecklist).Activate
End Sub

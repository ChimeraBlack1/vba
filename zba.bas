Attribute VB_Name = "Module7"
Sub LeasePriceModel20_Button2_Click()

'ActiveSheet.Unprotect Password:="sherpadoc1"

accountInfoSheet = "Account Info-DO NOT DELETE"
standardSheets = 15
mySheets = Worksheets.Count
zbaStart = 16
transType = Sheets(accountInfoSheet).Cells(14, 4).Value

For i = standardSheets To mySheets
    
    Cells(zbaStart, 41).Value = transType

    zbaStart = zbaStart + 2
Next i

'ActiveSheet.Protect Password:="sherpadoc1"

End Sub



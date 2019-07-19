Attribute VB_Name = "Module7"
Sub LeasePriceModel20_Button2_Click()

ActiveSheet.Unprotect Password:="sherpadoc1"

standardSheets = 15
mySheets = Worksheets.Count
zbaStart = 16

For i = standardSheets To mySheets
    Zba = Cells(zbaStart, 41).Value
    
    If Zba = "ZBA" Then
        Cells(zbaStart, 41).Value = ""
    Else
        Cells(zbaStart, 41).Value = "ZBA"
    End If
    
    zbaStart = zbaStart + 2
Next i

ActiveSheet.Protect Password:="sherpadoc1"

End Sub

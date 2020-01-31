Attribute VB_Name = "Module1"
Sub Button2048_Click()

For i = 1 To 2500
    If Sheets("Sheet1").Cells(i, 1).Value = "" Then
        Rows(i).EntireRow.Delete
    End If
Next i


    
End Sub

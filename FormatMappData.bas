Attribute VB_Name = "Module1"
Sub FormatData()


linesToProcess = Application.WorksheetFunction.Count(Range("A1:A1000"))

insertRowPointer = 1

For i = 1 To linesToProcess
    
    'insert row
    myRange = insertRowPointer & ":" & insertRowPointer
    Range(myRange).Insert
    
    'copy desc
    desc = Cells(insertRowPointer + 1, 2).Value
    Cells(insertRowPointer, 2).Value = desc
    
    insertRowPointer = insertRowPointer + 2

    
Next i

'Range("1:1").Insert

'Cells(1, 2).Value = desc

'Range("A1:A2").Merge
'Range("C1:C2").Merge
'Range("D1:D2").Merge
'Range("E1:E2").Merge
'Range("F1:F2").Merge

End Sub

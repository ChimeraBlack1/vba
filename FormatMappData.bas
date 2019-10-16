Attribute VB_Name = "Module1"
Sub FormatData()


'linesToProcess = Application.WorksheetFunction.CountIf(Range("A1:A1000"), "*")
'linesToProcess2 = Application.WorksheetFunction.Count(Range("A1:A1000"))
lineCount = Range("A1").End(xlDown).Row

insertRowPointer = 1

For i = 1 To lineCount
    
    'insert row
    myRange = insertRowPointer & ":" & insertRowPointer
    Range(myRange).Insert
    
    'copy desc
    desc = Cells(insertRowPointer + 1, 2).Value
    Cells(insertRowPointer, 2).Value = desc
    
    'merge cells
    mergeRangeA = "A" & insertRowPointer & ":" & "A" & (insertRowPointer + 1)
    mergeRangeC = "C" & insertRowPointer & ":" & "C" & (insertRowPointer + 1)
    mergeRangeD = "D" & insertRowPointer & ":" & "D" & (insertRowPointer + 1)
    mergeRangeE = "E" & insertRowPointer & ":" & "E" & (insertRowPointer + 1)
    mergeRangeF = "F" & insertRowPointer & ":" & "F" & (insertRowPointer + 1)
    
    Range(mergeRangeA).Merge
    Range(mergeRangeC).Merge
    Range(mergeRangeD).Merge
    Range(mergeRangeE).Merge
    Range(mergeRangeF).Merge
    
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

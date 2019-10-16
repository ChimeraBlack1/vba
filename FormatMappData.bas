Attribute VB_Name = "Module1"
Sub FormatData()

lineCount = Range("A1").End(xlDown)

If lineCount <= 1 Then
    lineCount = 1
End If

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


End Sub

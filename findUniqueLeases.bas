Attribute VB_Name = "Module1"
Sub FindUniqueLeases()

targetSheet = "Raw Data"
endOfLeaseSheet = Sheets(targetSheet).Cells(Rows.Count, "D").End(xlUp).Row
totalLeaseValue = 0
portfolioTotal = "PortFolio Total (calc)"
portfolioRowIndex = 1
thisReport = "Sherpa Report-September 06, 2019"

For i = 2 To endOfLeaseSheet
    leaseNumber = Sheets(targetSheet).Cells(i, 4).Value
    prevLeaseNumber = Sheets(targetSheet).Cells(i - 1, 4).Value
    
    If leaseNumber <> prevLeaseNumber Then
        portfolioRowIndex = portfolioRowIndex + 1
        For j = 1 To 12
            'Copy Value to new Sheet
            Sheets(portfolioTotal).Cells(portfolioRowIndex, j).Value = Sheets(targetSheet).Cells(i, j).Value
        Next j
        
    End If
Next i

Workbooks(thisReport).Sheets(portfolioTotal).Columns("A:R").AutoFit

End Sub

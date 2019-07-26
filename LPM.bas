Attribute VB_Name = "Module6"
Sub LeasePriceModel20_Button1_Click()
 
lpmSheet = 6
leaseSheet = "Lease Price Model 2.0"
Sheets(lpmSheet).Activate
ActiveSheet.Unprotect Password:="sherpadoc1"
 
'Align Colums
Columns("A").ColumnWidth = 0.94
Columns("B").ColumnWidth = 0.94
Columns("C").ColumnWidth = 0.94
Columns("D").ColumnWidth = 0.94
Columns("E").ColumnWidth = 0.94
Columns("F").ColumnWidth = 0.94
Columns("G").ColumnWidth = 0.94
Columns("H").ColumnWidth = 0.94
Columns("I").ColumnWidth = 0.94
Columns("J").ColumnWidth = 0.94
Columns("K").ColumnWidth = 0.94
Columns("L").ColumnWidth = 0.94
Columns("M").ColumnWidth = 0.94
Columns("N").ColumnWidth = 0.94
Columns("O").ColumnWidth = 0.94
Columns("P").ColumnWidth = 0.94
Columns("Q").ColumnWidth = 0.94
Columns("R").ColumnWidth = 0.94
Columns("S").ColumnWidth = 0.94
Columns("T").ColumnWidth = 0.94
Columns("U").ColumnWidth = 0.94
Columns("V").ColumnWidth = 0.94
Columns("W").ColumnWidth = 0.94
Columns("X").ColumnWidth = 0.94
Columns("Y").ColumnWidth = 0.94
Columns("Z").ColumnWidth = 0.94
Columns("AA").ColumnWidth = 0.94
Columns("AB").ColumnWidth = 0.94
Columns("AC").ColumnWidth = 0.94
Columns("AD").ColumnWidth = 4.33
Columns("AE").ColumnWidth = 0.94
Columns("AF").ColumnWidth = 0.94
Columns("AG").ColumnWidth = 0.94
Columns("AH").ColumnWidth = 2
Columns("AI").ColumnWidth = 2
Columns("AJ").ColumnWidth = 0.94
Columns("AK").ColumnWidth = 0.94
Columns("AL").ColumnWidth = 0.94
Columns("AM").ColumnWidth = 0.94
Columns("AN").ColumnWidth = 0.94
Columns("AO").ColumnWidth = 0.94
Columns("AP").ColumnWidth = 0.94
Columns("AQ").ColumnWidth = 0.94
Columns("AR").ColumnWidth = 0.94
Columns("AS").ColumnWidth = 0.94
Columns("AT").ColumnWidth = 0.94
Columns("AU").ColumnWidth = 0.94
Columns("AV").ColumnWidth = 0.94
Columns("AW").ColumnWidth = 0.94
Columns("AX").ColumnWidth = 0.94
Columns("AY").ColumnWidth = 0.94
Columns("AZ").ColumnWidth = 0.94
Columns("BA").ColumnWidth = 0.94
Columns("BB").ColumnWidth = 0.94
Columns("BC").ColumnWidth = 0.94
Columns("BD").ColumnWidth = 0.94
Columns("BE").ColumnWidth = 0.94
Columns("BF").ColumnWidth = 0.94
Columns("BG").ColumnWidth = 0.94
 
'Font Settings
Range(Cells(14, 1), Cells(moduleStart + 40, 58)).Font.Size = 11
Range(Cells(14, 1), Cells(moduleStart + 40, 58)).Font.Name = "Times New Roman"
 
'Air % font
Range(Cells(moduleStart + 23, 34), Cells(moduleStart + 23, 34)).Font.Size = 8
 
Cells(moduleStart + 4, 18).Value = "Delivery Charge"
Range(Cells(moduleStart + 4, 18), Cells(moduleStart + 14, 18)).Font.Size = 10
Range(Cells(moduleStart + 4, 18), Cells(moduleStart + 14, 18)).Font.Name = "Arial"
 
percentageFormat = "0.00%"
accountingFormat = "_-$* #,##0.00_-;-$* #,##0.00_-;_-$* ""-""??_-;_-@_-"
moduleStart = 16
standardSheets = 15
mySheets = Worksheets.Count
leaseRate = Sheets(3).Cells(13, 4).Value
ERPNumber = Sheets(1).Cells(17, 2).Value
 
'Set ERP Number in Header
Sheets(lpmSheet).Cells(10, 40).Value = ERPNumber
Range(Cells(10, 40), Cells(10, 40)).HorizontalAlignment = xlCenter
 
'insert mainframes
For i = standardSheets To mySheets
 
    'Align Rows
  Rows(moduleStart).RowHeight = 12
    Rows(moduleStart + 1).RowHeight = 6.6
   
    'Thick Side Border
  With Worksheets(leaseSheet).Range(Cells(moduleStart, 58), Cells(moduleStart, 58)).Borders(xlEdgeRight)
       .LineStyle = xlContinuous
       .Weight = xlMedium
    End With
     
    With Worksheets(leaseSheet).Range(Cells(moduleStart - 1, 58), Cells(moduleStart - 1, 58)).Borders(xlEdgeRight)
       .LineStyle = xlContinuous
       .Weight = xlMedium
    End With
   
    Range(Cells(moduleStart, 2), Cells(moduleStart, 13)).Merge
    Range(Cells(moduleStart, 2), Cells(moduleStart, 13)).NumberFormat = accountingFormat
    Range(Cells(moduleStart, 15), Cells(moduleStart, 24)).Merge
    Range(Cells(moduleStart, 15), Cells(moduleStart, 24)).NumberFormat = accountingFormat
    Range(Cells(moduleStart, 30), Cells(moduleStart, 39)).Merge
    Range(Cells(moduleStart, 30), Cells(moduleStart, 39)).NumberFormat = accountingFormat
    Range(Cells(moduleStart, 45), Cells(moduleStart, 54)).Merge
   
    'Mapp Price
  Cells(moduleStart, 30).Value = mappPrice
   
   
    'Get mainframe name
  Cells(moduleStart, 2).Value = Sheets(i).Cells(16, 2).Value
   
   
    moduleStart = moduleStart + 2
   
Next i
   
   
'Underline
With Worksheets(leaseSheet).Range(Cells(moduleStart - 2, 15), Cells(moduleStart - 2, 24)).Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .Weight = xlThin
End With
 
With Worksheets(leaseSheet).Range(Cells(moduleStart - 2, 30), Cells(moduleStart - 2, 39)).Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .Weight = xlThin
End With
 
With Worksheets(leaseSheet).Range(Cells(moduleStart - 2, 45), Cells(moduleStart - 2, 54)).Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .Weight = xlThin
End With
 
 
'Align Footer Row above totals
Rows(moduleStart - 1).RowHeight = 6.6
 
'Merging totals
Range(Cells(moduleStart, 15), Cells(moduleStart, 24)).Merge
Range(Cells(moduleStart, 30), Cells(moduleStart, 39)).Merge
Range(Cells(moduleStart, 45), Cells(moduleStart, 54)).Merge
 
'Borders for totals
Range(Cells(moduleStart, 15), Cells(moduleStart, 24)).BorderAround ColorIndex:=1, Weight:=xlThin
Range(Cells(moduleStart, 30), Cells(moduleStart, 39)).BorderAround ColorIndex:=1, Weight:=xlThin
Range(Cells(moduleStart, 45), Cells(moduleStart, 54)).BorderAround ColorIndex:=1, Weight:=xlThin
 
'Add Footer
Cells(moduleStart, 13).Value = "Totals"
Cells(moduleStart, 13).Font.Bold = True
sellTotalFormula = "=SUM(O16:" + "O" + CStr(moduleStart - 1) + ")"
Cells(moduleStart, 15).Value = sellTotalFormula
Range(Cells(moduleStart, 13), Cells(moduleStart, 13)).HorizontalAlignment = xlRight
 
'MAPP Total
mappTotalFormula = "=SUM(AD16:" + "AD" + CStr(moduleStart - 1) + ")"
Cells(moduleStart, 30).Value = mappTotalFormula
 
'Diff to MAPP total
diffToMappTotal = "=SUM(AS16:" + "AS" + CStr(moduleStart - 1) + ")"
Cells(moduleStart, 45).Value = diffToMappTotal
 
'Equipment Sub Total
Cells(moduleStart + 2, 13).Value = "Equipment Sub-Total"
Cells(moduleStart + 2, 13).Font.Bold = True
Range(Cells(moduleStart + 2, 13), Cells(moduleStart + 2, 13)).HorizontalAlignment = xlRight
Range(Cells(moduleStart + 2, 26), Cells(moduleStart + 2, 35)).Merge
Range(Cells(moduleStart + 2, 26), Cells(moduleStart + 2, 35)).BorderAround ColorIndex:=1, Weight:=xlThin
Range(Cells(moduleStart + 2, 46), Cells(moduleStart + 2, 55)).Merge
Range(Cells(moduleStart + 2, 46), Cells(moduleStart + 2, 55)).BorderAround ColorIndex:=1, Weight:=xlThin
eqpSubFormula = "=O" + CStr(moduleStart)
Cells(moduleStart + 2, 26).Formula = eqpSubFormula
Cells(moduleStart + 2, 46).NumberFormat = percentageFormat
 
'Delivery Charge
Cells(moduleStart + 4, 18).Value = "Delivery Charge"
Range(Cells(moduleStart + 4, 27), Cells(moduleStart + 4, 35)).Merge
Cells(moduleStart + 4, 27).Value = 0
Range(Cells(moduleStart + 4, 18), Cells(moduleStart + 4, 18)).HorizontalAlignment = xlRight
 
'Removal Charge
Cells(moduleStart + 5, 18).Value = "Removal Charge"
Range(Cells(moduleStart + 5, 18), Cells(moduleStart + 5, 18)).HorizontalAlignment = xlRight
Range(Cells(moduleStart + 5, 27), Cells(moduleStart + 5, 35)).Merge
Cells(moduleStart + 5, 27).Value = 0
Range(Cells(moduleStart + 5, 27), Cells(moduleStart + 5, 35)).Interior.ColorIndex = 6
 
'Service Allocation
Cells(moduleStart + 6, 18).Value = "Service Allocation"
Range(Cells(moduleStart + 6, 18), Cells(moduleStart + 6, 18)).HorizontalAlignment = xlRight
Range(Cells(moduleStart + 6, 27), Cells(moduleStart + 6, 35)).Merge
Cells(moduleStart + 6, 27).Value = 0
Range(Cells(moduleStart + 6, 27), Cells(moduleStart + 6, 35)).Interior.ColorIndex = 6
 
'Marketing Promotion
Cells(moduleStart + 7, 18).Value = "Marketing Promotion"
Range(Cells(moduleStart + 7, 18), Cells(moduleStart + 7, 18)).HorizontalAlignment = xlRight
Range(Cells(moduleStart + 7, 27), Cells(moduleStart + 7, 35)).Merge
Cells(moduleStart + 7, 27).Value = 0
Range(Cells(moduleStart + 7, 27), Cells(moduleStart + 7, 35)).Interior.ColorIndex = 6
 
'Trade-In Amount (Discount)
Cells(moduleStart + 8, 18).Value = "Trade-In Amount (Discount)"
Range(Cells(moduleStart + 8, 18), Cells(moduleStart + 8, 18)).HorizontalAlignment = xlRight
Range(Cells(moduleStart + 8, 27), Cells(moduleStart + 8, 35)).Merge
Cells(moduleStart + 8, 27).Value = 0
Range(Cells(moduleStart + 8, 27), Cells(moduleStart + 8, 35)).Interior.ColorIndex = 6
 
'Provincial Environmental Levy
Cells(moduleStart + 9, 18).Value = "Provincial Environmental Levy"
Range(Cells(moduleStart + 9, 18), Cells(moduleStart + 9, 18)).HorizontalAlignment = xlRight
Range(Cells(moduleStart + 9, 27), Cells(moduleStart + 9, 35)).Merge
Cells(moduleStart + 9, 27).Value = 0
 
'Net Equipment Value
Range(Cells(moduleStart + 10, 27), Cells(moduleStart + 10, 35)).Merge
Cells(moduleStart + 10, 18).Value = "Net Equipment Value"
Cells(moduleStart + 10, 18).Font.Bold = True
Range(Cells(moduleStart + 10, 27), Cells(moduleStart + 10, 35)).BorderAround ColorIndex:=1, Weight:=xlThin
Range(Cells(moduleStart + 10, 18), Cells(moduleStart + 10, 18)).HorizontalAlignment = xlRight
 
'Settlement Amount
Cells(moduleStart + 11, 18).Value = "Settlement Amount"
Range(Cells(moduleStart + 11, 18), Cells(moduleStart + 11, 18)).HorizontalAlignment = xlRight
Range(Cells(moduleStart + 11, 27), Cells(moduleStart + 11, 35)).Merge
 
'Discretionary Items
Cells(moduleStart + 12, 18).Value = "Discretionary Items"
Range(Cells(moduleStart + 12, 18), Cells(moduleStart + 12, 18)).HorizontalAlignment = xlRight
Range(Cells(moduleStart + 12, 27), Cells(moduleStart + 12, 35)).Merge
Cells(moduleStart + 12, 27).Value = 0
Range(Cells(moduleStart + 12, 27), Cells(moduleStart + 12, 35)).Interior.ColorIndex = 6
 
'Invoice Price
Range(Cells(moduleStart + 14, 27), Cells(moduleStart + 14, 35)).Merge
Cells(moduleStart + 14, 18).Value = "Invoice Price"
Cells(moduleStart + 14, 18).Font.Bold = True
Range(Cells(moduleStart + 14, 18), Cells(moduleStart + 14, 18)).HorizontalAlignment = xlRight
Range(Cells(moduleStart + 14, 27), Cells(moduleStart + 14, 35)).BorderAround ColorIndex:=1, Weight:=xlThin
 
'Lease Pricing Model (Ricoh Non-Note)
Range(Cells(moduleStart + 16, 1), Cells(moduleStart + 16, 58)).Merge
Cells(moduleStart + 16, 1).Value = "B) Lease Pricing Model (Ricoh Non-Note)"
Cells(moduleStart + 16, 1).Font.Bold = True
Range(Cells(moduleStart + 16, 1), Cells(moduleStart + 16, 58)).Interior.ColorIndex = 6
Range(Cells(moduleStart + 16, 1), Cells(moduleStart + 16, 58)).BorderAround ColorIndex:=1, Weight:=xlThin
 
'Please enter amounts in green boxes
Cells(moduleStart + 17, 2).Value = "Please enter amounts in green boxes"
 
'Equipment Value
Cells(moduleStart + 18, 22).Value = "Equipment Value"
 
'New Equipment
Cells(moduleStart + 19, 2).Value = "New Equipment"
Range(Cells(moduleStart + 19, 20), Cells(moduleStart + 19, 32)).Merge
Range(Cells(moduleStart + 19, 20), Cells(moduleStart + 19, 32)).BorderAround ColorIndex:=1, Weight:=xlThin
 
'Used/Refinance Equipment
Cells(moduleStart + 21, 2).Value = "Used/Refinance Equipment"
Range(Cells(moduleStart + 21, 20), Cells(moduleStart + 21, 32)).Merge
Range(Cells(moduleStart + 21, 20), Cells(moduleStart + 21, 32)).BorderAround ColorIndex:=1, Weight:=xlThin
 
'Settlement
Cells(moduleStart + 23, 2).Value = "Settlement"
Range(Cells(moduleStart + 23, 20), Cells(moduleStart + 23, 32)).Merge
Range(Cells(moduleStart + 23, 20), Cells(moduleStart + 23, 32)).BorderAround ColorIndex:=1, Weight:=xlThin
 
'Soft Costs
Cells(moduleStart + 25, 2).Value = "Soft Costs"
Range(Cells(moduleStart + 25, 20), Cells(moduleStart + 25, 32)).Merge
Range(Cells(moduleStart + 25, 20), Cells(moduleStart + 25, 32)).BorderAround ColorIndex:=1, Weight:=xlThin
 
'Air
Cells(moduleStart + 18, 34).Value = "% Air"
 
'New Equipment Rate
Cells(moduleStart + 18, 40).Value = "Rate"
Range(Cells(moduleStart + 19, 40), Cells(moduleStart + 19, 44)).Merge
Range(Cells(moduleStart + 19, 40), Cells(moduleStart + 19, 44)).BorderAround ColorIndex:=1, Weight:=xlThin
Cells(moduleStart + 19, 40).Value = leaseRate
 
'Used/Refinance Rate
Range(Cells(moduleStart + 21, 40), Cells(moduleStart + 21, 44)).Merge
Range(Cells(moduleStart + 21, 40), Cells(moduleStart + 21, 44)).BorderAround ColorIndex:=1, Weight:=xlThin
Cells(moduleStart + 21, 40).Value = leaseRate
 
'Settlement Rate
Range(Cells(moduleStart + 23, 40), Cells(moduleStart + 23, 44)).Merge
Range(Cells(moduleStart + 23, 40), Cells(moduleStart + 23, 44)).BorderAround ColorIndex:=1, Weight:=xlThin
Cells(moduleStart + 23, 40).Value = leaseRate
 
' % of Air
Range(Cells(moduleStart + 23, 34), Cells(moduleStart + 23, 36)).Merge
Range(Cells(moduleStart + 23, 34), Cells(moduleStart + 23, 36)).BorderAround ColorIndex:=1, Weight:=xlThin
 
'Soft Costs Rate
Range(Cells(moduleStart + 25, 40), Cells(moduleStart + 25, 44)).Merge
Range(Cells(moduleStart + 25, 40), Cells(moduleStart + 25, 44)).BorderAround ColorIndex:=1, Weight:=xlThin
Cells(moduleStart + 25, 40).Value = leaseRate
 
'Lease Payment
Cells(moduleStart + 18, 49).Value = "Lease Payment"
lpFormula = "=T"
lpFormula = lpFormula + CStr(moduleStart + 19)
lpFormula = lpFormula + "*"
lpFormula = lpFormula + "AN"
lpFormula = lpFormula + CStr(moduleStart + 19)
Cells(moduleStart + 19, 49).Formula = lpFormula
Range(Cells(moduleStart + 19, 49), Cells(moduleStart + 19, 49)).NumberFormat = "$0.00"
 
'New Eqp Lease Payment
Range(Cells(moduleStart + 19, 49), Cells(moduleStart + 19, 56)).Merge
Range(Cells(moduleStart + 19, 49), Cells(moduleStart + 19, 56)).BorderAround ColorIndex:=1, Weight:=xlThin
 
'Used / Refinanced EQP Lease payment
usedFormula = "=T"
usedFormula = usedFormula + CStr(moduleStart + 21)
usedFormula = usedFormula + "*"
usedFormula = usedFormula + "AN"
usedFormula = usedFormula + CStr(moduleStart + 21)
Range(Cells(moduleStart + 21, 49), Cells(moduleStart + 21, 56)).Merge
Range(Cells(moduleStart + 21, 49), Cells(moduleStart + 21, 56)).BorderAround ColorIndex:=1, Weight:=xlThin
Cells(moduleStart + 21, 49).Formula = usedFormula
Range(Cells(moduleStart + 21, 49), Cells(moduleStart + 21, 49)).NumberFormat = "$0.00"
 
'Settlement Lease Payment
slpFormula = "=T"
slpFormula = slpFormula + CStr(moduleStart + 23)
slpFormula = slpFormula + "*"
slpFormula = slpFormula + "AN"
slpFormula = slpFormula + CStr(moduleStart + 23)
Range(Cells(moduleStart + 23, 49), Cells(moduleStart + 23, 56)).Merge
Range(Cells(moduleStart + 23, 49), Cells(moduleStart + 23, 56)).BorderAround ColorIndex:=1, Weight:=xlThin
Cells(moduleStart + 23, 49).Formula = slpFormula
Range(Cells(moduleStart + 23, 49), Cells(moduleStart + 23, 49)).NumberFormat = "$0.00"
 
'Soft Costs Lease Payment
sclpFormula = "=T"
sclpFormula = sclpFormula + CStr(moduleStart + 25)
sclpFormula = sclpFormula + "*"
sclpFormula = sclpFormula + "AN"
sclpFormula = sclpFormula + CStr(moduleStart + 25)
Range(Cells(moduleStart + 25, 49), Cells(moduleStart + 25, 56)).Merge
Range(Cells(moduleStart + 25, 49), Cells(moduleStart + 25, 56)).BorderAround ColorIndex:=1, Weight:=xlThin
Cells(moduleStart + 25, 49).Formula = sclpFormula
Range(Cells(moduleStart + 25, 49), Cells(moduleStart + 25, 49)).NumberFormat = "$0.00"
 
'Rate Total
Range(Cells(moduleStart + 31, 40), Cells(moduleStart + 31, 44)).Merge
Range(Cells(moduleStart + 31, 40), Cells(moduleStart + 31, 44)).BorderAround ColorIndex:=1, Weight:=xlThin
 
'Lease Payment
Range(Cells(moduleStart + 31, 49), Cells(moduleStart + 31, 56)).Merge
Range(Cells(moduleStart + 31, 49), Cells(moduleStart + 31, 56)).BorderAround ColorIndex:=1, Weight:=xlThin
 
'Total
Cells(moduleStart + 27, 2).Value = "Total"
Range(Cells(moduleStart + 27, 20), Cells(moduleStart + 27, 32)).Merge
Range(Cells(moduleStart + 27, 20), Cells(moduleStart + 27, 32)).BorderAround ColorIndex:=1, Weight:=xlThin
 
'Total Amount Financed
TAFormula = "=T"
TAFormula = TAFormula + CStr(moduleStart + 31)
TAFormula = TAFormula + "*"
TAFormula = TAFormula + "AN"
TAFormula = TAFormula + CStr(moduleStart + 31)
Cells(moduleStart + 31, 2).Value = "Total Amount Financed"
Cells(moduleStart + 31, 40).Value = leaseRate
Range(Cells(moduleStart + 31, 20), Cells(moduleStart + 31, 32)).Merge
Range(Cells(moduleStart + 31, 20), Cells(moduleStart + 31, 32)).BorderAround ColorIndex:=1, Weight:=xlThin
Cells(moduleStart + 31, 49).Formula = TAFormula
Range(Cells(moduleStart + 31, 49), Cells(moduleStart + 31, 49)).NumberFormat = "$0.00"
 
'Misc
Cells(moduleStart + 35, 58).Value = "*Please indicate any other pertinent information necessary"
Cells(moduleStart + 35, 58).HorizontalAlignment = xlRight
Range(Cells(moduleStart + 35, 1), Cells(moduleStart + 35, 58)).Interior.ColorIndex = 15
Cells(moduleStart + 35, 1).Value = "E) Miscellaneous"
Cells(moduleStart + 35, 1).Font.Bold = True
Cells(moduleStart + 35, 2).HorizontalAlignment = xlLeft
Range(Cells(moduleStart + 35, 1), Cells(moduleStart + 35, 58)).BorderAround ColorIndex:=1, Weight:=xlThin
Rows(moduleStart + 36).RowHeight = 6.6
 
'Misc Section
Range(Cells(moduleStart + 37, 1), Cells(moduleStart + 39, 58)).Merge
Range(Cells(moduleStart + 37, 1), Cells(moduleStart + 39, 58)).BorderAround ColorIndex:=1, Weight:=xlThin
Rows(moduleStart + 40).RowHeight = 6.6
 
 
'Footer thick side border
  With Worksheets(leaseSheet).Range(Cells(moduleStart - 1, 58), Cells(moduleStart + 40, 58)).Borders(xlEdgeRight)
       .LineStyle = xlContinuous
       .Weight = xlMedium
    End With
 
    With Worksheets(leaseSheet).Range(Cells(moduleStart + 40, 1), Cells(moduleStart + 40, 58)).Borders(xlEdgeBottom)
       .LineStyle = xlContinuous
       .Weight = xlMedium
    End With
 
 
' //////////////////////////////
' //Fill Movement Forms Script//
' //////////////////////////////
currentWs = 15
currentItem = 16
lpmItem = 16
standardSheetNumber = 15
mySheets = Worksheets.Count
StandardRowsPCList = 15
StandardRowsInSheet = 15
StandardRowsInColumnD = 31
equipmentSheet = 2
Sheets(equipmentSheet).Activate
productListEnd = Sheets(equipmentSheet).Range("B16").End(xlDown).Row - StandardRowsInSheet
itemsInSheet = 0
filledItems = 0
itemsFound = 0
qtyFound = 0
totalSettlement = 0

 
For i = standardSheetNumber To mySheets
   
    Sheets(currentWs).Activate
    Sheets(currentWs).Range(Cells(16, 4), Cells(45, 7)).ClearContents
    Sheets(currentWs).Range(Cells(16, 6), Cells(45, 7)).NumberFormat = accountingFormat
    itemsInSheet = Sheets(currentWs).Range("A16").End(xlDown).Row - StandardRowsInSheet
    thisConfigSettlement = 0
    firstProduct = 0
    
    For j = itemsFound To productListEnd - 1
   
        ProductCode = Sheets(equipmentSheet).Cells(currentItem + itemsFound, 5).Value
        productQty = Sheets(equipmentSheet).Cells(currentItem + itemsFound, 1).Value
        productCost = Sheets(equipmentSheet).Cells(currentItem + itemsFound, 10).Value
        mappCost = Sheets(equipmentSheet).Cells(currentItem + itemsFound, 12).Value
       
        For k = firstProduct To itemsInSheet
           
                pcToCheck = Sheets(currentWs).Cells(currentItem + k, 1).Value
 
                If pcToCheck = ProductCode Then
                   
                    Sheets(currentWs).Cells(currentItem + k, 4).Value = productQty
 
                    If pcToCheck = "SETRI" Then
                        thisConfigSettlement = productCost
                        totalSettlement = totalSettlement + thisConfigSettlement
                    End If
                    
                    Sheets(currentWs).Cells(currentItem + k, 6).Value = productCost * productQty
                    Sheets(currentWs).Cells(currentItem + k, 7).Value = mappCost * productQty

                    firstProduct = firstProduct + 1
                    itemsFound = itemsFound + 1

                    Exit For
                End If
        Next k
       
        filledItems = Sheets(currentWs).Range("D16").End(xlDown).Row - StandardRowsInSheet
       
        If filledItems = itemsInSheet Then
            totalPrice = WorksheetFunction.Sum(Sheets(currentWs).Range(Cells(16, 6), Cells(16 + itemsInSheet, 6)))
            totalPrice = totalPrice - thisConfigSettlement
            totalMapp = WorksheetFunction.Sum(Sheets(currentWs).Range(Cells(16, 7), Cells(16 + itemsInSheet, 7)))
           
            'Totals
            Sheets(lpmSheet).Activate
            Sheets(lpmSheet).Cells(lpmItem, 15).Value = totalPrice
            Sheets(lpmSheet).Cells(lpmItem, 30).Value = totalMapp
           
            Sheets(lpmSheet).Range(Cells(moduleStart, 30), Cells(moduleStart, 30)).NumberFormat = accountingFormat
            Sheets(lpmSheet).Range(Cells(moduleStart, 15), Cells(moduleStart, 15)).NumberFormat = accountingFormat
            Sheets(lpmSheet).Range(Cells(moduleStart, 45), Cells(moduleStart, 45)).NumberFormat = accountingFormat
           
            'Diff from MAPP
            diffFromMappFormula = "=AD" + CStr(lpmItem) + "-O" + CStr(lpmItem)
            diffMappPercentFormula = "=AD" + CStr(moduleStart) + "/AS" + CStr(moduleStart)
            Sheets(lpmSheet).Cells(lpmItem, 45).Formula = diffFromMappFormula
            Sheets(lpmSheet).Cells(moduleStart + 2, 46).Formula = diffMappPercentFormula
           
            lpmItem = lpmItem + 2
            Exit For
        End If
 
    Next j
   
    If currentWs < mySheets Then
        currentWs = currentWs + 1
    Else
        Sheets(lpmSheet).Activate
    End If
 
 
Next i
 
Cells(moduleStart + 11, 27).Value = totalSettlement

'Net Equipment Value
Cells(moduleStart + 10, 27).Value = Cells(moduleStart, 15).Value
 
'New Equipment Value
Cells(moduleStart + 19, 20).Value = Cells(moduleStart + 10, 27).Value
 
'Invoice price
formulaString = "=SUM(AA"
formulaString = formulaString + CStr(moduleStart + 4)
formulaString = formulaString + ":"
formulaString = formulaString + "AA"
formulaString = formulaString + CStr(moduleStart + 12)
formulaString = formulaString + ")"
formulaString = CStr(formulaString)
Range(Cells(moduleStart + 14, 27), Cells(moduleStart + 14, 27)).Formula = formulaString
 
'Settlement
settlementFormulastring = "=AA"
settlementFormulastring = settlementFormulastring + CStr(moduleStart + 11)
Range(Cells(moduleStart + 23, 20), Cells(moduleStart + 23, 20)).Formula = settlementFormulastring
Range(Cells(moduleStart + 23, 20), Cells(moduleStart + 23, 20)).NumberFormat = accountingFormat
 
'%of Air
If Cells(moduleStart + 23, 20).Value = 0 Then
    Cells(moduleStart + 23, 34).Value = 0
Else
    airFormula = "=T" + CStr(moduleStart + 23) + "/AA" + CStr(moduleStart + 14)
    Air = Cells(moduleStart + 23, 20).Value / Cells(moduleStart + 14, 27).Value
    Cells(moduleStart + 23, 34).Formula = airFormula
End If
 
'Number Formats
Range(Cells(moduleStart + 23, 34), Cells(moduleStart + 23, 34)).NumberFormat = percentageFormat
Range(Cells(moduleStart + 4, 27), Cells(moduleStart + 14, 27)).NumberFormat = accountingFormat
 
'Soft Costs
softCostFormula = "=Sum(AA" + CStr(moduleStart + 4) + ":AA" + CStr(moduleStart + 9) + ")" + "+AA" + CStr(moduleStart + 12)
Cells(moduleStart + 25, 20).Formula = softCostFormula
 
'Total Costs
totalCostFormula = "=Sum(T" + CStr(moduleStart + 19) + ":" + "T" + CStr(moduleStart + 25) + ")"
Cells(moduleStart + 27, 20).Formula = totalCostFormula
 
'Total Amount Financed
Cells(moduleStart + 31, 20).Formula = totalCostFormula
 
'Lock cells and protect sheet
Range(Cells(moduleStart, 1), Cells(moduleStart + 30, 58)).Locked = False
Range(Cells(1, 1), Cells(moduleStart - 1, 58)).Locked = True
ActiveSheet.Protect Password:="sherpadoc1"
 
End Sub


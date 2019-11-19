Attribute VB_Name = "Module3"
Sub BoS20_Button1_Click()

bosSheet = "BoS 2.0"
equipmentSheet = "Equip. Info-DO NOT DELETE"
lpmSheet = "Lease Price Model 2.0"
Sheets(bosSheet).Activate

'Align columns
Columns("A").ColumnWidth = 1.67
Columns("B").ColumnWidth = 3.11
Columns("C").ColumnWidth = 13.67
Columns("D").ColumnWidth = 6.11
Columns("E").ColumnWidth = 32.89
Columns("F").ColumnWidth = 6.56
Columns("G").ColumnWidth = 15.33
Columns("H").ColumnWidth = 11.78
Columns("I").ColumnWidth = 1.56
 
'Align Header Rows
Rows(1).RowHeight = 11.4
Rows(2).RowHeight = 10.8
Rows(3).RowHeight = 19.2
Rows(4).RowHeight = 21
Rows(5).RowHeight = 14.4
Rows(6).RowHeight = 14.4
Rows(7).RowHeight = 12
Rows(8).RowHeight = 12
Rows(9).RowHeight = 12
Rows(10).RowHeight = 12
Rows(11).RowHeight = 12
Rows(12).RowHeight = 27
Rows(13).RowHeight = 14.4
Rows(14).RowHeight = 14.4
 
'Font Settings
Range(Cells(12, 1), Cells(moduleStart + 100, 9)).Font.Name = "Arial"
Range(Cells(12, 1), Cells(moduleStart + 100, 9)).Font.Size = 8
 
'Get Account info for Header
acctName = Sheets(1).Cells(21, 2).Value
acctAddy = Sheets(1).Cells(22, 2).Value
acctCity = Sheets(1).Cells(24, 2).Value
acctProv = Sheets(1).Cells(26, 2).Value
acctPostal = Sheets(1).Cells(27, 2).Value
acctContact = Sheets(1).Cells(30, 4).Value
acctPhone = Sheets(1).Cells(28, 4).Value
acctFax = Sheets(1).Cells(29, 4).Value
acctEmail = Sheets(1).Cells(31, 4).Value
acctRep = Sheets(1).Cells(12, 2).Value
acctBilling = Sheets(1).Cells(22, 4).Value
acctPO = Sheets(1).Cells(18, 2).Value
 
'Inject account info into Header
 
 
If acctRep = "" Then
    Range(Cells(5, 4), Cells(5, 4)).Interior.ColorIndex = 6
Else
    Cells(5, 4).Value = acctRep
End If
 
 
If acctName = "" Then
    Range(Cells(7, 4), Cells(7, 4)).Interior.ColorIndex = 6
Else
    Cells(7, 4).Value = acctName
End If
 
 
If acctBilling = "" Then
    Range(Cells(9, 4), Cells(9, 4)).Interior.ColorIndex = 6
Else
    Cells(9, 4).Value = acctBilling
End If
 
 
If acctPO = "" Then
    Range(Cells(5, 7), Cells(5, 7)).Interior.ColorIndex = 6
Else
    Cells(5, 7).Value = acctPO
End If
 
 
If acctFax = "" Then
    Range(Cells(8, 7), Cells(8, 7)).Interior.ColorIndex = 6
Else
    Cells(8, 7).Value = acctFax
End If
 
 
If acctContact = "" Then
    Range(Cells(8, 4), Cells(8, 4)).Interior.ColorIndex = 6
    Range(Cells(6, 7), Cells(6, 7)).Interior.ColorIndex = 6
Else
    Cells(8, 4).Value = acctContact
    Cells(6, 7).Value = acctContact
End If
 
 
If acctPhone = "" Then
    Range(Cells(7, 7), Cells(7, 7)).Interior.ColorIndex = 6
Else
    Cells(7, 7).Value = acctPhone
End If
 
 
If acctEmail = "" Then
    Range(Cells(9, 7), Cells(9, 7)).Interior.ColorIndex = 6
Else
    Cells(9, 7).Value = acctEmail
End If
 
 
 
Cells(4, 7).Value = Date
Range(Cells(4, 7), Cells(4, 7)).NumberFormat = "mmm dd, yyyy"
Range(Cells(4, 7), Cells(4, 7)).HorizontalAlignment = xlCenter
 
 
Dim thisModel As String
Dim lastModel As String
Dim lastLocation As String
Dim thisTax As Double
Dim onTax As Double
Dim bcTax As Double
Dim mbTax As Double
Dim nfTax As Double
Dim ntTax As Double
Dim nsTax As Double
Dim nuTax As Double
Dim peTax As Double
Dim qcTax As Double
Dim skTax As Double
Dim ykTax As Double
Dim appText As String
 
appText = "APPLICATION: " & Chr(10)
appText = appText & "You agree to purchase the equipment, software licenses and/or software maintenance and support products listed above in accordance with the terms and conditions stated above and set out overleaf. You agree to pay to us the payments set forth above. You agree that all information set out herein is correct and that all particulars were complete when this Agreement was signed by you. You acknowledge having read the terms and conditions of this Agreement set forth on this page and overleaf, and agree that no other terms and conditions, express or implied, are part of this Agreement unless they appear above or in a schedule or addendum, and in either event are initialed by both of us to indicate they form part of this Agreement." & Chr(10)
appText = appText & "RETURNS:" & Chr(10)
appText = appText & "No equipment, software or software maintenance and support products may be returned without Document Direction Limited’s prior written consent, which may be withheld at our sole discretion. Merchandise returned without written authorization may not be accepted at the receiving dock and remains your sole responsibility. On returns authorized by us in advance, you agree to pay a restocking charge determined by us. All claims for damaged equipment shall be deemed waived unless made in writing and delivered to us within five days after your receipt of the applicable equipment" & Chr(10)
appText = appText & "PAYMENT:" & Chr(10)
appText = appText & "The purchase price for equipment and software is invoiced upon Document Direction Limited’s shipment of the applicable product to you, notwithstanding date of installation. Software is deemed to be shipped at the time the applicable license key is downloaded. Software maintenance and support is invoiced immediately following execution of this Agreement. You agree to pay invoices in accordance with their terms." & Chr(10)
appText = appText & "SIGNATURE:"
 
'Sales tax Rates
abTax = 0.05
bcTax = 0.12
mbTax = 0.12
nbTax = 0.15
nfTax = 0.15
ntTax = 0.05
nsTax = 0.15
nuTax = 0.05
onTax = 0.13
peTax = 0.13
qcTax = 0.14975
skTax = 0.11
ykTax = 0.05




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
equipmentSheet = 2
Sheets(equipmentSheet).Activate
productListEnd = Sheets(equipmentSheet).Range("B16").End(xlDown).Row - StandardRowsPCList
itemsInSheet = 0
filledItems = 0
itemsFound = 0
bosSheet = 7
qtyFound = 0
totalDealSettlement = 0

For i = standardSheetNumber To mySheets
   
    Sheets(currentWs).Activate
    Sheets(currentWs).Range(Cells(16, 4), Cells(45, 7)).ClearContents
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
                    End If
                    
                    Sheets(currentWs).Cells(currentItem + k, 6).Value = productCost * productQty
                    Sheets(currentWs).Cells(currentItem + k, 7).Value = mappCost * productQty
                    
                    itemsFound = itemsFound + 1
                    firstProduct = firstProduct + 1
                                   
                    Exit For
                End If
        Next k
 
    Next j
   
    If currentWs < mySheets Then
        currentWs = currentWs + 1
    Else
        Sheets(bosSheet).Activate
    End If
 
 
Next i


' /////////////////////////////////////
' //End of Fill Movement Forms Script//
' /////////////////////////////////////

 
machineIndex = 16
modelQty = Sheets(2).Cells(machineIndex, 27).Value
moduleStart = 13
modelsToCheckStart = 33
standardSheetNumber = 15
mySheets = Worksheets.Count
statndardRowsInModelDesc = 15
modelTypes = Sheets(2).Cells(Rows.Count, "AB").End(xlUp).Row - statndardRowsInModelDesc
 
 
For i = standardSheetNumber To mySheets
       
    Sheets(i).Activate
    'Get
    thisLocation = Sheets(i).Cells(8, 2).Value
    thisProv = Sheets(i).Cells(10, 2).Value
    thisModel = Sheets(i).Cells(16, 2).Value
    thisConfigPrice = WorksheetFunction.Sum(Sheets(i).Range(Cells(16, 6), Cells(45, 6)))
    Sheets(bosSheet).Activate
   
        'Set Quantity
      Cells(moduleStart, 2).Value = 1
      
      'Set Model
      Cells(moduleStart, 3).Value = thisModel
   
        'Set Location
      Cells(moduleStart, 5).Value = thisLocation & " - " & thisProv
       
        'Set Province
      Cells(moduleStart, 6).Value = thisProv
 
        'Set Tax
      Select Case thisProv
            Case "ON"
                thisTax = onTax
            Case "BC"
                thisTax = bcTax
            Case "MB"
                thisTax = mbTax
            Case "NB"
                thisTax = nbTax
            Case "NF"
                thisTax = nfTax
            Case "NT"
                thisTax = ntTax
            Case "NS"
                thisTax = nsTax
            Case "NU"
                thisTax = nuTax
            Case "PE"
                thisTax = peTax
            Case "AB"
                thisTax = abTax
            Case "QC"
                thisTax = qcTax
            Case "SK"
                thisTax = skTax
            Case "YK"
                thisTax = ykTax
        End Select
           
        'if quantity < 1 then add 1 to machineIndex
        If modelQty <= 1 Then
            machineIndex = machineIndex + 1
        Else
            modelQty = modelQty - 1
        End If
 
        'Increment line pointer
        moduleStart = moduleStart + 1
       
        lastLocation = thisLocation
        lastModel = thisModel

   
   
    For j = 0 To modelTypes - 1
        modelToCheck = Sheets(2).Cells(modelsToCheckStart + j, 18).Value
       
        If thisModel = modelToCheck Then
            Cells(moduleStart - 1, 7).Value = thisConfigPrice
            Range(Cells(moduleStart - 1, 8), Cells(moduleStart - 1, 8)).Formula = "=G" + CStr(moduleStart - 1) + "*" + CStr(thisTax)
            Exit For
        End If
    Next j
 
    'Multiply amount by quantity
    Cells(moduleStart - 1, 7).Value = Cells(moduleStart - 1, 7).Value * Cells(moduleStart - 1, 2).Value
    Cells(moduleStart - 1, 8).Value = Cells(moduleStart - 1, 8).Value * Cells(moduleStart - 1, 2).Value
   
    'Add in borders for layout
    Range(Cells(moduleStart - 1, 9), Cells(moduleStart - 1, 9)).Borders(xlEdgeRight).LineStyle = xlContinuous
    Range(Cells(moduleStart - 1, 8), Cells(moduleStart - 1, 8)).Borders(xlEdgeRight).LineStyle = xlContinuous
    Range(Cells(moduleStart - 1, 2), Cells(moduleStart - 1, 2)).Borders(xlEdgeLeft).LineStyle = xlContinuous
    Range("G13:H100").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
   
    'Align text
  Range(Cells(moduleStart - 1, 2), Cells(moduleStart - 1, 6)).HorizontalAlignment = xlCenter
   
Next i
 
'Insert Footer
Range(Cells(moduleStart - 1, 2), Cells(moduleStart - 1, 8)).Borders(xlEdgeBottom).LineStyle = xlContinuous
Range(Cells(moduleStart + 33, 1), Cells(moduleStart + 33, 9)).Borders(xlEdgeBottom).LineStyle = xlContinuous
Range(Cells(moduleStart - 1, 9), Cells(moduleStart + 33, 9)).Borders(xlEdgeRight).LineStyle = xlContinuous
Range(Cells(moduleStart - 5, 1), Cells(moduleStart + 33, 1)).Borders(xlEdgeLeft).LineStyle = xlContinuous
 
 
For i = 4 To 10
    Range(Cells(moduleStart + i, 6), Cells(moduleStart + i, 7)).Merge
Next i
 
Range(Cells(moduleStart + 4, 6), Cells(moduleStart + 10, 7)).HorizontalAlignment = xlRight
 
'Settlement Details
Cells(moduleStart + 3, 2).Value = "Settlement Details:"
Cells(moduleStart + 3, 2).Font.Bold = True
Range(Cells(moduleStart + 4, 2), Cells(moduleStart + 4, 5)).Merge
Range(Cells(moduleStart + 4, 2), Cells(moduleStart + 4, 5)).Borders(xlEdgeBottom).LineStyle = xlContinuous
Range(Cells(moduleStart + 5, 2), Cells(moduleStart + 5, 5)).Merge
Range(Cells(moduleStart + 5, 2), Cells(moduleStart + 5, 5)).Borders(xlEdgeBottom).LineStyle = xlContinuous
Range(Cells(moduleStart + 6, 2), Cells(moduleStart + 6, 5)).Merge
Range(Cells(moduleStart + 6, 2), Cells(moduleStart + 6, 5)).Borders(xlEdgeBottom).LineStyle = xlContinuous
Range(Cells(moduleStart + 7, 2), Cells(moduleStart + 7, 5)).Merge
Range(Cells(moduleStart + 7, 2), Cells(moduleStart + 7, 5)).Borders(xlEdgeBottom).LineStyle = xlContinuous
Range(Cells(moduleStart + 8, 2), Cells(moduleStart + 8, 5)).Merge
Range(Cells(moduleStart + 8, 2), Cells(moduleStart + 8, 5)).Borders(xlEdgeBottom).LineStyle = xlContinuous
Range(Cells(moduleStart + 9, 2), Cells(moduleStart + 9, 5)).Merge
Range(Cells(moduleStart + 9, 2), Cells(moduleStart + 9, 5)).Borders(xlEdgeBottom).LineStyle = xlContinuous
Range(Cells(moduleStart + 10, 2), Cells(moduleStart + 10, 5)).Merge
Rows(moduleStart + 4).RowHeight = 14.4
Rows(moduleStart + 5).RowHeight = 14.4
Rows(moduleStart + 6).RowHeight = 14.4
Rows(moduleStart + 7).RowHeight = 14.4
Rows(moduleStart + 8).RowHeight = 14.4
Rows(moduleStart + 9).RowHeight = 14.4
Rows(moduleStart + 10).RowHeight = 14.4
 
'Net Value before Tax
Cells(moduleStart + 4, 6).Value = "Net Value Before Tax:"
Cells(moduleStart + 4, 6).Font.Bold = True
Cells(moduleStart + 4, 8).Font.Italic = True
Cells(moduleStart + 4, 8).Borders(xlEdgeBottom).LineStyle = xlContinuous
Cells(moduleStart + 4, 8).Value = Application.Sum(Range(Cells(12, 7), Cells(moduleStart, 7)))
 
'Total Taxes
Cells(moduleStart + 6, 6).Value = "Total Taxes:"
Cells(moduleStart + 6, 6).Font.Bold = True
Cells(moduleStart + 6, 8).Font.Italic = True
Cells(moduleStart + 6, 8).Borders(xlEdgeBottom).LineStyle = xlContinuous
Cells(moduleStart + 6, 8).Value = Application.Sum(Range(Cells(12, 8), Cells(moduleStart, 8)))
 
'TOTAL
Cells(moduleStart + 8, 6).Value = "TOTAL:"
Cells(moduleStart + 8, 6).Font.Bold = True
Cells(moduleStart + 8, 8).Font.Italic = True
Cells(moduleStart + 8, 8).Borders(xlEdgeBottom).LineStyle = xlContinuous
Cells(moduleStart + 8, 8).Value = Application.Sum(Range(Cells(moduleStart + 4, 8), Cells(moduleStart + 6, 8)))
 
'Special Provisions
Cells(moduleStart + 11, 2).Value = "Special Provisions:"
Cells(moduleStart + 11, 2).VerticalAlignment = xlVAlignTop
Cells(moduleStart + 11, 2).Font.Size = 8
Rows(moduleStart + 11).RowHeight = 27
Rows(moduleStart + 12).RowHeight = 6
 
'Customer Initial
Cells(moduleStart + 11, 8).Value = "Customer" & Chr(10) & " Initial:"
Cells(moduleStart + 11, 8).VerticalAlignment = xlVAlignTop
Cells(moduleStart + 11, 8).Font.Size = 8
Range(Cells(moduleStart + 11, 2), Cells(moduleStart + 11, 8)).BorderAround ColorIndex:=1
Range(Cells(moduleStart + 11, 2), Cells(moduleStart + 11, 8)).IndentLevel = 0
Range(Cells(moduleStart + 11, 2), Cells(moduleStart + 11, 6)).Merge
 
'Application Text
Range(Cells(moduleStart + 12, 2), Cells(moduleStart + 25, 8)).Merge
Cells(moduleStart + 12, 2).Value = appText
Rows(moduleStart + 12).RowHeight = 10
Cells(moduleStart + 12, 2).Characters(1, 15).Font.Bold = True
Cells(moduleStart + 12, 2).Characters(15, 748).Font.Bold = False
Cells(moduleStart + 12, 2).Characters(749, 757).Font.Bold = True
Cells(moduleStart + 12, 2).Characters(758, 1334).Font.Bold = False
Cells(moduleStart + 12, 2).Characters(1335, 1345).Font.Bold = True
Cells(moduleStart + 12, 2).Characters(1346, 1759).Font.Bold = False
Cells(moduleStart + 12, 2).Characters(1760, 1780).Font.Bold = True
 
'Signature
Range(Cells(moduleStart + 26, 2), Cells(moduleStart + 26, 6)).Merge
Range(Cells(moduleStart + 26, 2), Cells(moduleStart + 27, 6)).HorizontalAlignment = xlCenter
Range(Cells(moduleStart + 26, 2), Cells(moduleStart + 26, 6)).BorderAround ColorIndex:=1
Cells(moduleStart + 26, 2).Value = "Signature(s) of Customer(s)"
 
'Acceptance
Range(Cells(moduleStart + 26, 7), Cells(moduleStart + 26, 8)).Merge
Range(Cells(moduleStart + 26, 7), Cells(moduleStart + 26, 8)).BorderAround ColorIndex:=1
Cells(moduleStart + 26, 7).Value = "Acceptance by Document Direction Limited"
Cells(moduleStart + 26, 7).Font.Size = 7
Rows(moduleStart + 26).RowHeight = 14
 
'Signature (2)
Range(Cells(moduleStart + 27, 2), Cells(moduleStart + 27, 4)).Merge
Range(Cells(moduleStart + 27, 2), Cells(moduleStart + 27, 4)).BorderAround ColorIndex:=1
Range(Cells(moduleStart + 27, 2), Cells(moduleStart + 27, 8)).Font.Size = 8
Cells(moduleStart + 27, 2).Value = "Signature"
 
'Print name and Position
Cells(moduleStart + 27, 5).Value = "Print Name and Position"
Cells(moduleStart + 27, 5).BorderAround ColorIndex:=1
 
'Date Signed
Cells(moduleStart + 27, 6).Value = "Date" & Chr(10) & "Signed"
Cells(moduleStart + 27, 6).VerticalAlignment = xlCenter
Cells(moduleStart + 27, 6).BorderAround ColorIndex:=1
 
'DDL Signature
Cells(moduleStart + 27, 7).Value = "Signature of Document Direction Limited"
Range(Cells(moduleStart + 27, 7), Cells(moduleStart + 27, 8)).Merge
Range(Cells(moduleStart + 27, 7), Cells(moduleStart + 27, 8)).BorderAround ColorIndex:=1
 
'bind to customer
Range(Cells(moduleStart + 28, 2), Cells(moduleStart + 32, 4)).Merge
Range(Cells(moduleStart + 28, 2), Cells(moduleStart + 32, 4)).WrapText = True
Range(Cells(moduleStart + 28, 2), Cells(moduleStart + 32, 4)).BorderAround ColorIndex:=1
Cells(moduleStart + 28, 2).Value = "I have the authority to bind the Customer"
Cells(moduleStart + 28, 2).HorizontalAlignment = xlCenter
Cells(moduleStart + 28, 2).VerticalAlignment = xlBottom
 
'Name Slot
Range(Cells(moduleStart + 28, 5), Cells(moduleStart + 32, 5)).Merge
Range(Cells(moduleStart + 28, 5), Cells(moduleStart + 32, 5)).BorderAround ColorIndex:=1
 
'Date Signed Slot
Range(Cells(moduleStart + 28, 6), Cells(moduleStart + 32, 6)).Merge
Range(Cells(moduleStart + 28, 6), Cells(moduleStart + 32, 6)).BorderAround ColorIndex:=1
 
'Signature Slot
Range(Cells(moduleStart + 28, 7), Cells(moduleStart + 32, 8)).Merge
Range(Cells(moduleStart + 28, 7), Cells(moduleStart + 32, 8)).BorderAround ColorIndex:=1




End Sub


Attribute VB_Name = "Module1"
Sub Button2_Click()

serviceAgreementSheet = "Service Contract 2.0"
Sheets(serviceAgreementSheet).Activate

moduleStart = 12
standardSheetNumber = 15
mySheets = Worksheets.Count
 
'Get account details
accountName = Sheets(1).Cells(21, 2).Value
accountNumber = Sheets(1).Cells(17, 2).Value
accountAddress = Sheets(1).Cells(22, 2).Value
accountCity = Sheets(1).Cells(24, 2).Value
accountState = Sheets(1).Cells(26, 2).Value
accountZip = Sheets(1).Cells(27, 2).Value
accountContact = Sheets(1).Cells(30, 4).Value
accountPhone = Sheets(1).Cells(28, 4).Value
accountFax = Sheets(1).Cells(29, 4).Value
accountEmail = Sheets(1).Cells(31, 4).Value
accountRep = Sheets(1).Cells(12, 2).Value
machine = 16
machineQty = Sheets(2).Cells(machine, 27).Value
 
 
'App Text
appText = "You apply to us to service the equipment listed above to you for the Initial Period referred to above and thereafter in accordance with the terms and conditions stated above and set out overleaf. You agree to pay to us the payments set forth above in accordance with the frequency set out above. You agree that all information set out herein is correct and that all particulars were complete when this application was signed. You acknowledge having read the terms and conditions of this Agreement set forth on this page and overleaf, and agree that no other terms and conditions, express or implied, are part of this agreement unless they appear above or in a schedule or addendum, and in either event are initialed by both of us to indicate they form part of this Agreement"
 
 
'Fill out Header
Range(Cells(6, 2), Cells(6, 5)).Merge
Range(Cells(6, 2), Cells(6, 5)).Borders(xlEdgeBottom).LineStyle = xlDash
Range(Cells(7, 2), Cells(7, 5)).Merge
Range(Cells(7, 2), Cells(7, 5)).Borders(xlEdgeBottom).LineStyle = xlDash
Range(Cells(7, 7), Cells(8, 9)).Merge
Range(Cells(8, 2), Cells(8, 5)).Merge
Range(Cells(8, 2), Cells(8, 5)).Borders(xlEdgeBottom).LineStyle = xlDash
Range(Cells(9, 2), Cells(9, 5)).Merge
Range(Cells(9, 2), Cells(9, 5)).Borders(xlEdgeBottom).LineStyle = xlDash
Range(Cells(10, 2), Cells(10, 5)).Merge
Range(Cells(10, 2), Cells(10, 5)).Borders(xlEdgeBottom).LineStyle = xlDash
Range(Cells(11, 2), Cells(11, 5)).Merge
Range(Cells(11, 2), Cells(11, 5)).Borders(xlEdgeBottom).LineStyle = xlDash
Range(Cells(12, 2), Cells(12, 5)).Merge
Range(Cells(12, 2), Cells(12, 5)).Borders(xlEdgeBottom).LineStyle = xlDash
 

Cells(6, 2).Value = accountName
Cells(7, 2).Value = accountAddress
Cells(8, 2).Value = accountCity & " - " & accountState & " - " & accountZip
Cells(9, 2).Value = accountContact
Cells(10, 2).Value = accountPhone

If accountNumber = 0 Then
    Range(Cells(6, 7), Cells(6, 7)).Interior.ColorIndex = 6
Else
    Cells(6, 7).Value = accountNumber
End If
 
If (accountFax = "") Then
    Range(Cells(11, 2), Cells(11, 2)).Interior.ColorIndex = 6
Else
    Cells(11, 2).Value = accountFax
End If

If accountEmail = "" Then
    Range(Cells(12, 2), Cells(12, 2)).Interior.ColorIndex = 6
    Range(Cells(7, 7), Cells(7, 7)).Interior.ColorIndex = 6
Else
    Cells(12, 2).Value = accountEmail
    Cells(7, 7).Value = accountEmail
End If
 
If accountRep = "" Then
    Range(Cells(10, 7), Cells(10, 7)).Interior.ColorIndex = 6
Else
    Cells(10, 7).Value = accountRep
End If
 
 
'Create Layout
For i = standardSheetNumber To mySheets
 
Range(Cells(moduleStart + 1, 9), Cells(moduleStart + 20, 9)).RowHeight = 10.2
Range(Cells(moduleStart + 1, 1), Cells(moduleStart + 20, 9)).Font.Bold = False
Range(Cells(moduleStart + 1, 9), Cells(moduleStart + 20, 9)).Borders(xlEdgeRight).LineStyle = xlContinuous
 
Cells(moduleStart + 2, 1).Value = "Contract Type: CPC"
Range(Cells(moduleStart + 2, 1), Cells(moduleStart + 2, 9)).Merge
Range(Cells(moduleStart + 2, 1), Cells(moduleStart + 2, 9)).Font.Bold = True
 
Cells(moduleStart + 3, 1).Value = "Model"
Range(Cells(moduleStart + 3, 1), Cells(moduleStart + 3, 3)).Merge
Range(Cells(moduleStart + 3, 1), Cells(moduleStart + 3, 3)).Font.Bold = True
 
Cells(moduleStart + 5, 1).Value = "Additional Fixed Charge"
Range(Cells(moduleStart + 5, 1), Cells(moduleStart + 5, 3)).Merge
Range(Cells(moduleStart + 5, 1), Cells(moduleStart + 5, 3)).Font.Bold = True
 
Cells(moduleStart + 6, 1).Value = "NO"
Cells(moduleStart + 7, 1).Value = "Fixed Charge Description (if applicable):"
Range(Cells(moduleStart + 7, 1), Cells(moduleStart + 7, 9)).Merge
 
'Black
Cells(moduleStart + 9, 1).Value = "Blk"
Range(Cells(moduleStart + 5, 1), Cells(moduleStart + 5, 3)).Font.Bold = True
Range(Cells(moduleStart + 8, 1), Cells(moduleStart + 10, 9)).Borders.LineStyle = xlContinuous
 
'Color
Cells(moduleStart + 10, 1).Value = "Clr"
Cells(moduleStart + 12, 1).Value = "SHIP TO:"
Range(Cells(moduleStart + 12, 1), Cells(moduleStart + 12, 9)).Merge
Cells(moduleStart + 12, 1).Font.Bold = True
 
Cells(moduleStart + 13, 1).Value = "Account #:"
 
Cells(moduleStart + 14, 1).Value = "Name"
Cells(moduleStart + 15, 1).Value = "Address"
 
Cells(moduleStart + 19, 1).Value = "Special Provisions:"
Range(Cells(moduleStart + 19, 1), Cells(moduleStart + 20, 8)).Merge
Range(Cells(moduleStart + 19, 1), Cells(moduleStart + 20, 8)).VerticalAlignment = xlVAlignTop
Range(Cells(moduleStart + 19, 1), Cells(moduleStart + 20, 8)).Borders.LineStyle = xlContinuous
 
Cells(moduleStart + 6, 2).Value = "YES"
 
'Image Charge
Cells(moduleStart + 8, 2).Value = "Image Charge"
Cells(moduleStart + 9, 2).Value = Sheets(2).Cells(machine, 37).Value
Cells(moduleStart + 10, 2).Value = Sheets(2).Cells(machine, 38).Value
Range(Cells(moduleStart + 8, 2), Cells(moduleStart + 8, 9)).Font.Bold = True
Range(Cells(moduleStart + 8, 2), Cells(moduleStart + 8, 4)).HorizontalAlignment = xlCenter
Range(Cells(moduleStart + 8, 2), Cells(moduleStart + 8, 4)).Merge
 
Range(Cells(moduleStart + 9, 2), Cells(moduleStart + 9, 4)).HorizontalAlignment = xlCenter
Range(Cells(moduleStart + 9, 2), Cells(moduleStart + 9, 4)).Merge
 
Range(Cells(moduleStart + 10, 2), Cells(moduleStart + 10, 4)).HorizontalAlignment = xlCenter
Range(Cells(moduleStart + 10, 2), Cells(moduleStart + 10, 4)).Merge
 
Cells(moduleStart + 3, 4).Value = "Serial#"
Range(Cells(moduleStart + 4, 4), Cells(moduleStart + 4, 5)).Interior.ColorIndex = 6
Range(Cells(moduleStart + 3, 4), Cells(moduleStart + 3, 5)).Merge
Range(Cells(moduleStart + 3, 4), Cells(moduleStart + 3, 9)).Font.Bold = True
 
Cells(moduleStart + 5, 4).Value = "Fix Change Amount"
Range(Cells(moduleStart + 5, 4), Cells(moduleStart + 5, 9)).Font.Bold = True
Range(Cells(moduleStart + 6, 4), Cells(moduleStart + 6, 5)).Merge
Range(Cells(moduleStart + 6, 4), Cells(moduleStart + 6, 8)).Interior.ColorIndex = 6
 
'Agreed Volume
Cells(moduleStart + 8, 5).Value = "Agreed Volume"
Cells(moduleStart + 9, 5).Value = Sheets(2).Cells(machine, 39).Value
Cells(moduleStart + 10, 5).Value = Sheets(2).Cells(machine, 40).Value
Range(Cells(moduleStart + 8, 5), Cells(moduleStart + 8, 5)).HorizontalAlignment = xlCenter
Range(Cells(moduleStart + 8, 5), Cells(moduleStart + 8, 7)).Merge
 
Range(Cells(moduleStart + 9, 5), Cells(moduleStart + 9, 7)).HorizontalAlignment = xlCenter
Range(Cells(moduleStart + 9, 5), Cells(moduleStart + 9, 7)).Merge
 
Range(Cells(moduleStart + 10, 5), Cells(moduleStart + 10, 5)).HorizontalAlignment = xlCenter
Range(Cells(moduleStart + 10, 5), Cells(moduleStart + 10, 7)).Merge
 
'Installed date
Cells(moduleStart + 3, 6).Value = "Installed Date"
Range(Cells(moduleStart + 3, 6), Cells(moduleStart + 3, 8)).Merge
Range(Cells(moduleStart + 4, 6), Cells(moduleStart + 4, 8)).Merge
 
Cells(moduleStart + 5, 6).Value = "Billing Frequency"
Range(Cells(moduleStart + 5, 6), Cells(moduleStart + 5, 8)).Merge
Range(Cells(moduleStart + 6, 6), Cells(moduleStart + 6, 8)).Merge
 
Cells(moduleStart + 13, 6).Value = "Meter Read:"
Range(Cells(moduleStart + 13, 7), Cells(moduleStart + 13, 9)).Merge
Range(Cells(moduleStart + 13, 7), Cells(moduleStart + 13, 9)).Borders(xlEdgeBottom).LineStyle = xlDash
Range(Cells(moduleStart + 13, 7), Cells(moduleStart + 13, 8)).Interior.ColorIndex = 6
 
Cells(moduleStart + 14, 6).Value = "Phone #:"
Range(Cells(moduleStart + 14, 7), Cells(moduleStart + 14, 9)).Merge
Range(Cells(moduleStart + 14, 7), Cells(moduleStart + 14, 9)).Borders(xlEdgeBottom).LineStyle = xlDash
 
Cells(moduleStart + 15, 6).Value = "Fax #:"
Range(Cells(moduleStart + 15, 7), Cells(moduleStart + 15, 9)).Merge
Range(Cells(moduleStart + 15, 7), Cells(moduleStart + 15, 9)).Borders(xlEdgeBottom).LineStyle = xlDash
 
Cells(moduleStart + 16, 6).Value = "Email:"
Range(Cells(moduleStart + 16, 7), Cells(moduleStart + 16, 9)).Merge
Range(Cells(moduleStart + 16, 7), Cells(moduleStart + 16, 9)).Borders(xlEdgeBottom).LineStyle = xlDash
 
'Cells at bottom of Module
Range(Cells(moduleStart + 17, 7), Cells(moduleStart + 17, 9)).Merge
Range(Cells(moduleStart + 18, 1), Cells(moduleStart + 18, 9)).Merge
 
Cells(moduleStart + 8, 8).Value = "Meter Start"
Range(Cells(moduleStart + 8, 8), Cells(moduleStart + 8, 8)).HorizontalAlignment = xlCenter
Range(Cells(moduleStart + 8, 8), Cells(moduleStart + 8, 9)).Merge
 
'under meter start
Range(Cells(moduleStart + 9, 8), Cells(moduleStart + 9, 8)).HorizontalAlignment = xlCenter
Range(Cells(moduleStart + 9, 8), Cells(moduleStart + 9, 9)).Merge
 
'second cell under meter start
Range(Cells(moduleStart + 10, 8), Cells(moduleStart + 10, 8)).HorizontalAlignment = xlCenter
Range(Cells(moduleStart + 10, 8), Cells(moduleStart + 10, 9)).Merge
 
Cells(moduleStart + 3, 9).Value = "Service Fee"
 
Cells(moduleStart + 4, 9).Interior.ColorIndex = 6
 
Cells(moduleStart + 5, 9).Value = "Initial Period (Mths)"
 
Cells(moduleStart + 6, 9).HorizontalAlignment = xlLeft
 
'line break the customer and initial
Cells(moduleStart + 19, 9).Value = "Customer " & Chr(10) & "Initial:"
Range(Cells(moduleStart + 19, 9), Cells(moduleStart + 20, 9)).Merge
Range(Cells(moduleStart + 19, 9), Cells(moduleStart + 20, 9)).Borders.LineStyle = xlContinuous
 
 
    'Get Sheet name
  myActiveSheet = Sheets(i).Name
 
    'Get Model
  Range(Cells(moduleStart + 4, 1), Cells(moduleStart + 4, 3)).Merge
    Model = Sheets(myActiveSheet).Cells(16, 2).Value
    Cells(moduleStart + 4, 1).Value = Model
   
    'insert Account number
  Range(Cells(moduleStart + 13, 2), Cells(moduleStart + 13, 5)).Merge
    Range(Cells(moduleStart + 13, 2), Cells(moduleStart + 13, 5)).Borders(xlEdgeBottom).LineStyle = xlDash
    Cells(moduleStart + 13, 2).Value = accountNumber
    Cells(moduleStart + 13, 2).HorizontalAlignment = xlLeft
   
    'Get CompanyName from sheet
  Range(Cells(moduleStart + 14, 2), Cells(moduleStart + 14, 5)).Merge
    Range(Cells(moduleStart + 14, 2), Cells(moduleStart + 14, 5)).Borders(xlEdgeBottom).LineStyle = xlDash
    CompanyName = Sheets(myActiveSheet).Cells(7, 2).Value
    Cells(moduleStart + 14, 2).Value = CompanyName
   
    'Get Address from sheet
  Range(Cells(moduleStart + 15, 2), Cells(moduleStart + 15, 5)).Merge
    Range(Cells(moduleStart + 15, 2), Cells(moduleStart + 15, 5)).Borders(xlEdgeBottom).LineStyle = xlDash
    Address = Sheets(myActiveSheet).Cells(8, 2).Value
    Cells(moduleStart + 15, 2).Value = Address
   
    'Get City from sheet
  Range(Cells(moduleStart + 16, 2), Cells(moduleStart + 16, 5)).Merge
    Range(Cells(moduleStart + 16, 2), Cells(moduleStart + 16, 5)).Borders(xlEdgeBottom).LineStyle = xlDash
    City = Sheets(myActiveSheet).Cells(9, 2).Value
    Cells(moduleStart + 16, 2).Value = City
   
    'Get State from sheet
  State = Sheets(myActiveSheet).Cells(10, 2).Value
   
    'Get ZipCode from sheet
  Range(Cells(moduleStart + 17, 2), Cells(moduleStart + 17, 5)).Merge
    Range(Cells(moduleStart + 17, 2), Cells(moduleStart + 17, 5)).Borders(xlEdgeBottom).LineStyle = xlDash
    ZipCode = Sheets(myActiveSheet).Cells(11, 2).Value
    Cells(moduleStart + 17, 2).Value = State & " - " & ZipCode
   
    Cells(moduleStart + 6, 9).Value = 12
   
    moduleStart = moduleStart + 20
   
    If machineQty <= 1 Then
        machine = machine + 1
    End If
   
    machineQty = machineQty - 1
   
    If machineQty = 0 Then
        machineQty = Sheets(2).Cells(machine, 25).Value
    End If
   
   
Next i
 
'Insert Footer
 
'Application
Range(Cells(moduleStart + 1, 1), Cells(moduleStart + 1, 9)).Merge
Cells(moduleStart + 1, 1).Value = "APPLICATION:" & Chr(10) & appText
Cells(moduleStart + 1, 1).Font.Size = 6.5
Cells(moduleStart + 1, 1).Font.Bold = False
Cells(moduleStart + 1, 1).Characters(1, 12).Font.Bold = True
Cells(moduleStart + 1, 1).Characters(1, 12).Font.Size = 7.5
Cells(moduleStart + 1, 1).RowHeight = 72.6
Cells(moduleStart + 1, 1).WrapText = True
Cells(moduleStart + 1, 1).IndentLevel = 1
Range(Cells(moduleStart + 1, 1), Cells(moduleStart + 1, 9)).VerticalAlignment = xlCenter
Range(Cells(moduleStart + 1, 9), Cells(moduleStart + 6, 9)).Borders(xlEdgeRight).LineStyle = xlContinuous
 
'Signature
Range(Cells(moduleStart + 2, 1), Cells(moduleStart + 2, 9)).Merge
Cells(moduleStart + 2, 1).Value = "SIGNATURE"
Cells(moduleStart + 2, 1).Font.Bold = True
Cells(moduleStart + 2, 1).IndentLevel = 1
Cells(moduleStart + 2, 1).RowHeight = 14.4
Range(Cells(moduleStart + 3, 1), Cells(moduleStart + 5, 9)).Borders.LineStyle = xlContinuous
 
'Signatures of Customers
Range(Cells(moduleStart + 3, 1), Cells(moduleStart + 3, 8)).Merge
Cells(moduleStart + 3, 1).Value = "Signaure(s) of Customer(s)"
Cells(moduleStart + 3, 1).IndentLevel = 33
Cells(moduleStart + 3, 1).Font.Bold = False
 
'Acceptance
Cells(moduleStart + 3, 9).Value = "Acceptance by Document Direction Limited"
Cells(moduleStart + 3, 9).Font.Size = 6.5
Cells(moduleStart + 3, 9).Font.Bold = False
 
'SignatureField2
Range(Cells(moduleStart + 4, 1), Cells(moduleStart + 4, 3)).Merge
Cells(moduleStart + 4, 1).Value = "Signature"
Cells(moduleStart + 4, 1).Font.Size = 7.5
Cells(moduleStart + 4, 1).Font.Bold = False
Cells(moduleStart + 4, 1).HorizontalAlignment = xlCenter
 
'Print name and Position
Range(Cells(moduleStart + 4, 4), Cells(moduleStart + 4, 6)).Merge
Cells(moduleStart + 4, 4).Value = "Print name and Position"
Cells(moduleStart + 4, 4).Font.Size = 7.5
Cells(moduleStart + 4, 4).Font.Bold = False
Cells(moduleStart + 4, 4).HorizontalAlignment = xlCenter
 
'Date Signed
Range(Cells(moduleStart + 4, 7), Cells(moduleStart + 4, 8)).Merge
Cells(moduleStart + 4, 7).Value = "Date Signed"
Cells(moduleStart + 4, 7).Font.Size = 7.5
Cells(moduleStart + 4, 7).Font.Bold = False
Cells(moduleStart + 4, 7).HorizontalAlignment = xlCenter
 
'Signature of Document Direction Limited
Cells(moduleStart + 4, 9).Value = "Signature of Document Direction Limited"
Cells(moduleStart + 4, 9).Font.Size = 7.5
Cells(moduleStart + 4, 9).Font.Bold = False
Cells(moduleStart + 4, 7).HorizontalAlignment = xlCenter
 
'Signature Field
Range(Cells(moduleStart + 5, 1), Cells(moduleStart + 5, 3)).Merge
Cells(moduleStart + 5, 1).RowHeight = 46.8
 
'Print name and Position
Range(Cells(moduleStart + 5, 4), Cells(moduleStart + 5, 6)).Merge
 
'Date signed
Range(Cells(moduleStart + 5, 7), Cells(moduleStart + 5, 8)).Merge
 
'Service start date
Cells(moduleStart + 5, 9).Font.Bold = False
Cells(moduleStart + 5, 9).Value = "Service Start Date:"
Cells(moduleStart + 5, 9).VerticalAlignment = xlBottom

'Erase line
Range(Cells(moduleStart + 6, 9), Cells(moduleStart + 6, 9)).Borders.LineStyle = xlNone

End Sub

Attribute VB_Name = "Module2"
Sub Button1_Click()
leaseSheet = "Lease Agreement 2.0"
Sheets(leaseSheet).Activate

Dim LA_model As String
Dim lastStreet As String
Dim lastCity As String
Dim lastProv As String
Dim LA_prevModel As String
Dim LA_lastAddy As String
Dim appText As String
Dim accepText As String
Dim finalLine As String
 
LA_moduleStart = 16
LA_machineIndex = 0
LA_standardSheetNumber = 15
LA_mySheets = Worksheets.Count
LA_modelQty = Sheets(2).Cells(LA_machineIndex + 16, 1).Value
 
appText = Chr(10) & "APPLICATION:" & Chr(10) & "You apply to us to lease the equipment listed above to you for the Initial Period referred to above and thereafter in accordance with the terms and conditions stated above and set out overleaf. You agree to pay to us the payments set forth above (which are for lease of the specified equipment, and may include amounts for delivery and installation) in accordance with the frequency set out above. You agree that all information set out herein is correct and that all particulars were complete when this application was signed. You acknowledge having read the terms and conditions of this Agreement set forth on this page and overleaf, and agree that no other terms and conditions, express or implied, are part of this agreement unless they appear above or in a schedule or addendum, and in either event are initialed by both of us to indicate they form part of this Agreement."
appText = appText & Chr(10) & Chr(10) & "PRE-AUTHORIZED DEBIT AUTHORIZATION You hereby authorize us to debit your bank account identified on the void cheque delivered to us (the ""account"") with the amount of each payment or other amount owing from time to time to us under this Agreement on or shortly after the due date thereof as set out in this Agreement, by issuing pre-authorized debit requests (each a ""PAD"") to the financial institution where the account is held (the ""processing institution""). The processing institution is hereby authorized to pay from and to debit against the account, any payment order or request whatsoever, payable to us and drawn on the account by bank acting for us. Any such payment order or request shall be considered as having been signed by you. You acknowledge that this authorization also constitutes delivery thereof by you to the processing institution. "
appText = appText & "You hereby agree that each PAD may be processed without prior written notice from us of either the amount of the PAD or the date that the PAD is to be processed. You may revoke this authorization at any time by giving a 10day written prior notice to us at the address set forth above. "
appText = appText & "You may obtain a sample cancellation form, or further information on your right to cancel this authorization at the processing institution or by visiting www.cdnpay.ca? you have certain recourse rights if any debit does not comply with this PAD agreement. To obtain more information on your recourse rights, contact your financial institution or visit www.cdnpay.ca. ? For example, you have the right to receive reimbursement for any debit that is not authorized or is not consistent with this PAD agreement. To obtain moreinformation on your recourse rights, contact your financial institution or visit www.cdnpay.ca.? Each person whose signature is required on the account must sign below. The Payee may assign or transfer its rights under this PAD Agreement."
appText = appText & Chr(10) & Chr(10) & Chr(10)
appText = appText & Chr(10) & Chr(10) & "Authorized Cheque Signature(s):  ____________________________________________________________________________________Please attach 'void' cheque"
 
 
accepText = "ACCEPTANCE: By signing below, you, as customer, certify that all of the equipment has been delivered, fully installed and accepted as of the date of your signature below and you direct and authorize us to purchase the equipment."
 
finalLine = "Under this Agreement the Equipment remains our property and you may not sell it."
 
'Align columns
Columns("A").ColumnWidth = 6.56
Columns("B").ColumnWidth = 20.67
Columns("C").ColumnWidth = 13.33
Columns("D").ColumnWidth = 2.67
Columns("E").ColumnWidth = 13.78
Columns("F").ColumnWidth = 37.22
Columns("G").ColumnWidth = 0.56
 
'Align Header Rows
Rows(1).RowHeight = 9.6
Rows(2).RowHeight = 14.4
Rows(3).RowHeight = 26.4
Rows(4).RowHeight = 19.8
Rows(5).RowHeight = 13.8
Rows(6).RowHeight = 12.6
Rows(7).RowHeight = 13.8
Rows(8).RowHeight = 12.6
Rows(9).RowHeight = 13.8
Rows(10).RowHeight = 12.6
Rows(11).RowHeight = 13.8
Rows(12).RowHeight = 12.6
Rows(13).RowHeight = 33
Rows(14).RowHeight = 12.6
 
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
 
Cells(6, 2).Value = acctName
Cells(7, 2).Value = acctAddy
Cells(8, 2).Value = acctCity
Cells(9, 2).Value = acctProv
Cells(10, 2).Value = acctPostal

If acctContact = "" Then
    Range(Cells(6, 6), Cells(6, 6)).Interior.ColorIndex = 6
Else
    Cells(6, 6).Value = acctContact
End If

If acctPhone = "" Then
    Range(Cells(7, 6), Cells(7, 6)).Interior.ColorIndex = 6
Else
    Cells(7, 6).Value = acctPhone
End If

If acctFax = "" Then
    Range(Cells(8, 6), Cells(8, 6)).Interior.ColorIndex = 6
Else
    Cells(8, 6).Value = acctFax
End If

If acctEmail = "" Then
    Range(Cells(9, 6), Cells(9, 6)).Interior.ColorIndex = 6
Else
    Cells(9, 6).Value = acctEmail
End If

If acctRep = "" Then
    Range(Cells(10, 6), Cells(10, 6)).Interior.ColorIndex = 6
Else
    Cells(10, 6).Value = acctRep
End If

 
'Get Quantity
 
For i = LA_standardSheetNumber To LA_mySheets
       
    'Get
    addyStreet = Sheets(i).Cells(8, 2).Value
    addyCity = Sheets(i).Cells(9, 2).Value
    addyProv = Sheets(i).Cells(10, 2).Value
    LA_address = addyStreet & " - " & addyCity & ", " & addyProv
    LA_model = Sheets(i).Cells(16, 2).Value
       
   
    'if address of this machine matches address of the last machine, check if the machine is the same model
   If LA_lastAddy <> LA_address Or LA_prevModel <> LA_model Then
       
        'Layout
        Range(Cells(LA_moduleStart, 4), Cells(LA_moduleStart, 6)).Merge
        Range(Cells(LA_moduleStart, 3), Cells(LA_moduleStart, 3)).Interior.ColorIndex = 6
        Range(Cells(LA_moduleStart, 1), Cells(LA_moduleStart, 6)).HorizontalAlignment = xlCenter
        Range(Cells(LA_moduleStart, 1), Cells(LA_moduleStart, 6)).Borders.LineStyle = xlContinuous
        Rows(LA_moduleStart).RowHeight = 14.4
   
        'Set Quantity
      Cells(LA_moduleStart, 1).Value = 1
   
        'Set Address
      Cells(LA_moduleStart, 4).Value = LA_address
 
        'Set Model
      Cells(LA_moduleStart, 2).Value = LA_model
   
        'if quantity < 1 then add 1 to LA_machineIndex
      If LA_modelQty <= 1 Then
            LA_machineIndex = LA_machineIndex + 1
        Else
            LA_modelQty = LA_modelQty - 1
        End If
 
        'Increment line pointer
      LA_moduleStart = LA_moduleStart + 1
       
        lastStreet = addyStreet
        lastCity = addyCity
        lastProv = addyProv
        LA_lastAddy = LA_address
        LA_prevModel = LA_model
       
 
    Else
        'increment Qty and move to the next iteration
      Cells(LA_moduleStart - 1, 1).Value = Cells(LA_moduleStart - 1, 1).Value + 1
       
    End If
   
Next i
 
 
'Border header and body
 
 
'Footer layout
Range(Cells(LA_moduleStart, 1), Cells(LA_moduleStart, 6)).Merge
Range(Cells(LA_moduleStart + 1, 1), Cells(LA_moduleStart + 1, 6)).Merge
Range(Cells(LA_moduleStart + 2, 1), Cells(LA_moduleStart + 2, 2)).Merge
Range(Cells(LA_moduleStart + 2, 4), Cells(LA_moduleStart + 2, 6)).Merge
Range(Cells(LA_moduleStart + 2, 4), Cells(LA_moduleStart + 2, 6)).Merge
Range(Cells(LA_moduleStart + 3, 1), Cells(LA_moduleStart + 3, 2)).Merge
Range(Cells(LA_moduleStart + 3, 4), Cells(LA_moduleStart + 3, 5)).Merge
Range(Cells(LA_moduleStart + 4, 1), Cells(LA_moduleStart + 4, 6)).Merge
Range(Cells(LA_moduleStart + 5, 1), Cells(LA_moduleStart + 5, 5)).Merge
Range(Cells(LA_moduleStart + 6, 1), Cells(LA_moduleStart + 6, 6)).Merge
Range(Cells(LA_moduleStart + 7, 1), Cells(LA_moduleStart + 7, 6)).Merge
Range(Cells(LA_moduleStart + 8, 1), Cells(LA_moduleStart + 8, 2)).Merge
Range(Cells(LA_moduleStart + 8, 3), Cells(LA_moduleStart + 8, 4)).Merge
Range(Cells(LA_moduleStart + 9, 1), Cells(LA_moduleStart + 9, 2)).Merge
Range(Cells(LA_moduleStart + 9, 3), Cells(LA_moduleStart + 9, 4)).Merge
Range(Cells(LA_moduleStart + 9, 6), Cells(LA_moduleStart + 10, 6)).Merge
Range(Cells(LA_moduleStart + 10, 1), Cells(LA_moduleStart + 10, 5)).Merge
 
'Font Settings
Range(Cells(5, 1), Cells(LA_moduleStart + 10, 6)).Font.Name = "arial"
Range(Cells(5, 1), Cells(LA_moduleStart + 10, 6)).Font.Size = 8
 
'Borders
Range(Cells(1, 1), Cells(LA_moduleStart + 1, 6)).BorderAround ColorIndex:=1
Cells(LA_moduleStart + 5, 6).BorderAround ColorIndex:=1
Range(Cells(LA_moduleStart + 2, 1), Cells(LA_moduleStart + 2, 6)).BorderAround ColorIndex:=1
 
Range(Cells(LA_moduleStart + 3, 1), Cells(LA_moduleStart + 3, 3)).BorderAround ColorIndex:=1
Range(Cells(LA_moduleStart + 3, 4), Cells(LA_moduleStart + 3, 6)).BorderAround ColorIndex:=1
 
Range(Cells(LA_moduleStart + 4, 1), Cells(LA_moduleStart + 4, 6)).BorderAround ColorIndex:=1
 
Range(Cells(LA_moduleStart + 5, 1), Cells(LA_moduleStart + 5, 5)).BorderAround ColorIndex:=1
 
Range(Cells(LA_moduleStart + 6, 1), Cells(LA_moduleStart + 6, 6)).BorderAround ColorIndex:=1
 
Range(Cells(LA_moduleStart + 7, 1), Cells(LA_moduleStart + 7, 6)).BorderAround ColorIndex:=1
 
Range(Cells(LA_moduleStart + 8, 1), Cells(LA_moduleStart + 8, 2)).BorderAround ColorIndex:=1
Range(Cells(LA_moduleStart + 8, 3), Cells(LA_moduleStart + 8, 4)).BorderAround ColorIndex:=1
Cells(LA_moduleStart + 8, 5).BorderAround ColorIndex:=1
Cells(LA_moduleStart + 8, 6).BorderAround ColorIndex:=1
 
Range(Cells(LA_moduleStart + 9, 6), Cells(LA_moduleStart + 10, 6)).BorderAround ColorIndex:=1
Range(Cells(LA_moduleStart + 9, 1), Cells(LA_moduleStart + 9, 2)).BorderAround ColorIndex:=1
Range(Cells(LA_moduleStart + 9, 3), Cells(LA_moduleStart + 9, 4)).BorderAround ColorIndex:=1
Cells(LA_moduleStart + 9, 5).BorderAround ColorIndex:=1
 
Range(Cells(LA_moduleStart + 10, 1), Cells(LA_moduleStart + 10, 5)).BorderAround ColorIndex:=1
 
'Align Row heights
Rows(LA_moduleStart).RowHeight = 12
Rows(LA_moduleStart + 1).RowHeight = 10.2
Rows(LA_moduleStart + 2).RowHeight = 19.8
Rows(LA_moduleStart + 3).RowHeight = 15
Rows(LA_moduleStart + 4).RowHeight = 24
Rows(LA_moduleStart + 5).RowHeight = 21.6
Rows(LA_moduleStart + 6).RowHeight = 185.4
Rows(LA_moduleStart + 7).RowHeight = 21
Rows(LA_moduleStart + 8).RowHeight = 21
Rows(LA_moduleStart + 9).RowHeight = 49.2
Rows(LA_moduleStart + 10).RowHeight = 12
 
 
'Layout Base Content
Cells(LA_moduleStart + 2, 1).Value = "Payment Amount: "
Cells(LA_moduleStart + 2, 1).HorizontalAlignment = xlCenter
Cells(LA_moduleStart + 2, 1).VerticalAlignment = xlCenter
Cells(LA_moduleStart + 2, 3).Value = FormatCurrency(Sheets(3).Cells(25, 5).Value, 2)
Cells(LA_moduleStart + 2, 4).Value = "+ all applicable taxes per period"
 
Cells(LA_moduleStart + 3, 1).Value = "Payment Frequency: "
Cells(LA_moduleStart + 3, 1).HorizontalAlignment = xlCenter
Cells(LA_moduleStart + 3, 3).Value = Sheets(3).Cells(16, 4).Value
Cells(LA_moduleStart + 3, 3).HorizontalAlignment = xlCenter
Cells(LA_moduleStart + 3, 4).Value = "Term (in Months): "
Cells(LA_moduleStart + 3, 6).Value = Sheets(3).Cells(15, 4).Value
 
Cells(LA_moduleStart + 4, 1).Value = "The first lease payment is payable on acceptance of this Agreement and thereafter on the first day of each lease period according to the lease payment frequency selected."
Cells(LA_moduleStart + 4, 1).WrapText = True
 
Cells(LA_moduleStart + 5, 1).Value = "Special Provisions: "
Cells(LA_moduleStart + 5, 6).Value = "Customer" & Chr(10) & "Initial: "
Range(Cells(LA_moduleStart + 5, 1), Cells(LA_moduleStart + 5, 6)).VerticalAlignment = xlTop
 
 
'App Text
Cells(LA_moduleStart + 6, 1).Font.Size = 6
Cells(LA_moduleStart + 6, 1).VerticalAlignment = xlTop
Cells(LA_moduleStart + 6, 1).Value = appText
Cells(LA_moduleStart + 6, 1).Characters(1, 15).Font.Bold = True
Cells(LA_moduleStart + 6, 1).Characters(15, 891).Font.Bold = False
Cells(LA_moduleStart + 6, 1).Characters(892, 927).Font.Bold = True
Cells(LA_moduleStart + 6, 1).Characters(928, 2015).Font.Bold = False
 
'Acceptance Text
Cells(LA_moduleStart + 7, 1).Value = accepText
Cells(LA_moduleStart + 7, 1).Font.Size = 6.5
Cells(LA_moduleStart + 7, 1).WrapText = True
 
'CUSTOMER Signature
Cells(LA_moduleStart + 8, 1).Value = "CUSTOMER Signature"
Cells(LA_moduleStart + 8, 1).Font.Bold = True
Cells(LA_moduleStart + 8, 1).Font.Italic = True
Cells(LA_moduleStart + 8, 1).HorizontalAlignment = xlCenter
Cells(LA_moduleStart + 8, 1).VerticalAlignment = xlCenter
 
'Print Name and Position
Cells(LA_moduleStart + 8, 3).Value = "Print Name and Position"
Cells(LA_moduleStart + 8, 3).Font.Bold = True
Cells(LA_moduleStart + 8, 3).Font.Italic = True
Cells(LA_moduleStart + 8, 3).HorizontalAlignment = xlCenter
Cells(LA_moduleStart + 8, 3).VerticalAlignment = xlCenter
 
'Date Signed
Cells(LA_moduleStart + 8, 5).Value = "Date Signed"
Cells(LA_moduleStart + 8, 5).Font.Bold = True
Cells(LA_moduleStart + 8, 5).Font.Italic = True
Cells(LA_moduleStart + 8, 5).HorizontalAlignment = xlCenter
Cells(LA_moduleStart + 8, 5).VerticalAlignment = xlCenter
 
'OWNER (Document Direction Limited)
Cells(LA_moduleStart + 8, 6).Value = "OWNER (Document Direction Limited)"
Cells(LA_moduleStart + 8, 6).Font.Bold = True
Cells(LA_moduleStart + 8, 6).Font.Italic = True
Cells(LA_moduleStart + 8, 6).HorizontalAlignment = xlCenter
Cells(LA_moduleStart + 8, 6).VerticalAlignment = xlCenter
 
'Final line
Cells(LA_moduleStart + 10, 1).Value = finalLine
End Sub

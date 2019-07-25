Attribute VB_Name = "Module8"
Private Sub Workbook_Open()

'Get Account, Customer, and Rep
orderChecklist = "Order Checklist"
accountSheet = "Account Info-DO NOT DELETE"

repName = Sheets(accountSheet).Cells(12, 2).Value
accountNumber = Sheets(accountSheet).Cells(17, 2).Value
customerName = Sheets(accountSheet).Cells(21, 2).Value

'insert values
Sheets(orderChecklist).Cells(1, 10).Value = repName
Sheets(orderChecklist).Cells(2, 10).Value = Date
Sheets(orderChecklist).Cells(4, 3).Value = customerName
Sheets(orderChecklist).Cells(5, 3).Value = accountNumber

End Sub


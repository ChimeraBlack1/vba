Attribute VB_Name = "Module4"
Sub IQ20_Button1_Click()

inqSheet = "INQ 2.0"
Sheets(inqSheet).Activate

'Align Columns
Columns("A").ColumnWidth = 0
Columns("B").ColumnWidth = 0
Columns("C").ColumnWidth = 4.22
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
Columns("AD").ColumnWidth = 0.94
Columns("AE").ColumnWidth = 0.94
Columns("AF").ColumnWidth = 0.94
Columns("AG").ColumnWidth = 0.94
Columns("AH").ColumnWidth = 0.94
Columns("AI").ColumnWidth = 0.94
Columns("AJ").ColumnWidth = 0.94
Columns("AK").ColumnWidth = 0.94
Columns("AL").ColumnWidth = 0.94
Columns("AM").ColumnWidth = 0.94
Columns("AN").ColumnWidth = 4.89
Columns("AO").ColumnWidth = 3.11
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
Columns("BF").ColumnWidth = 5.78
Columns("BG").ColumnWidth = 5.78
Columns("BH").ColumnWidth = 3.33
Columns("BI").ColumnWidth = 0.94
Columns("BJ").ColumnWidth = 0.94
Columns("BK").ColumnWidth = 0.94
Columns("BL").ColumnWidth = 0.94

'Font settings
Range(Cells(1, 1), Cells(1000, 60)).Font.Name = "Times New Roman"
Range(Cells(8, 1), Cells(1000, 60)).Font.Size = 8

'Get account details
accountName = Sheets(1).Cells(21, 2).Value
accountRepPhone = Sheets(1).Cells(13, 2).Value

If accountRepPhone = "" Then
    Range(Cells(moduleStart + 8, 50), Cells(moduleStart + 8, 50)).Interior.ColorIndex = 6
Else
    Cells(moduleStart + 8, 50).Value = accountRepPhone
End If


accountRep = Sheets(1).Cells(12, 2).Value
accountITContact = Sheets(1).Cells(37, 2).Value
accountITPhone = Sheets(1).Cells(39, 2).Value
acctITEmail = Sheets(1).Cells(38, 2).Value

specialNoteText = "Special Note: Due to the ever-changing infrastructure and/or security measures taken with most computer networks, it cannot be guaranteed that the scan to email or folder options will remain operational without periodic configuration adjustments. It is strongly recommended that you (the customer) take an active role in the scan feature setup at the time of installation of your new equipment. Scan related failures and associated service calls that are a result of changes to network infrastructure (hardware, software, passwords and security) are not covered under a maintenance contract and could be subject to a charge at current service rates."

standardSheetNumber = 15
mySheets = Worksheets.Count
moduleStart = 0

For i = standardSheetNumber To mySheets

    ' Need to get model
    
    thisModel = Sheets(i).Cells(16, 2).Value
    accountAddress = Sheets(i).Cells(8, 2).Value
    accountContact = Sheets(i).Cells(7, 6).Value
    accountCity = Sheets(i).Cells(9, 2).Value & ", " & Sheets(i).Cells(10, 2).Value
    accountPhone = Sheets(i).Cells(8, 6).Value
    acctPostal = Sheets(i).Cells(11, 2).Value
    acctEmail = Sheets(i).Cells(9, 6).Value
    
    If accountRep = "" Then
        Range(Cells(moduleStart + 8, 16), Cells(moduleStart + 8, 16)).Interior.ColorIndex = 6
    Else
        Cells(moduleStart + 8, 16).Value = accountRep
    End If
        
    If acctITEmail = "" Then
        Range(Cells(moduleStart + 25, 45), Cells(moduleStart + 25, 45)).Interior.ColorIndex = 6
    Else
        Cells(moduleStart + 25, 45).Value = acctITEmail
    End If
    
    If accountPhone = "" Then
        Range(Cells(moduleStart + 21, 16), Cells(moduleStart + 21, 16)).Interior.ColorIndex = 6
    Else
        Cells(moduleStart + 21, 16).Value = accountPhone
    End If
    
    If accountContact = "" Then
        Range(Cells(moduleStart + 15, 16), Cells(moduleStart + 15, 16)).Interior.ColorIndex = 6
    Else
        Cells(moduleStart + 15, 16).Value = accountContact
    End If
    
    If accountITContact = "" Then
        Range(Cells(moduleStart + 23, 16), Cells(moduleStart + 23, 16)).Interior.ColorIndex = 6
    Else
        Cells(moduleStart + 23, 16).Value = accountITContact
    End If
    
    If accountITPhone = "" Then
        Range(Cells(moduleStart + 25, 16), Cells(moduleStart + 25, 16)).Interior.ColorIndex = 6
    Else
        Cells(moduleStart + 25, 16).Value = accountITPhone
    End If
    
    If acctEmail = "" Then
        Range(Cells(moduleStart + 21, 45), Cells(moduleStart + 21, 45)).Interior.ColorIndex = 6
    Else
        Cells(moduleStart + 21, 45).Value = acctEmail
    End If
    
    'Align Rows
    Rows(moduleStart + 1).RowHeight = 5
    Rows(moduleStart + 2).RowHeight = 15.6
    Rows(moduleStart + 3).RowHeight = 5
    Rows(moduleStart + 4).RowHeight = 16
    Rows(moduleStart + 5).RowHeight = 5
    Rows(moduleStart + 6).RowHeight = 5
    Rows(moduleStart + 7).RowHeight = 4
    Rows(moduleStart + 8).RowHeight = 9
    Rows(moduleStart + 9).RowHeight = 5
    Rows(moduleStart + 10).RowHeight = 9
    Rows(moduleStart + 11).RowHeight = 5
    Rows(moduleStart + 12).RowHeight = 5
    Rows(moduleStart + 13).RowHeight = 9
    Rows(moduleStart + 14).RowHeight = 4.2
    Rows(moduleStart + 15).RowHeight = 9
    Rows(moduleStart + 16).RowHeight = 4
    Rows(moduleStart + 17).RowHeight = 9
    Rows(moduleStart + 18).RowHeight = 4
    Rows(moduleStart + 19).RowHeight = 9
    Rows(moduleStart + 20).RowHeight = 4
    Rows(moduleStart + 21).RowHeight = 9
    Rows(moduleStart + 22).RowHeight = 9
    Rows(moduleStart + 23).RowHeight = 9
    Rows(moduleStart + 24).RowHeight = 4
    Rows(moduleStart + 25).RowHeight = 9
    Rows(moduleStart + 26).RowHeight = 4
    Rows(moduleStart + 27).RowHeight = 15.8
    Rows(moduleStart + 28).RowHeight = 8.8
    Rows(moduleStart + 29).RowHeight = 10
    Rows(moduleStart + 30).RowHeight = 4
    Rows(moduleStart + 31).RowHeight = 9
    Rows(moduleStart + 32).RowHeight = 4
    Rows(moduleStart + 33).RowHeight = 9
    Rows(moduleStart + 34).RowHeight = 9
    Rows(moduleStart + 35).RowHeight = 9
    Rows(moduleStart + 36).RowHeight = 9
    Rows(moduleStart + 37).RowHeight = 10
    Rows(moduleStart + 38).RowHeight = 9
    Rows(moduleStart + 39).RowHeight = 10
    Rows(moduleStart + 40).RowHeight = 9
    Rows(moduleStart + 41).RowHeight = 9
    Rows(moduleStart + 42).RowHeight = 10
    Rows(moduleStart + 43).RowHeight = 9
    Rows(moduleStart + 44).RowHeight = 9
    Rows(moduleStart + 45).RowHeight = 9
    Rows(moduleStart + 46).RowHeight = 9
    Rows(moduleStart + 47).RowHeight = 10
    Rows(moduleStart + 48).RowHeight = 9
    Rows(moduleStart + 49).RowHeight = 6
    Rows(moduleStart + 50).RowHeight = 9
    Rows(moduleStart + 51).RowHeight = 9
    Rows(moduleStart + 52).RowHeight = 9
    Rows(moduleStart + 53).RowHeight = 9
    Rows(moduleStart + 54).RowHeight = 4
    Rows(moduleStart + 55).RowHeight = 12
    Rows(moduleStart + 56).RowHeight = 12
    Rows(moduleStart + 57).RowHeight = 9
    Rows(moduleStart + 58).RowHeight = 10
    Rows(moduleStart + 59).RowHeight = 9
    Rows(moduleStart + 60).RowHeight = 10
    Rows(moduleStart + 61).RowHeight = 9
    Rows(moduleStart + 62).RowHeight = 11
    Rows(moduleStart + 63).RowHeight = 4.8
    Rows(moduleStart + 64).RowHeight = 10
    Rows(moduleStart + 65).RowHeight = 4
    Rows(moduleStart + 66).RowHeight = 10
    Rows(moduleStart + 67).RowHeight = 8.6
    Rows(moduleStart + 68).RowHeight = 11
    Rows(moduleStart + 69).RowHeight = 8
    Rows(moduleStart + 70).RowHeight = 10
    Rows(moduleStart + 71).RowHeight = 9
    Rows(moduleStart + 72).RowHeight = 9
    Rows(moduleStart + 73).RowHeight = 8
    Rows(moduleStart + 74).RowHeight = 5
    Rows(moduleStart + 75).RowHeight = 10
    Rows(moduleStart + 76).RowHeight = 9
    Rows(moduleStart + 77).RowHeight = 4
    Rows(moduleStart + 78).RowHeight = 12.2
    Rows(moduleStart + 79).RowHeight = 9
    Rows(moduleStart + 80).RowHeight = 9
    Rows(moduleStart + 81).RowHeight = 9
    Rows(moduleStart + 82).RowHeight = 9
    Rows(moduleStart + 83).RowHeight = 6.9
    Rows(moduleStart + 84).RowHeight = 7
    Rows(moduleStart + 85).RowHeight = 9.8
    Rows(moduleStart + 86).RowHeight = 9.8
    Rows(moduleStart + 87).RowHeight = 9.8
    Rows(moduleStart + 88).RowHeight = 10.4
    Rows(moduleStart + 89).RowHeight = 9.8
    Rows(moduleStart + 90).RowHeight = 11.4
    
    
    'Work order
    Cells(moduleStart + 2, 46).Value = "WO#"
    Range(Cells(moduleStart + 2, 50), Cells(moduleStart + 2, 59)).Merge
    Range(Cells(moduleStart + 2, 50), Cells(moduleStart + 2, 59)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range(Cells(moduleStart + 2, 50), Cells(moduleStart + 2, 59)).Interior.ColorIndex = 6
    
    Cells(moduleStart + 4, 19).Value = "Installation Network Questionnaire"
    Cells(moduleStart + 4, 19).Font.Bold = True
    Cells(moduleStart + 4, 19).Font.Size = 14
    Cells(moduleStart + 4, 19).VerticalAlignment = xlCenter
    
    'Sales Rep Details
    Cells(moduleStart + 8, 4).Value = "Sales Representative:"
    Range(Cells(moduleStart + 8, 16), Cells(moduleStart + 8, 40)).Merge
    Range(Cells(moduleStart + 8, 16), Cells(moduleStart + 8, 40)).Borders(xlEdgeBottom).LineStyle = xlContinuous

    Cells(moduleStart + 8, 16).HorizontalAlignment = xlCenter
    Cells(moduleStart + 8, 16).VerticalAlignment = xlCenter
    
    Cells(moduleStart + 8, 43).Value = "Rep Phone:"
    Range(Cells(moduleStart + 8, 50), Cells(moduleStart + 8, 59)).Merge
    Range(Cells(moduleStart + 8, 50), Cells(moduleStart + 8, 59)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    Cells(moduleStart + 8, 50).HorizontalAlignment = xlCenter
    Cells(moduleStart + 8, 50).VerticalAlignment = xlCenter
    
    Cells(moduleStart + 10, 4).Value = "Customer Name:"
    Range(Cells(moduleStart + 10, 16), Cells(moduleStart + 10, 40)).Merge
    Range(Cells(moduleStart + 10, 16), Cells(moduleStart + 10, 40)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    Cells(moduleStart + 10, 16).Value = accountName
    Cells(moduleStart + 10, 16).HorizontalAlignment = xlCenter
    Cells(moduleStart + 10, 16).VerticalAlignment = xlCenter
    
    Range(Cells(moduleStart + 8, 4), Cells(moduleStart + 10, 60)).VerticalAlignment = xlCenter
    
    Range(Cells(moduleStart + 7, 3), Cells(moduleStart + 12, 60)).BorderAround ColorIndex:=1, Weight:=xlMedium
    
    'Contact Details
    Range(Cells(moduleStart + 5, 4), Cells(moduleStart + 90, 60)).VerticalAlignment = xlCenter
    
    Cells(moduleStart + 15, 4).Value = "Primary Contact:"
    Range(Cells(moduleStart + 15, 16), Cells(moduleStart + 15, 40)).Merge
    Range(Cells(moduleStart + 15, 16), Cells(moduleStart + 15, 40)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    Cells(moduleStart + 15, 16).HorizontalAlignment = xlCenter
    Cells(moduleStart + 15, 16).VerticalAlignment = xlCenter
    
    Cells(moduleStart + 17, 4).Value = "Address:"
    Range(Cells(moduleStart + 17, 16), Cells(moduleStart + 17, 40)).Merge
    Range(Cells(moduleStart + 17, 16), Cells(moduleStart + 17, 40)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    Cells(moduleStart + 17, 16).Value = accountAddress
    Cells(moduleStart + 17, 16).HorizontalAlignment = xlCenter
    Cells(moduleStart + 17, 16).VerticalAlignment = xlCenter
    
    Cells(moduleStart + 19, 4).Value = "City / Province:"
    Range(Cells(moduleStart + 19, 16), Cells(moduleStart + 19, 40)).Merge
    Range(Cells(moduleStart + 19, 16), Cells(moduleStart + 19, 40)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    Cells(moduleStart + 19, 16).Value = accountCity
    Cells(moduleStart + 19, 16).HorizontalAlignment = xlCenter
    Cells(moduleStart + 19, 16).VerticalAlignment = xlCenter
    
    Cells(moduleStart + 21, 4).Value = "Phone:"
    Range(Cells(moduleStart + 21, 16), Cells(moduleStart + 21, 31)).Merge
    Range(Cells(moduleStart + 21, 16), Cells(moduleStart + 21, 31)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    
    Cells(moduleStart + 21, 16).HorizontalAlignment = xlCenter
    Cells(moduleStart + 21, 16).VerticalAlignment = xlCenter
    
    Cells(moduleStart + 23, 4).Value = "IT Contact:"
    Range(Cells(moduleStart + 23, 16), Cells(moduleStart + 23, 40)).Merge
    Range(Cells(moduleStart + 23, 16), Cells(moduleStart + 23, 40)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    
    Cells(moduleStart + 23, 16).VerticalAlignment = xlCenter
    
    Cells(moduleStart + 25, 4).Value = "IT Phone:"
    Range(Cells(moduleStart + 25, 16), Cells(moduleStart + 25, 31)).Merge
    Range(Cells(moduleStart + 25, 16), Cells(moduleStart + 25, 31)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    
    Cells(moduleStart + 25, 16).HorizontalAlignment = xlCenter
    Cells(moduleStart + 25, 16).VerticalAlignment = xlCenter
    
    Cells(moduleStart + 19, 42).Value = "Postal Code:"
    Range(Cells(moduleStart + 19, 50), Cells(moduleStart + 19, 59)).Merge
    Range(Cells(moduleStart + 19, 50), Cells(moduleStart + 19, 59)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    Cells(moduleStart + 19, 50).Value = acctPostal
    Cells(moduleStart + 19, 50).VerticalAlignment = xlCenter
    Cells(moduleStart + 19, 50).HorizontalAlignment = xlCenter
    
    Cells(moduleStart + 21, 36).Value = "E-Mail:"
    Range(Cells(moduleStart + 21, 45), Cells(moduleStart + 21, 59)).Merge
    Range(Cells(moduleStart + 21, 45), Cells(moduleStart + 21, 59)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    
    Cells(moduleStart + 21, 45).HorizontalAlignment = xlCenter
    
    
    Cells(moduleStart + 25, 36).Value = "IT E-Mail:"
    Range(Cells(moduleStart + 25, 45), Cells(moduleStart + 25, 59)).Merge
    Range(Cells(moduleStart + 25, 45), Cells(moduleStart + 25, 59)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    
    Cells(moduleStart + 25, 45).HorizontalAlignment = xlCenter
    
    'SECTION BORDERS
    Range(Cells(moduleStart + 14, 3), Cells(moduleStart + 26, 60)).BorderAround ColorIndex:=1, Weight:=xlMedium
    Range(Cells(moduleStart + 14, 3), Cells(moduleStart + 26, 60)).BorderAround ColorIndex:=1, Weight:=xlMedium
      
      '///////////////////////////////////// Does the cust have @remote /////////////////////////////////////////
      
    Range(Cells(moduleStart + 28, 3), Cells(moduleStart + 31, 60)).BorderAround ColorIndex:=1, Weight:=xlMedium
    Cells(moduleStart + 29, 4).Value = "THE CLIENT HAS @REMOTE APPLIANCES"
    Cells(moduleStart + 29, 4).Font.Bold = True
    Cells(moduleStart + 29, 45).Value = "Yes"
    Cells(moduleStart + 29, 45).Font.Bold = True
    Range(Cells(moduleStart + 29, 44), Cells(moduleStart + 29, 44)).BorderAround ColorIndex:=1
    Cells(moduleStart + 29, 53).Value = "No"
    Cells(moduleStart + 29, 53).Font.Bold = True
    Range(Cells(moduleStart + 29, 52), Cells(moduleStart + 29, 52)).BorderAround ColorIndex:=1
    moduleStart = moduleStart + 6
    
    '//////////////////////////////////// MAIN SECTION ///////////////////////////////////////////
    Range(Cells(moduleStart + 28, 3), Cells(moduleStart + 67, 60)).BorderAround ColorIndex:=1, Weight:=xlMedium
    Cells(moduleStart + 29, 4).Value = "Server"
    Cells(moduleStart + 29, 4).Font.Bold = True
    Range(Cells(moduleStart + 29, 11), Cells(moduleStart + 29, 11)).BorderAround ColorIndex:=1
    
    Cells(moduleStart + 29, 19).Value = "No"
    Cells(moduleStart + 29, 13).Value = "Yes"
    Range(Cells(moduleStart + 29, 17), Cells(moduleStart + 29, 17)).BorderAround ColorIndex:=1
    
    Cells(moduleStart + 29, 45).Value = "32 Bit"
    Range(Cells(moduleStart + 29, 44), Cells(moduleStart + 29, 44)).BorderAround ColorIndex:=1
    
    Cells(moduleStart + 29, 53).Value = "64 Bit"
    Range(Cells(moduleStart + 29, 52), Cells(moduleStart + 29, 52)).BorderAround ColorIndex:=1
    
    Cells(moduleStart + 29, 29).Value = "Server Operating System"
    Cells(moduleStart + 29, 29).Font.Bold = True
    
    Cells(moduleStart + 31, 6).Value = "Windows"
    Range(Cells(moduleStart + 31, 4), Cells(moduleStart + 31, 4)).BorderAround ColorIndex:=1
    
    Cells(moduleStart + 31, 13).Value = "Ver:"
    Cells(moduleStart + 31, 23).Value = "SP:"
    
    Cells(moduleStart + 31, 32).Value = "Macintosh"
    Range(Cells(moduleStart + 31, 30), Cells(moduleStart + 31, 30)).BorderAround ColorIndex:=1
    
    Cells(moduleStart + 31, 44).Value = "Ver:"
    Cells(moduleStart + 31, 54).Value = "SP:"
    
    Cells(moduleStart + 33, 6).Value = "Unix"
    Range(Cells(moduleStart + 33, 4), Cells(moduleStart + 33, 4)).BorderAround ColorIndex:=1
    
    Cells(moduleStart + 33, 13).Value = "Ver:"
    Cells(moduleStart + 33, 23).Value = "SP:"
    
    Cells(moduleStart + 33, 32).Value = "Other"
    Range(Cells(moduleStart + 33, 30), Cells(moduleStart + 33, 30)).BorderAround ColorIndex:=1
    
    Cells(moduleStart + 35, 4).Value = "Print Driver Requirements"
    Cells(moduleStart + 35, 4).Font.Bold = True
    
    Cells(moduleStart + 37, 6).Value = "PCL"
    Range(Cells(moduleStart + 37, 4), Cells(moduleStart + 37, 4)).BorderAround ColorIndex:=1
    
    Cells(moduleStart + 37, 22).Value = "P.S. (Option)"
    Range(Cells(moduleStart + 37, 19), Cells(moduleStart + 37, 19)).BorderAround ColorIndex:=1
    
    '39
    Cells(moduleStart + 39, 4).Value = "Workstation Operating Systems"
    Cells(moduleStart + 39, 4).Font.Bold = True
    Cells(moduleStart + 39, 33).Value = "Note: Install incl 4 stations at time of initial Install"
    
    '40
    Cells(moduleStart + 40, 5).Value = "(Indicate # of clients and version)"
    Cells(moduleStart + 40, 33).Value = "Additional Clients (@ $35.00)"
    
    '42
    Range(Cells(moduleStart + 42, 4), Cells(moduleStart + 42, 4)).BorderAround ColorIndex:=1
    Range(Cells(moduleStart + 42, 34), Cells(moduleStart + 42, 34)).BorderAround ColorIndex:=1
    Cells(moduleStart + 42, 6).Value = "Windows"
    Cells(moduleStart + 42, 13).Value = "Ver:"
    Cells(moduleStart + 42, 20).Value = "# of clients"
    Cells(moduleStart + 42, 36).Value = "Other, specify"
    
    '44
    Range(Cells(moduleStart + 44, 4), Cells(moduleStart + 44, 4)).BorderAround ColorIndex:=1
    Cells(moduleStart + 44, 6).Value = "Macintosh"
    Cells(moduleStart + 44, 13).Value = "Ver:"
    Cells(moduleStart + 44, 20).Value = "# of clients"
    
    '46
    Range(Cells(moduleStart + 46, 17), Cells(moduleStart + 46, 17)).BorderAround ColorIndex:=1
    Range(Cells(moduleStart + 46, 31), Cells(moduleStart + 46, 31)).BorderAround ColorIndex:=1
    Range(Cells(moduleStart + 46, 50), Cells(moduleStart + 46, 50)).BorderAround ColorIndex:=1
    Cells(moduleStart + 46, 52).Value = "Not Req"
    Cells(moduleStart + 46, 33).Value = "Scan to email - Mail Server **"
    Cells(moduleStart + 46, 19).Value = "Scan to Folder"
    Cells(moduleStart + 46, 4).Value = "Embedded Scanning"
    Cells(moduleStart + 46, 4).Font.Bold = True
    Cells(moduleStart + 46, 19).Font.Bold = False
    Cells(moduleStart + 46, 33).Font.Bold = False
    Cells(moduleStart + 46, 52).Font.Bold = False
    
    '47
    Cells(moduleStart + 47, 8).Value = "** if this is a new installation and not an upgrade to your existing Ricoh device, please enter Mail Server SMTP Address:"
    Range(Cells(moduleStart + 47, 4), Cells(moduleStart + 60, 50)).Font.Bold = False
    
    '48
    Cells(moduleStart + 48, 8).Value = "**"
    Cells(moduleStart + 48, 9).Value = "Please note: SMTP (email server) name and password will be required for scan to email setup"
    
    
    '51
    Range(Cells(moduleStart + 51, 4), Cells(moduleStart + 51, 4)).BorderAround ColorIndex:=1
    Range(Cells(moduleStart + 51, 19), Cells(moduleStart + 51, 19)).BorderAround ColorIndex:=1
    Range(Cells(moduleStart + 51, 27), Cells(moduleStart + 51, 27)).BorderAround ColorIndex:=1
    Range(Cells(moduleStart + 51, 35), Cells(moduleStart + 51, 35)).BorderAround ColorIndex:=1
    Cells(moduleStart + 51, 6).Value = "Inbound Fax routing"
    Cells(moduleStart + 51, 21).Value = "No"
    Cells(moduleStart + 51, 29).Value = "Folder"
    Cells(moduleStart + 51, 37).Value = "Email - Address:"
    
    
    '53
    Range(Cells(moduleStart + 53, 19), Cells(moduleStart + 53, 19)).BorderAround ColorIndex:=1
    Range(Cells(moduleStart + 53, 38), Cells(moduleStart + 53, 38)).BorderAround ColorIndex:=1
    Range(Cells(moduleStart + 53, 51), Cells(moduleStart + 53, 51)).BorderAround ColorIndex:=1
    Cells(moduleStart + 53, 4).Value = "@ REMOTE Embedded"
    Cells(moduleStart + 53, 22).Value = "Auto Toner Replenishment**"
    Cells(moduleStart + 53, 4).Font.Bold = True
    Cells(moduleStart + 53, 40).Value = "Other (see notes)"
    Cells(moduleStart + 53, 53).Value = "Opt Out"
    
    
    '55
    Cells(moduleStart + 55, 4).Value = "Meter reads will be submitted to Ricoh through @Remote or MYRICOH.  A $30 Administration Fee will be applied if this process is not used."
    Cells(moduleStart + 55, 4).Font.Bold = True
    Cells(moduleStart + 55, 4).Font.Name = "Times New Roman"
    Cells(moduleStart + 55, 4).Font.Size = 9
    
    '56
    Cells(moduleStart + 56, 7).Value = "**Please note you need to fill in the following cells below: Toner Recipient Name, email, ship to location and notes:"
    Cells(moduleStart + 56, 7).Font.Bold = False
    Cells(moduleStart + 56, 7).Font.Name = "Times New Roman"
    Cells(moduleStart + 56, 7).Font.Size = 9
    
    '58
    Range(Cells(moduleStart + 58, 4), Cells(moduleStart + 58, 4)).BorderAround ColorIndex:=1
    Cells(moduleStart + 58, 6).Value = "Network Drop"
    Cells(moduleStart + 58, 6).Font.Bold = True
    Range(Cells(moduleStart + 58, 24), Cells(moduleStart + 58, 24)).BorderAround ColorIndex:=1
    Cells(moduleStart + 58, 26).Value = "Wireless (Select One)"
    Cells(moduleStart + 58, 26).Characters(0, 8).Font.Bold = True
    Cells(moduleStart + 58, 42).Value = "Network Cables not included"
    
    '60
    Cells(moduleStart + 60, 4).Value = "Device IP:"
    Cells(moduleStart + 60, 24).Value = "Subnet:"
    Cells(moduleStart + 60, 42).Value = "Gateway:"
    Range(Cells(moduleStart + 60, 4), Cells(moduleStart + 60, 60)).Font.Bold = True
    
    '62
    Range(Cells(moduleStart + 62, 4), Cells(moduleStart + 62, 4)).BorderAround ColorIndex:=1
    Range(Cells(moduleStart + 62, 24), Cells(moduleStart + 62, 24)).BorderAround ColorIndex:=1
    Cells(moduleStart + 62, 6).Value = "Net New"
    Cells(moduleStart + 62, 26).Value = "Replacement/Upgrade"
    Cells(moduleStart + 62, 42).Value = "Location:"
    Range(Cells(moduleStart + 62, 6), Cells(moduleStart + 62, 60)).Font.Bold = True
    
    '64
    Cells(moduleStart + 64, 4).Value = "Toner Recipient name:"
    Cells(moduleStart + 64, 33).Value = "Toner Recipient email:"
    
    '66
    Cells(moduleStart + 66, 4).Value = "Toner Recipient Phone #:"
    Cells(moduleStart + 66, 33).Value = "Notes:"
    
    
    '////////// POWER (image) SECTION ///////////
    Range(Cells(moduleStart + 69, 3), Cells(moduleStart + 77, 60)).BorderAround ColorIndex:=1, Weight:=xlMedium
    Cells(moduleStart + 70, 24).Value = "For Internal Use only"
    Cells(moduleStart + 71, 4).Value = "Model"
    Cells(moduleStart + 72, 4).Value = thisModel
    
    
    '72
    Range(Cells(moduleStart + 72, 15), Cells(moduleStart + 72, 15)).BorderAround ColorIndex:=1
    Range(Cells(moduleStart + 72, 24), Cells(moduleStart + 72, 24)).BorderAround ColorIndex:=1
    Range(Cells(moduleStart + 72, 36), Cells(moduleStart + 72, 36)).BorderAround ColorIndex:=1
    Range(Cells(moduleStart + 72, 48), Cells(moduleStart + 72, 48)).BorderAround ColorIndex:=1
    
    '75
    Range(Cells(moduleStart + 75, 4), Cells(moduleStart + 76, 60)).Font.Bold = True
    Cells(moduleStart + 75, 4).Value = "Power Requirement"
    Cells(moduleStart + 75, 19).Value = "NEMA 5-15R"
    Cells(moduleStart + 75, 30).Value = "NEMA 5-20R"
    Cells(moduleStart + 75, 41).Value = "NEMA 6-20R"
    Cells(moduleStart + 75, 50).Value = "NEMA 6-30R"
    
    '76
    Cells(moduleStart + 76, 19).Value = "120V, 15A"
    Cells(moduleStart + 76, 30).Value = "120V, 20A"
    Cells(moduleStart + 76, 41).Value = "208V, 20A"
    Cells(moduleStart + 76, 51).Value = "208V, 20A"
    
    
    '///////////////// SPECIAL NOTE ///////////////////
    Range(Cells(moduleStart + 79, 3), Cells(moduleStart + 84, 60)).BorderAround ColorIndex:=1, Weight:=xlMedium
    Range(Cells(moduleStart + 79, 4), Cells(moduleStart + 84, 59)).Merge
    Range(Cells(moduleStart + 79, 4), Cells(moduleStart + 84, 59)).WrapText = True
    Range(Cells(moduleStart + 79, 4), Cells(moduleStart + 84, 59)).VerticalAlignment = xlTop
    Range(Cells(moduleStart + 79, 4), Cells(moduleStart + 84, 59)).HorizontalAlignment = xlLeft
    Cells(moduleStart + 79, 4).Value = specialNoteText
    
    Range(Cells(moduleStart + 85, 3), Cells(moduleStart + 88, 16)).BorderAround ColorIndex:=1, Weight:=xlMedium
    Range(Cells(moduleStart + 85, 17), Cells(moduleStart + 88, 31)).BorderAround ColorIndex:=1, Weight:=xlMedium
    Range(Cells(moduleStart + 85, 32), Cells(moduleStart + 88, 48)).BorderAround ColorIndex:=1, Weight:=xlMedium
    Range(Cells(moduleStart + 85, 49), Cells(moduleStart + 88, 60)).BorderAround ColorIndex:=1, Weight:=xlMedium
    
    Range(Cells(moduleStart + 89, 3), Cells(moduleStart + 90, 16)).BorderAround ColorIndex:=1, Weight:=xlMedium
    Range(Cells(moduleStart + 89, 17), Cells(moduleStart + 90, 31)).BorderAround ColorIndex:=1, Weight:=xlMedium
    Range(Cells(moduleStart + 89, 32), Cells(moduleStart + 90, 48)).BorderAround ColorIndex:=1, Weight:=xlMedium
    Range(Cells(moduleStart + 89, 49), Cells(moduleStart + 90, 60)).BorderAround ColorIndex:=1, Weight:=xlMedium
    
    Cells(moduleStart + 87, 49).Value = "Install Completed"
    Cells(moduleStart + 88, 49).Value = "Date"
    Cells(moduleStart + 89, 5).Value = "Date (mm/dd/yy)"
    Cells(moduleStart + 89, 18).Value = "Client Signature"
    Cells(moduleStart + 89, 34).Value = "Client Name"
    Cells(moduleStart + 90, 34).Value = "(Please Print)"
    Cells(moduleStart + 89, 49).Value = "Customer"
    Cells(moduleStart + 90, 49).Value = "Sign Off:"
    
    moduleStart = moduleStart + 93

Next i


End Sub



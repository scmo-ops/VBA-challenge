Sub loopingThroughWorksheets()


'sheet loop start
For Each ws In ActiveWorkbook.Worksheets
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total stock volume"
ws.Range("O1").Value = "Ticker"
ws.Range("P1").Value = "Value"
ws.Range("N2").Value = "Greatest % Increase"
ws.Range("N3").Value = "Greatest % Decrease"
ws.Range("N4").Value = "Greatest Total Volume"
'Variables needed

Dim TickerLetter As String
Dim counter As Double
Dim percentinperfect As Double
Dim totalvolume As Double
Dim nepo As Double
Dim esolc As Double 'For some reason I can't write open or close as variables
Dim change As Double
Dim previousamount As Double
Dim TableRow As Double
Dim lastline As Double
Dim LastRow As Double
'counter of the variables

TableRow = 2
previousamount = 2
totalvolume = 0
lastline = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Loop to get the values (2)
For i = 2 To lastline
totalvolume = totalvolume + ws.Cells(i, 7).Value

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
TickerLetter = ws.Cells(i + 1, 1).Value
ws.Cells(TableRow, 9).Value = TickerLetter
nepo = ws.Cells(previousamount, 3)
esolc = ws.Cells(i, 6)
change = esolc - nepo
ws.Cells(TableRow, 10) = change
ws.Cells(TableRow, 12).Value = totalvolume

'IF (2) here
    If nepo = 0 Then
    percentinperfect = 0
    Else
    percentinperfect = change / nepo
    ws.Cells(TableRow, 11).NumberFormat = ".00%"
    ws.Cells(TableRow, 11).Value = percentinperfect
    End If 'End if (2)
    
    'if(3) here
    If ws.Cells(TableRow, 10).Value >= 0 Then
    ws.Cells(TableRow, 10).Interior.Color = RGB(0, 255, 0)
    Else
    ws.Cells(TableRow, 10).Interior.Color = RGB(255, 0, 0)
    
    End If 'End if (3)
    'End if (4)
    TableRow = TableRow + 1
    previousamount = i + 1
    totalvolume = 0
    End If
    
    Next i 'aaaaaaaaa
    
    
    LastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
            For i = 2 To LastRow
                If ws.Cells(i, 12).Value > ws.Range("P4").Value Then
                    ws.Range("P4").Value = ws.Cells(i, 12).Value
                    ws.Range("O4") = ws.Cells(i, 9).Text
                End If

                If ws.Cells(i, 10).Value > ws.Range("P2").Value Then
                    ws.Range("P2").Value = ws.Cells(i, 12).Value
                    ws.Range("O2") = ws.Cells(i, 9).Text
                End If

                If ws.Cells(i, 10).Value < ws.Range("P3").Value Then
                    ws.Range("P3").Value = ws.Cells(i, 13).Value
                    ws.Range("O3") = ws.Cells(i, 9).Text
                End If

            Next i
            ws.Range("P2").NumberFormat = "0.00%"
            ws.Range("P3").NumberFormat = "0.00%"
            
    Next   'end of the sheet loop
End Sub

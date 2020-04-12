Sub VBAStocks():
'Cycle through each worksheet in the workbook
For Each ws In Worksheets
    
    'Define variables
    Dim Company As String
    Dim TotalVolume As Double
    TotalVolume = 0
    
    Dim FirstOpenPrice As Double
    FirstOpenPrice = Cells(2, 3).Value
    Dim LastClosePrice As Double
    
    Dim ChangeInPrice As Double
    Dim ChangeInPercent As Double
    ChangeInPrice = 0
    ChangeInPercent = 0
        
    Dim SummaryTableRow As Integer
    SummaryTableRow = 2
    
    'Add Headers on the Summary Table
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
        
    'Find the row number of the last row on the list
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Cycle through each row in the worksheet 
    For i = 2 To lastrow

        'Check if the current ticker is the same as the next ticker
        If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1) Then
            'If true, add the current Volume to the total Volume
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value
        Else
            'If not true, add the current ticker symbol to the Summary Table
            Company = ws.Cells(i, 1).Value
            ws.Cells(SummaryTableRow, 9).Value = Company
            
            'Print the current Total Volume on the Summary Table
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            ws.Cells(SummaryTableRow, 12).Value = TotalVolume
                                    
            'Calculate the change in price and put on the Summary Table
            LastClosePrice = ws.Cells(i, 6).Value
            ChangeInPrice = LastClosePrice - FirstOpenPrice
            ws.Cells(SummaryTableRow, 10).Value = ChangeInPrice

            'Format the change in price red if negative, greeen if positive 
            If ChangeInPrice < 0 Then
                ws.Cells(SummaryTableRow, 10).Interior.ColorIndex = 3
            Else: ws.Cells(SummaryTableRow, 10).Interior.ColorIndex = 4
            End If
            
            'Calculate the percent change and put on the Summary Table
            If FirstOpenPrice = 0 Then
                ChangeInPercent = 0
            Else
                ChangeInPercent = ChangeInPrice / FirstOpenPrice
                ws.Cells(SummaryTableRow, 11).Value = ChangeInPercent
                ws.Cells(SummaryTableRow, 11).NumberFormat = "0.00%"
            End If
            
            'Set the opening price of the next ticker
            FirstOpenPrice = ws.Cells(i + 1, 3).Value
            
            'Advance to the next row of the Summary Table
            SummaryTableRow = SummaryTableRow + 1
            
            'Reset the Total Volume for the next ticker
            TotalVolume = 0
        End If
    
    Next i
    
    'Add header and row titles for the Bonus Table
    ws.Range("n2").Value = "Greatest % Increase"
    ws.Range("n3").Value = "Greatest % Decrease"
    ws.Range("n4").Value = "Greatest Total Volume"
    ws.Range("o1").Value = "Ticker"
    ws.Range("p1").Value = "Value"
      
    'Determine the row number of the last row of the Summary Table
    lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    MaxIncrease = 0
    MaxDecrease = 0
    MaxVolume = 0
        
    For i = 2 To lastrow
        'Find the ticker with highest % change in price
        If MaxIncrease < ws.Cells(i, 11).Value Then
            MaxIncrease = ws.Cells(i, 11).Value
                        
            TickerOfMaxIncrease = ws.Cells(i, 9).Value
        End If
        
        'Find the ticker with lowest % change in price
        If MaxDecrease > ws.Cells(i, 11).Value Then
            MaxDecrease = ws.Cells(i, 11).Value
                        
            TickerOfMaxDecrease = ws.Cells(i, 9).Value
        End If
        
        'Find the ticker with highest trading volume
        If MaxVolume < ws.Cells(i, 12).Value Then
            MaxVolume = ws.Cells(i, 12).Value
            
            TickerOfMaxVolume = ws.Cells(i, 9).Value
        End If
              
    Next i
    
    'Place the ticker and appropriate values in the Bonus Table
    ws.Range("o2").Value = TickerOfMaxIncrease
    ws.Range("p2").Value = MaxIncrease
    ws.Range("p2").NumberFormat = "0.00%"
    ws.Range("o3").Value = TickerOfMaxDecrease
    ws.Range("p3").Value = MaxDecrease
    ws.Range("p3").NumberFormat = "0.00%"
    ws.Range("o4").Value = TickerOfMaxVolume
    ws.Range("p4").Value = MaxVolume

Next ws

MsgBox ("Stock Ticker Analysis Complete.")

End Sub
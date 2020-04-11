Sub VBAStocks():

For Each ws In Worksheets
'Create ticker variable for ticker, Yearly change, Percent Change
    
    
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
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    
'read ticker symbol column
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    
    For i = 2 To lastrow
        
        If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1) Then
        
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value
        
        Else
            Company = ws.Cells(i, 1).Value
            ws.Cells(SummaryTableRow, 9).Value = Company
            
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            ws.Cells(SummaryTableRow, 12).Value = TotalVolume
                                    
            LastClosePrice = ws.Cells(i, 6).Value
            ChangeInPrice = LastClosePrice - FirstOpenPrice
            ws.Cells(SummaryTableRow, 10).Value = ChangeInPrice
            If ChangeInPrice < 0 Then
                ws.Cells(SummaryTableRow, 10).Interior.ColorIndex = 3
            Else: ws.Cells(SummaryTableRow, 10).Interior.ColorIndex = 4
            End If
            
            If FirstOpenPrice = 0 Then
                ChangeInPercent = 0
                
            Else
                ChangeInPercent = ChangeInPrice / FirstOpenPrice
                ws.Cells(SummaryTableRow, 11).Value = ChangeInPercent
                ws.Cells(SummaryTableRow, 11).NumberFormat = "0.00%"
            End If
            
            FirstOpenPrice = ws.Cells(i + 1, 3).Value
            
            SummaryTableRow = SummaryTableRow + 1
            TotalVolume = 0
        End If
    
    Next i
    
    ws.Range("n2").Value = "Greatest % Increase"
    ws.Range("n3").Value = "Greatest % Decrease"
    ws.Range("n4").Value = "Greatest Total Volume"
    ws.Range("o1").Value = "Ticker"
    ws.Range("p1").Value = "Value"
    
    
    
    lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    MaxIncrease = 0
    MaxDecrease = 0
    MaxVolume = 0
        
    For i = 2 To lastrow
    
        If MaxIncrease < ws.Cells(i, 11).Value Then
            MaxIncrease = ws.Cells(i, 11).Value
                        
            TickerOfMaxIncrease = ws.Cells(i, 9).Value
            
        End If
        
        If MaxDecrease > ws.Cells(i, 11).Value Then
            MaxDecrease = ws.Cells(i, 11).Value
                        
            TickerOfMaxDecrease = ws.Cells(i, 9).Value
            
        End If
        
        If MaxVolume < ws.Cells(i, 12).Value Then
            MaxVolume = ws.Cells(i, 12).Value
            
            TickerOfMaxVolume = ws.Cells(i, 9).Value
            
        End If
              
    Next i
    
    ws.Range("o2").Value = TickerOfMaxIncrease
    ws.Range("p2").Value = MaxIncrease
    ws.Range("p2").NumberFormat = "0.00%"
    ws.Range("o3").Value = TickerOfMaxDecrease
    ws.Range("p3").Value = MaxDecrease
    ws.Range("p3").NumberFormat = "0.00%"
    ws.Range("o4").Value = TickerOfMaxVolume
    ws.Range("p4").Value = MaxVolume
   

    
'while ticker symbol is the same, add volume
'take first open price
'take last close price
'solve for yearly change by finding the difference between last close and first open prices
'solve for percent change by dividing yearly change by first open price
Next ws

MsgBox ("Stock Ticker Analysis Complete.")


End Sub


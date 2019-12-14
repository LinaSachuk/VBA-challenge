Sub StocksRealData()
    
Dim ws As Worksheet
Dim WorksheetName As String
Dim Ticker_Name As String


Total_Stock_Volume = 0

For Each ws In Worksheets
    
    WorksheetName = ws.Name
    Summary_Table_Row = 2
    RowCounter = 0
    
     ws.Range("I1").Value = "Ticker"
     ws.Range("J1").Value = "Yearly Changes"
     ws.Range("K1").Value = "Percent Change"
     ws.Range("L1").Value = "Total Stock Volume"
     
     ws.Range("P1").Value = "Ticker"
     ws.Range("Q1").Value = "Volume"
     ws.Range("O2").Value = "Greatest % Increase"
     ws.Range("O3").Value = "Greatest % Decrease"
     ws.Range("O4").Value = "Greatest Total Volume"
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    

    For i = 2 To LastRow
        
         If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                
                 'clear the data
                ws.Range("I" & Summary_Table_Row).Value = ""
                ws.Range("J" & Summary_Table_Row).Value = ""
                ws.Range("K" & Summary_Table_Row).Value = ""
                ws.Range("L" & Summary_Table_Row).Value = ""
                
                Ticker_Name = ws.Cells(i, 1).Value
                
                 ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
                 
                FirstRowForTicker = i - RowCounter
               
                
                Yearly_Changes = Round(ws.Cells(i, 6).Value - ws.Cells(FirstRowForTicker, 3).Value, 2)
                
                ws.Range("J" & Summary_Table_Row).Value = Yearly_Changes ' /// close - open
                
                
                If Yearly_Changes >= 0 Then
                    
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4

                Else
                    
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                    
                End If
                
                
                
                If ws.Cells(FirstRowForTicker, 3).Value = 0 Then
                    Percent_Change = 0
                Else
                    Percent_Change = Round(Yearly_Changes / ws.Cells(FirstRowForTicker, 3).Value, 4)
                End If
                
                
                ws.Range("K" & Summary_Table_Row).Value = Percent_Change  '/// year_change / open * 100
                ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                
                
                
                ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume + ws.Cells(i, 7).Value   '/// Sum(vol) + last row
                
                
                RowCounter = 0
                
                Summary_Table_Row = Summary_Table_Row + 1
                
                Total_Stock_Volume = 0
            
         Else
         
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
                
                RowCounter = RowCounter + 1
                
         
         End If
    
    
    Next i


    
    ws.Range("K" & Summary_Table_Row).Style = "Percent"
    MaxIncrease = 0
    MaxDecrease = 0
    MaxTotalVolume = 0
    
    For j = 2 To Summary_Table_Row
           
           If ws.Range("K" & j).Value > MaxIncrease Then
           
                Ticker = ws.Range("I" & j).Value
                MaxIncrease = ws.Range("K" & j).Value
                
           
           End If
           
           If ws.Range("K" & j).Value < MaxDecrease Then
           
                
                TickerD = ws.Range("I" & j).Value
                MaxDecrease = ws.Range("K" & j).Value
               
           
           End If
           
           If ws.Range("L" & j).Value > MaxTotalVolume Then
           
                
                TickerV = ws.Range("I" & j).Value
                MaxTotalVolume = ws.Range("L" & j).Value
                'Debug.Print MaxTotalVolume
           
           End If
           
           
    Next j
  
    ws.Range("P2").Value = Ticker
    ws.Range("Q2").Value = MaxIncrease
    ws.Range("Q2").NumberFormat = "0.00%"
    
    ws.Range("P3").Value = TickerD
    ws.Range("Q3").Value = MaxDecrease
    ws.Range("Q3").NumberFormat = "0.00%"
    
    ws.Range("P4").Value = TickerV
    ws.Range("Q4").Value = MaxTotalVolume
    
Next ws



End Sub



Sub StockMarketAnalysis()


    'Loop for worksheets
        For Each ws In Worksheets
      
    'Declare variables
        Dim Ticker As String
        Dim Yearly_Change As Double
        Dim Percent_Change As Double
        Dim TotalVolume As Double
            TotalVolume = 0
        Dim LastRow As Long
        Dim Summary_Table_Row As Integer
            Summary_Table_Row = 2
        Dim Amount As Long
            Amount = 2
     
     
    'Coloumn headers
        ws.Range("i1").Value = "Ticker"
        ws.Range("j1").Value = "Yearly Change"
        ws.Range("k1").Value = "Percent Change"
        ws.Range("l1").Value = "Total Stock Volume"
     
    'Assign LastRow
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        
    For i = 2 To LastRow
    
        TotalVolume = TotalVolume + ws.Cells(i, 7).Value
          
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
       
           
                'Set Ticker
                    Ticker = Cells(i, 1).Value
                    
                'Print Ticker into the SummaryTable
                    ws.Range("i" & Summary_Table_Row).Value = Ticker
                    
                'Print TotalVolume to the SummaryTable
                   
                    ws.Range("l" & Summary_Table_Row).Value = TotalVolume
                    
                'TotalVolume reset
                     TotalVolume = 0
                     
                'Set StockOpen, StockClose and YearlyChange
                     StockOpen = ws.Range("c" & Amount).Value
                     StockClose = ws.Range("F" & i)
                     YearlyChange = StockClose - StockOpen
                     ws.Range("J" & Summary_Table_Row).Value = YearlyChange
                     
                'Percent Change
                    If StockOpen = 0 Then
                       PercentChange = 0
                    Else
                        StockOpen = ws.Range("c" & Amount)
                        PercentChange = YearlyChange / StockOpen
                    End If
                    
                        'Including % symbol
                        ws.Range("k" & Summary_Table_Row).NumberFormat = "0.00%"
                        ws.Range("k" & Summary_Table_Row).Value = PercentChange
                        
                        
                 'Conditional formatting
                    If ws.Range("J" & Summary_Table_Row).Value >= 0 Then
                       ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                    Else
                        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                    End If
                    
                    Summary_Table_Row = Summary_Table_Row + 1
                    
                    Amount = i + 1
                     
                                  
                             
          
           End If
            
     Next i
     
 Next ws
            
             
           
End Sub



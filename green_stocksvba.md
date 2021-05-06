Sub DQAnalysis()
    
    Worksheets("DQ Analysis").Activate
    
    Range("A1").Value = "DAQO (Ticker: DQ)"
    
    'Create a header row
    Cells(3, 1).Value = "Year"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
    Worksheets("2018").Activate
    
    'set intial volume to zero
    totalVolume = 0
    
    Dim startingPrice As Double
    Dim endingPrice As Double
    
    'find the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    'loop over all the rows
    For i = 2 To RowCount
    
        
        If Cells(i, 1).Value = "DQ" Then
            
            totalVolume = totalVolume + Cells(i, 8).Value
        
        End If
        
        If Cells(i - 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then
    
            startingPrice = Cells(i, 6).Value
        
        End If
        If Cells(i + 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then
        
            endingPrice = Cells(i, 6).Value
            
        End If
    
        
    Next i
    
    
    Worksheets("DQ Analysis").Activate
    Cells(4, 1).Value = 2018
    Cells(4, 2).Value = totalVolume
    Cells(4, 3).Value = (endingPrice / startingPrice) - 1
        
  
End Sub


Sub AllStocksAnalysis()

    Worksheets("All Stocks Analysis").Activate
    
        Range("A1").Value = "All Stocks(2018)"
    
        Cells(3, 1).Value = "Ticker"
        Cells(3, 2).Value = "Total Daily Volume"
        Cells(3, 3).Value = "Return"

    Dim tickers(12) As String
        tickers(0) = "AY"
        tickers(1) = "CSIQ"
        tickers(2) = "DQ"
        tickers(3) = "ENPH"
        tickers(4) = "FSLR"
        tickers(5) = "HASI"
        tickers(6) = "JKS"
        tickers(7) = "RUN"
        tickers(8) = "SEDG"
        tickers(9) = "SPWR"
        tickers(10) = "TERP"
        tickers(11) = "VSLR"
        
    Dim startingPrice As Single
    Dim endingPrice As Single
    
    Worksheets("2018").Activate
    
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
        For i = 0 To 11
    
            ticker = tickers(i)
            totalVolume = 0
            
            Worksheets("2018").Activate
            
            For j = 2 To RowCount
            
            If Cells(j, 1).Value = ticker Then
                
                totalVolume = totalVolume + Cells(j, 8).Value
            
                End If
            
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            
                startingPrice = Cells(j, 6).Value
                
                End If
    
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            
                endingPrice = Cells(j, 6).Value
                
                End If
                
            Next j
            
            Worksheets("All Stocks Analysis").Activate
            Cells(4 + i, 1).Value = ticker
            Cells(4 + i, 2).Value = totalVolume
            Cells(4 + i, 3).Value = (endingPrice / startingPrice) - 1
            
        Next i
    
End Sub

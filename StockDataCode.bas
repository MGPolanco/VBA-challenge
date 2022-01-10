Attribute VB_Name = "Module11"
Sub stock_market()

        Dim wss As Worksheet
        Dim TotalVolume As Double
        Dim Ticker As String
        Dim TickerNext As String
        Dim OpenPrice As Double
        Dim ClosePrice As Double
        Dim YearlyChange As Double
        Dim PercentChange As Double
        For Each wss In Worksheets
        Dim Percent As String
        
            wss.Range("j1").Value = "Ticker"
            wss.Range("k1").Value = "Yearly Change"
            wss.Range("l1").Value = "Percent Change"
            wss.Range("m1").Value = "Total Stock Volume"
            
            
            row_count = Cells(Rows.Count, "A").End(xlUp).Row
    
           
            PositionTicker = 2
            OpenPrice = wss.Cells(2, 3).Value
            'PercentChange = YearlyChange / OpenPrice
           
           
           
          
            For i = 2 To row_count
            
            vol = wss.Cells(i, 7).Value
            TickerNext = wss.Cells(i + 1, 1).Value
            If Ticker = TickerNext Then
            TotalVolume = TotalVolume + vol
            Else
            TotalVolume = TotalVolume + vol
            
            End If
            
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                Ticker = wss.Cells(i, 1).Value
                vol = wss.Cells(i, 7).Value
                TickerNext = wss.Cells(i + 1, 1).Value
                'TotalVolume = 0
                ClosePrice = wss.Cells(i, 6).Value
                YearlyChange = ClosePrice - OpenPrice
                
                
            If wss.Cells(PositionTicker, 11).Value < 0 Then
                wss.Cells(PositionTicker, 11).Interior.ColorIndex = 3
            Else
                wss.Cells(PositionTicker, 11).Interior.ColorIndex = 4
            End If
            
            
            If OpenPrice = 0 Then
                PercentChange = 0
            Else
                PercentChange = YearlyChange / OpenPrice
                
                wss.Cells(PositionTicker, 12).Value = Format(PerChange, "Percent")
                wss.Cells(PositionTicker, 12).Value = Format(0, "Percent")
                
            End If
                
                OpenPrice = wss.Cells(i + 1, 3).Value
                Ticker = TickerNext
                    TotalVolume = TotalVolume + vol
                    
                    wss.Cells(PositionTicker, 10).Value = Ticker
                    wss.Cells(PositionTicker, 13).Value = TotalVolume
                    wss.Cells(PositionTicker, 11).Value = YearlyChange
                    wss.Cells(PositionTicker, 12).Value = PercentChange
                    PositionTicker = PositionTicker + 1
                    TotalVolume = 0
                    
            
           
            
                End If
                    
            Next i
                    
        Next wss
                    
End Sub









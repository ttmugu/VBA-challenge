
Sub VBAStocks()

    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    For Each ws In Worksheets
    
        Dim WorksheetName As String
        'variable
        Dim ticker_symbol As String
        Dim opening_price As Double
        Dim closing_price As Double
        Dim highst_price As Double
        Dim lowest_price As Double
        Dim stock_volume As Double
        Dim PercentChange As Double
        
        ' Grabbed the WorksheetName
        WorksheetName = ws.Name
         
        ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        ' Determine the Last Column Number
        LastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
        
        Dim j As Integer
        Dim k As Integer
            j = 2
            k = 2
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock volume"
        
        For i = 2 To LastRow
             'check below column has same tiker
             If (ws.Range("A" & i) = ws.Range("A" & (i + 1))) Then
                'check Ticker change in next column or not
                ticker_symbol = ws.Range("A" & (i + 1)).Value
                If (ws.Range("A" & (i - 1)) <> ws.Range("A" & i)) Then
                    opening_price = ws.Range("C" & i).Value
                End If
                
                stock_volume = stock_volume + ws.Range("G" & i).Value
               
            Else
                closing_price = ws.Range("F" & i).Value
                stock_volume = stock_volume + ws.Range("G" & i).Value
                'opening price zero handling
                If (opening_price <> 0) Then
                    PercentChange = (closing_price - opening_price) / opening_price
                Else
                    PercentChange = 0
                    
                End If
                    
                
                'write in cells
                ws.Range("I" & j).Value = ticker_symbol
                
                ws.Range("J" & j).Value = (closing_price - opening_price)
                
                ws.Range("K" & j).Value = PercentChange
                ws.Range("K" & j).NumberFormat = "0.00%"
                
                ws.Range("L" & j).Value = stock_volume
                ' increse counter
                j = j + 1
                k = k + 1
                                
                'reset values
                closing_price = 0
                opening_price = 0
                closing_price = 0
                stock_volume = 0
                PercentChange = 0
                
                
            End If
            
        Next i
 
        'find last row in calculated column
        lRow = ws.Cells(Rows.Count, 10).End(xlUp).Row
        
        Dim high As Double
        Dim rowHigh As Integer
        Dim low As Double
        Dim rowLow As Integer
        Dim hVolume As Double
        Dim hVRow As Integer
        
        'change cell color based on cell value
        For Each Cell In ws.Range("J2:J" & lRow)
            If Cell.Value > 0 Then
                Cell.Interior.ColorIndex = 4
            Else
                Cell.Interior.ColorIndex = 3
            End If
         Next Cell
         
        'find high and low % change
        ws.Range("P1") = "Ticker"
        ws.Range("Q1") = "Value"
        ws.Range("O2").Value = "Gretest%Increase"
        ws.Range("O3").Value = "Gretest%Decrease"
        ws.Range("O4").Value = "Gretest Total Volume"
            
        For Each Cell In ws.Range("K2:K" & lRow)
            If high < Cell.Value Then
                high = Cell.Value
                rowHigh = Cell.Row()
            End If
            If low > Cell.Value Then
                low = Cell.Value
                rowLow = Cell.Row()
            End If
         Next Cell
         
         ws.Range("P2").Value = ws.Range("I" & rowHigh).Value
         ws.Range("Q2").Value = high
         ws.Range("Q2").NumberFormat = "#0.00%"
         ws.Range("P3").Value = ws.Range("I" & rowLow).Value
         ws.Range("Q3").Value = low
         ws.Range("Q3").NumberFormat = "#0.00%"
       
         'find highest volumn in a year
        For Each Cell In ws.Range("L2:L" & lRow)
            If hVolume < Cell.Value Then
                hVolume = Cell.Value
                hVRow = Cell.Row()
            End If
        Next Cell
        
        ws.Range("P4").Value = ws.Range("I" & hVRow).Value
        ws.Range("Q4").Value = hVolume
         
        'reset value before next calculation
        high = 0
        rowHigh = 0
        low = 0
        rowLow = 0
        
        hVolume = 0
        hVRow = 0
        lRow = 0
        
     'got to next worksheet
    Next ws
    
End Sub



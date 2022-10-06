# VBA-Challenge
Sub AlphabeticalTesting()
    
    'Name Worksheet
        Dim ws As Worksheet
        
    'Loop Worksheets
        For Each ws In Worksheets
        
    'Define variables
        Dim Ticker As String
        
        Dim Year_Open As Double
        
        Dim Year_Close As Double
        
        Dim Yearly_Change As Double
        
        Dim Total_Stock_Volume As Double
        
        Dim Percent_Change As Double
        
        Dim Start_Data As Integer
        
    'Create Column Headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
    'Define EndTicker
    
        EndTicker = ws.Cells(Rows.Count, 1).End(x1Up).Row
        
    'Loop Rows
        For i = 2 To EndTicker
        
            'Ticker Name
                TickerName = ws.Cells(i, 1).Value
        
            'If Ticker Name changes
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Ticker = ws.Cells(i, 1).Value
                
    End If
    
                
            'Calculate Yearly Change in Dollars and Format
                Yearly_Change = ClosePrice - OpenPrice
                OpenPrice = ws.Range(i, 3).Value
                ClosePrice = ws.Range(i, 6).Value
                ws.Range(10).NumberFormat = "0.00"
                    If ws.Range(10).Value >= 0 Then
                    ws.Range(10).Interior.ColorIndex = 4
                        Else
                            ws.Range(10).Interior.ColorIndex = 3
                            
        
                            
                
                
        
    End If
    
Next ws

        
End Sub

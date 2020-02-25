Attribute VB_Name = "Module1"
Sub VBAWSHW():

    For Each ws In Worksheets
        ' Label headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "YearlyChange"
        ws.Range("K1").Value = "%Change"
        ws.Range("L1").Value = "TotalVolume"
        ws.Range("O2").Value = "Greatest%Increase"
        ws.Range("O3").Value = "Greatest%Decrease"
        ws.Range("O4").Value = "GreatestTotalVolume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
    
        ' Declare vars
        Dim Ticker As String
        
        Dim OpeningPriceYR As Double
        Dim ClosingPriceYR As Double
        Dim YearlyChange As Double
        Dim PercentChange As Double
        Dim TotalVolume As Double
        Dim GreatestPercentIncrease As Double
        Dim GreatestPercentDecrease As Double
        Dim GreatestTotalVolume As Double
        
        Dim ResultRow As Long
        'start resultrow in row 2 to avoid header
        ResultRow = 2
        Dim FinalRow As Long
        Dim FinalRowValue As Long
        
        'Find yearly closing
        FinalRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For q = 2 To FinalRow
            
            TotalVolume = TotalVolume + ws.Cells(q, 7).Value
            If ws.Cells(q + 1, 1).Value <> ws.Cells(q, 1).Value Then
            'continue if ticket matches otherwise break
            
            Ticker = ws.Cells(q, 1).Value
            ws.Range("I" & ResultRow).Value = Ticker
            ws.Range("L" & ResultRow).Value = TotalVolume
            'reset volume counter
            TotalVolume = 0
    
            'Calculate yearly statistics.... our for loop and ws loop are still open'
            OpeningPriceYR = 0
            ClosingPriceYR = 0
            '^ needed for the 0 case to avoid DBZ error
            OpeningPriceYR = ws.Range("C" & ResultRow)
            ClosingPriceYR = ws.Range("F" & q)
            YearlyChange = ClosingPriceYR - OpeningPriceYR
            ws.Range("J" & ResultRow).Value = YearlyChange
            
            If OpeningPriceYR = 0 Then
                PercentChange = 0
            Else
                OpeningPriceYR = ws.Range("C" & ResultRow)
                'calculate % change
                PercentChange = YearlyChange / OpeningPriceYR
            End If
            ws.Range("K" & ResultRow).Value = PercentChange
            'determine color based on yearly gain or loss'
            If ws.Range("J" & ResultRow).Value >= 0 Then
                ws.Range("J" & ResultRow).Interior.ColorIndex = 4
            ElseIf ws.Range("J" & ResultRow).Value <= 0 Then
                ws.Range("J" & ResultRow).Interior.ColorIndex = 3
            Else: ws.Range("J" & ResultRow).Interior.ColorIndex = 2
            End If
            'increment result row for new ticker'
            ResultRow = ResultRow + 1
            End If
            
        Next q
    
    FinalRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
    
        ' Loop for final results
        For q = 2 To FinalRow
        
            If ws.Range("K" & q).Value > ws.Range("Q2").Value Then
                ws.Range("Q2").Value = ws.Range("K" & q).Value
                ws.Range("P2").Value = ws.Range("I" & q).Value
            End If
            
            If ws.Range("K" & q).Value < ws.Range("Q3").Value Then
                ws.Range("Q3").Value = ws.Range("K" & q).Value
                ws.Range("P3").Value = ws.Range("I" & q).Value
            End If
            
            If ws.Range("L" & q).Value > ws.Range("Q4").Value Then
                ws.Range("Q4").Value = ws.Range("L" & q).Value
                ws.Range("P4").Value = ws.Range("I" & q).Value
            End If
            
        Next q
    
    Next ws

End Sub


Attribute VB_Name = "Module1"
Sub QuarterlyAnalysis():
    
    For Each Sheet In Worksheets
    
        'Set Dimensions
        Dim WorksheetName As String
        'Current row
        Dim i As Long
        'Start row of ticker
        Dim j As Long
        'Ticker count
        Dim Tickercount As Long
        'Variable for last row column A
        Dim FinalrowA As Long
        'Variable for last row column I
        Dim FinalrowI As Long
        'Variable for Percent change
        Dim PercentChange As Double
        'Variable for greatest Increase
        Dim GreatInc As Double
        'Variable for greatest decrease
        Dim GreatDec As Double
        'Variable for greatest stock volume
        Dim GreatStockVol As Double
    
        'Worksheet Name
        WorksheetName = Sheet.Name
    
        'Create Column titles
        Sheet.Cells(1, 9).Value = "Ticker"
        Sheet.Cells(1, 10).Value = "Quarterly Change"
        Sheet.Cells(1, 11).Value = "Percent Change"
        Sheet.Cells(1, 12).Value = "Total Stock Volume"
        Sheet.Cells(2, 15).Value = "Greatest %Increase"
        Sheet.Cells(3, 15).Value = "Greatest %Decrease"
        Sheet.Cells(4, 15).Value = "Greatest Total Stock Volume"
        Sheet.Cells(1, 16).Value = "Ticker"
        Sheet.Cells(1, 17).Value = "Value"
    
        'Column width auto-adjust
        Worksheets(WorksheetName).Columns("A:R").AutoFit
    
     
        'Set first row Ticker counter
        Tickercount = 2
    
        'Set start row to 2
        j = 2
    
        'Locate the last content cell in Col A
        FinalrowA = Sheet.Cells(Rows.Count, 1).End(xlUp).Row
    
            'Loop through all rows
            For i = 2 To FinalrowA
            
                'Ticker name change
                If Sheet.Cells(i + 1, 1).Value <> Sheet.Cells(i, 1).Value Then
            
                'Write ticker for Column I (cell9)
                Sheet.Cells(Tickercount, 9).Value = Sheet.Cells(i, 1).Value
            
                'Calculate and write Quarterly Change in column J (cell10)
                Sheet.Cells(Tickercount, 10).Value = Sheet.Cells(i, 6).Value - Sheet.Cells(j, 3).Value
                
                    'Conditional Formatting
                    If Sheet.Cells(Tickercount, 10).Value < 0 Then
                
                    'Set cell background color to red
                    Sheet.Cells(Tickercount, 10).Interior.ColorIndex = 3
                    
                    ElseIf Sheet.Cells(Tickercount, 10).Value > 0 Then
                
                    'Set cell background color to green
                    Sheet.Cells(Tickercount, 10).Interior.ColorIndex = 4
                
                    End If
                    
                    'Calculate and write percent change in column K (cell11)
                    If Sheet.Cells(j, 3).Value <> 0 Then
                    PercentChange = ((Sheet.Cells(i, 6).Value - Sheet.Cells(j, 3).Value) / Sheet.Cells(j, 3).Value)
                    
                    'Percentage change
                    Sheet.Cells(Tickercount, 11).Value = Format(PercentChange, "Percent")
                    
                    Else
                    
                    Sheet.Cells(Tickercount, 11).Value = Format(0, "Percent")
                    
                    End If
                    
                'Calculate and write total volume in column L (cell12)
                Sheet.Cells(Tickercount, 12).Value = WorksheetFunction.Sum(Range(Sheet.Cells(j, 7), Sheet.Cells(i, 7)))
                
                'Increase Tickercount by 1
                Tickercount = Tickercount + 1
                
                'Set new start row of the ticker block
                j = i + 1
                
                End If
            
            Next i
        
        'Locate last content cell in column I
        FinalrowI = Sheet.Cells(Rows.Count, 9).End(xlUp).Row
        
        'Prepare for summary
        GreatStockVol = Sheet.Cells(2, 12).Value
        GreatInc = Sheet.Cells(2, 11).Value
        GreatDec = Sheet.Cells(2, 11).Value
        
            'Loop for summary
            For i = 2 To FinalrowI
            
                'For greatest stock total volume, check if next value is larger if yes take over a new value and populate Sheet.Cells
                If Sheet.Cells(i, 12).Value > GreatStockVol Then
                GreatStockVol = Sheet.Cells(i, 12).Value
                Sheet.Cells(4, 16).Value = Sheet.Cells(i, 9).Value
                
                Else
                
                GreatStockVol = GreatStockVol
                
                End If
                
                'For greatest increase--check if next value is larger--if yes take over a new value and populate ws.Cells
                If Sheet.Cells(i, 11).Value > GreatInc Then
                GreatInc = Sheet.Cells(i, 11).Value
                Sheet.Cells(2, 16).Value = Sheet.Cells(i, 9).Value
                
                Else
                
                GreatInc = GreatInc
                
                End If
                
                'For greatest decrease, check if next value is smaller, if yes take over a new value and populate Sheet.Cells
                If Sheet.Cells(i, 11).Value < GreatDec Then
                GreatDec = Sheet.Cells(i, 11).Value
                Sheet.Cells(3, 16).Value = Sheet.Cells(i, 9).Value
                
                Else
                
                GreatDec = GreatDec
                
                End If
                
            'Write summary results in Sheet.Cells
            Sheet.Cells(2, 17).Value = Format(GreatInc, "Percent")
            Sheet.Cells(3, 17).Value = Format(GreatDec, "Percent")
            Sheet.Cells(4, 17).Value = Format(GreatStockVol, "Scientific")
        
            Next i

        Next Sheet

End Sub




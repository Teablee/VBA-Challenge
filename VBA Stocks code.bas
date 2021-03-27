Attribute VB_Name = "Module1"
Sub ws_loop()
'Make loop through worksheets
    Dim Current As Worksheet
    
    For Each Current In Worksheets
        
        'MsgBox Current.Name check
        
            'Define variables
        Dim Ticker As String
        Dim Yearly_change As Double
        Dim Percent_change As Double
        Dim Volume As Double
        Dim summary_row As Long
        summary_row = 2
        Dim start As Long
        start = 2
        
          'Make summary table
        Current.Range("J1").Value = "Ticker"
        Current.Range("K1").Value = "Yearly_Change"
        Current.Range("L1").Value = "Percent_Change"
        Current.Range("M1").Value = "Total_Volume"
        Current.Range("J1:M1").Font.Bold = True

        'Make LastRow check the rows for the sheet
        Dim LastRow As Long
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row
        
        'MsgBox LastRow check
        
            'Loop through tickers
        For i = 2 To LastRow
                
                If Current.Cells(i, 1).Value <> Current.Cells(i + 1, 1) Then
                
                   'Define ticker cells
                    Ticker = Current.Cells(i, 1).Value
                   
                   'Define running total for total stock vol
                    Volume = Volume + Current.Cells(i, 7).Value
                   
                   'Define opening value as first <open> value displayed per ticket
                    Opening = Current.Cells(start, 3).Value
                   
                   'progress the start with ticker changing
                    start = i + 1
                    
                    'Define close value as last row
                    Closing = Current.Cells(i, 6).Value
                    
                    'Define formula for yearly change
                    Yearly_change = (Closing - Opening)
                    
                    'Change colors red for negative and green for positive
                    If Yearly_change < 0 Then
                        Current.Cells(summary_row, 11).Interior.ColorIndex = 3
                    Else
                        Current.Cells(summary_row, 11).Interior.ColorIndex = 4
                    End If
                   
                   'Define formula for percent change
                    'Percent_change = (Closing - Opening) / Opening
                    'Fix Problem with 0 opening value
                   If Opening <> 0 Then
                        Percent_change = (Closing - Opening) / Opening
                   ElseIf Opening = 0 Then
                        MsgBox ("Open value is zero for " + Ticker + ". Percent change cannot be calculated and will be set to 0 by default.")
                        Percent_change = 0
                   End If
                  
                   'Update summary table
                   Current.Cells(summary_row, 10).Value = Ticker
                   Current.Cells(summary_row, 11).Value = Yearly_change
                   Current.Cells(summary_row, 12).Value = Percent_change
                   Current.Cells(summary_row, 12).NumberFormat = ("0.00%")
                   Current.Cells(summary_row, 13).Value = Volume
                   Current.Columns("J:M").EntireColumn.AutoFit
                   
                   'Reset running total
                   Volume = 0
                   
                   'move to the next summary row
                   summary_row = summary_row + 1
                
                
                 Else
                    'running total for total volume
                    Volume = Volume + Current.Cells(i, 7).Value
                
                End If
        Next i
                              
         'Bonus
         
         'Make summary table
         Current.Cells(2, 15).Value = "Greatest % Increase"
         Current.Cells(3, 15).Value = "Greatest % Decrease"
         Current.Cells(4, 15).Value = "Greatest Total Volume"
         Current.Cells(1, 16).Value = "Ticker"
         Current.Cells(1, 17).Value = "Value"
         Current.Columns("O:Q").EntireColumn.AutoFit
         
         Dim Greatest_increase As Double
         
         Dim Greatest_decrease As Double
         
         Dim Greatest_total_volume As Double
         
         'Find greatest increase with large function
         Greatest_increase = Application.WorksheetFunction.Large(Current.Range("L:L"), 1)
         Current.Cells(2, 17).Value = Greatest_increase
         Current.Cells(2, 17).NumberFormat = ("0.00%")
         
         'Find greatest decrease with min function
         Greatest_decrease = Application.WorksheetFunction.Min(Current.Range("L:L"))
         Current.Cells(3, 17).Value = Greatest_decrease
         Current.Cells(3, 17).NumberFormat = ("0.00%")
         
         'Find greatest total volume with large function
         Greatest_total_volume = Application.WorksheetFunction.Large(Current.Range("M:M"), 1)
         Current.Cells(4, 17).Value = Greatest_total_volume
         
         'Index/Match to find the ticker symbol associated with each value
         Current.Range("P2") = "=Index(J:J, match(Q2,L:L,0))"
         Current.Range("P3") = "=Index(J:J, match(Q3,L:L,0))"
         Current.Range("P4") = "=Index(J:J, match(Q4,M:M,0))"

        
    Next Current

End Sub


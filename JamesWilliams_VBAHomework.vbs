Sub VBAHomework()
    For Each ws In Worksheets
        ws.Activate
        Call CalculateSummary
    Next ws
End Sub

Sub CalculateSummary()
    ' Start writing your code here

'setup dimensions for all variables used
Dim current_ticker As String
Dim next_ticker As String
Dim yearly_change As Double
Dim percent_change As Double
Dim total_stock_volume As Double
Dim open_price As Double
Dim close_price As Double



'initialize variables
total_stock_volume = 0
open_price = Cells(2, 3).Value

'initialize summary table first row
Dim summary_row As Integer
summary_row = 2

'calculate total rows of the sheet
Dim totalrows As Single
totalrows = Cells(Rows.Count, "A").End(xlUp).Row


'for loop from first row of data to the final row
    For currentrow = 2 To totalrows
        
        current_ticker = Cells(currentrow, 1).Value
        
        next_ticker = Cells(currentrow + 1, 1).Value
        
        total_stock_volume = total_stock_volume + Cells(currentrow, 7).Value
        
        'if statement when the next ticker appears
        If current_ticker <> next_ticker Then
            
            Cells(summary_row, 9).Value = current_ticker
            
            Cells(summary_row, 12).Value = total_stock_volume
            
            close_price = Cells(currentrow, 6).Value
            
            yearly_change = (close_price - open_price)
            
            Cells(summary_row, 10).Value = yearly_change
            
            Cells(summary_row, 10).NumberFormat = "0.00"
            
            'If statement for green/red formatting based on yearly_change
            If yearly_change >= 0 Then
            
                Cells(summary_row, 10).Interior.ColorIndex = 4
            
            Else
            
                Cells(summary_row, 10).Interior.ColorIndex = 3
            
            End If
            
            'If statement to fix divide by zero error
            If open_price = 0 And closing_price = 0 Then
                percent_change = 0
            ElseIf open_price = 0 And closing_price <> 0 Then
                percent_change = 1
            Else
            percent_change = (yearly_change / open_price)
            End If
        
            
            Cells(summary_row, 11) = percent_change
            
            Cells(summary_row, 11).NumberFormat = "0.00%"
            
            
            'move to the next row in the summary table
            summary_row = summary_row + 1
            
            'set variables back to zero for new ticker
            total_stock_volume = 0
            open_price = Cells(currentrow + 1, 3).Value

        End If
        
    Next currentrow
    
    
    
    
    
    
    Debug.Print ActiveSheet.Name
    Call SetTitle
End Sub


Sub SetTitle()
' Set title row
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    'this is for challenge only
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("I:O").Columns.AutoFit
End Sub


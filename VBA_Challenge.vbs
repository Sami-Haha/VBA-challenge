'VBA Challenge
Sub VBA_Challenge()
For Each ws In Worksheets
     'create variable to hold the ticker
    Dim ticker_symbol As String
    'create variable for the main row checking loop
    Dim row As LongLong
    'set location for each ticker in the summary table
    Dim Sum_Tab_R As Integer
    'set initial variable to hold volume
    Dim Total_stock As LongLong
    'create variables for storing dates
    Dim opendate As String
    Dim closedate As String
    'create variables to store new date
    Dim dateOne As String
    Dim date_value As String
    'create variables for storing the prices
    Dim O_Price As Double
    Dim C_Price As Double
    'create vaiables for storing the change and percent change
    Dim Y_Change As Double
    Dim P_Change As Double
    
    
        'define starting values for variables
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).row
        Sum_Tab_R = 2
        Total_stock = 0
        opendate = "0102"
        closedate = "1231"
        O_Price = 0
        C_Price = 0
        Y_Change = 0
        P_Change = 0
    
   'Apply conditional formatting and column headings
        myFormatting
   
    'Loop through all ticker symbols
     For row = 2 To LastRow
        'check for open and close dates
         dateOne = ws.Cells(row, 2).Value
         date_value = Right(dateOne, 4)
            ' get and store the open and close price values
            ' if year value equals open date then store open price in O_Price
                    If date_value = opendate Then
                    O_Price = ws.Cells(row, 3).Value
            'else if year value equals close date then store close price in C_Price
                    Else: date_value = closedate
                    C_Price = ws.Cells(row, 6).Value
            'else do nothing
                    End If


        'check if ticker is the same, if not do ...
            If ws.Cells(row + 1, 1).Value <> ws.Cells(row, 1).Value Then
                'set the ticker name
                ticker_symbol = ws.Cells(row, 1).Value
                'complete the total stock volume
                Total_stock = ws.Cells(row, 7).Value + Total_stock
                'complete yearly change value
                Y_Change = C_Price - O_Price
                'complete percentage change value
                P_Change = Y_Change / O_Price
                 'Printing the final results
                  'print the ticker into the summary table
                  ws.Range("I" & Sum_Tab_R).Value = ticker_symbol
                  'print the yearly change into the summary table
                  ws.Range("j" & Sum_Tab_R).Value = Y_Change
                  'print the percentage change into the summary table
                  ws.Range("k" & Sum_Tab_R).Value = P_Change
                  'print the total stock volume into the summary table
                  ws.Range("L" & Sum_Tab_R).Value = Total_stock
                    'color yearly cells for negative or positive value
                    If ws.Range("j" & Sum_Tab_R).Value > 0 Then
                        ws.Range("j" & Sum_Tab_R).Interior.ColorIndex = 4
                    ElseIf ws.Range("j" & Sum_Tab_R).Value < 0 Then
                        ws.Range("j" & Sum_Tab_R).Interior.ColorIndex = 3
                    End If
                    'color percentage cells for negative or positive value
                    If ws.Range("K" & Sum_Tab_R).Value > 0 Then
                        ws.Range("K" & Sum_Tab_R).Interior.ColorIndex = 4
                    ElseIf ws.Range("K" & Sum_Tab_R).Value < 0 Then
                        ws.Range("K" & Sum_Tab_R).Interior.ColorIndex = 3
                    End If

                'increment the summary table row number
                Sum_Tab_R = Sum_Tab_R + 1
                
                'reset all table variables
                O_Price = 0
                C_Price = 0
                Y_Change = 0
                P_Change = 0
                Total_stock = 0

            Else
                'add volume to total stock volume
                Total_stock = ws.Cells(row, 7).Value + Total_stock
            End If
      
       Next row
    Next ws

End Sub

Sub myFormatting():
    For Each ws In Worksheets
    'Column Headings
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percentage Change"
    ws.Range("L1").Value = "Total Stock Volume"
    'Column widths
    ws.Columns("I:K").AutoFit
    ws.Columns("L").ColumnWidth = 20
    'Column types
    ws.columns ("J").numberformat = "$0.00"
    ws.Columns("K").NumberFormat = "0.00%"
    Next ws
End Sub

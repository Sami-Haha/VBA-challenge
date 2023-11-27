sub bonus ():

    ' Loop through all sheets
    For Each ws In Worksheets
'create variable to hold the ticker symbols
    Dim tick_increase As String
    Dim tick_decrease As String
    Dim tick_Gtot As String
    'create variable for the main row checking loop
    Dim row As Long
  
    'set initial variable to hold volume
    Dim GreatestTotal As LongLong
    'create variables for greatest %increase and decrease
    Dim G_P_Increase As Double
    Dim G_P_Decrease As Double
    
    
    
        'define starting values for variables
        LastRow = ws.Cells(Rows.Count, 9).End(xlUp).row
        GreatestTotal = 0
        G_P_Increase = 0
        G_P_Decrease =0
        
    'loop through all values
    for row = 2 to LastRow
        'check % change for increase or decrease
        if ws.Cells(row , 11).value > G_P_Increase then
        'new value
        G_P_Increase = ws.Cells(row,11).value
        'new ticker value
        tick_increase = ws.Cells(row,9).value
        'else check if it is greatest % decrease
        ElseIf ws.Cells(row , 11).value < G_P_Decrease then
        'new value
        G_P_Decrease = ws.Cells(row , 11).value 
        'new ticker value
        tick_decrease = ws.Cells(row , 9).value 
        end if
        'next check for greatest total
        if ws.Cells(row,12).value > GreatestTotal then
        'new value
        GreatestTotal = ws.Cells(row,12).value 
        'new ticker value
        tick_Gtot = ws.Cells(row,9).value  
        end if
    next row

    'Output results of search
    myFormatting
    'Greatest % increase values
    ws.Range ("P2").value = tick_increase
    ws.Range ("Q2").value = G_P_Increase
    'Greatest % decearse values
    ws.Range ("P3").value = tick_decrease
    ws.Range ("Q3").value = G_P_Decrease
    'Greatest total volume
    ws.Range ("P4").value = tick_Gtot
    ws.Range ("Q4").value = GreatestTotal
  Next ws
  
end sub   
sub myFormatting ():
    for each ws in worksheets
    'Row Headings
    ws.Range ("O2").value = "Greatest % Increase"
    ws.Range ("O3").value = "Greatest % Decrease"
    ws.Range ("O4").value = "Greatest Total Volume"
    ws.Range ("P1").value = "Ticker"
    ws.Range ("Q1").value = "Value"
    'Column widths
    ws.Columns("O:P").Autofit
    ws.Columns("Q").ColumnWidth = 20
    'percentage format
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
    next ws
 end sub   
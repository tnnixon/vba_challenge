Attribute VB_Name = "Module1"
Sub runWorkbook()
    'run macro on multiple worksheets at same time 'extend office code
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call alphaTest
    Next
    Application.ScreenUpdating = True
End Sub
Sub alphaTest()

    'loop through all of the worksheets in a workbook
    For Each ws In Worksheets
    
    'variable to hold worksheet name
    Dim worksheetName As String
    
    'stores the name of the worksheet
    worksheetName = ws.Name
    
    'create variable to hold results
    Dim ticker As String
    Dim yearlyChng As Double
    Dim percentChng As Double
    'Dim totalStoVol As Long
    
    'variable to hold total stock volume
    totalStoVol = 0
    
    'variable to hold last row of the tickers
    Dim lastRow As Long
    
    'summary table row
    Dim summaryTableRow As Integer
    summaryTableRow = 2 'starts at row 2 in summary table
    
    'count the number of rows
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'loop through the rows in the ticker column
    For Row = 2 To lastRow
    
        'check to see of we are still within the same ticker
        'if not, do the following
        If (Cells(Row, 1).Value <> Cells(Row + 1, 1).Value) Then
        
        'set (reset) the ticker name
        ticker = Range("A" & Row).Value
        
        'add to the total stock volume one last time before change in ticker
        totalStoVol = totalStoVol + Range("G" & Row).Value
        
        'add the values to the summary table
        'add the ticker name to the I column on the summary table
        Range("I" & summaryTableRow).Value = ticker
        
        'add the final total stock volume to the L column on the summary table
        Range("L" & summaryTableRow).Value = totalStoVol
        
        'once the summary table is populated, then add one to the summary row count
        summaryTableRow = summaryTableRow + 1
        
        'then reset the total stock volume to 0
        totalStoVol = 0
        
        
        Else
        
        'if we are in the same ticker, add on to the running total
        totalStoVol = totalStoVol + Range("G" & Row).Value
        
        
        End If
        
    Next Row
    
    Next ws
    
End Sub

Attribute VB_Name = "Module1"
Sub Multi_Year_Stock():
    
    'Declare and set worksheet
    Dim ws As Worksheet
    For Each ws In Worksheets
    
    ' set values in column I to "Ticker"
    ws.Range("I1").Value = "Ticker"
    
    ' set values in column J to "Yearly change"
    ws.Range("J1").Value = "Yearly Change"
    
    ' set value in column K to "Percent Change"
    ws.Range("K1").Value = "Percent Change"
    
    ' set vlaue for column L to "Total Volume"
    ws.Range("L1").Value = "Total Volume"
    
    ' Set an intial variable for holding the ticker name
    Dim Ticker As String
    
    ' Set an intial variable for holding the opening price
    Dim OpeningPrice As Double
    OpeningPrice = 0
    
    ' Set an intial variable for holding the closing price
    Dim ClosingPrice As Double
    ClosingPrice = 0
    
    ' Set an intial variable for holding yearly change
    Dim YearlyChange As Double
    YearlyChange = 0
    
    ' Set an intial variable for holding the percent change
    Dim PercentChange As Double
    PercentChange = 0
    
    ' Set an intial variable for holding the total volume
    ' Set an intial variable for holding the total volume
    Dim TotalVolume As Double
    TotalVolume = 0
    
    ' Keep track of location for each new value in the summary table
    Dim Summary_Table_Row As Long
    Summary_Table_Row = 2
    
    ' Loop throught all the stock
    Dim Lastrow As Long
    
    'Count the # of rows
    Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' assign open date
    OpeningPrice = ws.Cells(2, 3).Value
    For i = 2 To Lastrow
    
    ' Set the ticker name
    
    'Check if all still within the same tickers,if it isnt,
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
    'Set the ticker
    Ticker = ws.Cells(i, 1).Value
    
    ' Set closingprice
    ClosingPrice = ws.Cells(i, 6).Value
    
    ' Add YearlyChange
    YearlyChange = ClosingPrice - OpeningPrice
    
    'Add to the PercentChange
    If OpeningPrice <> 0 Then
    PercentChange = (ClosingPrice - OpeningPrice) / OpeningPrice
    End If
    
    ' Add to the volume
    Volume = Volume + ws.Cells(i, 7).Value
    
    ' print the ticket symbol in the summary table
    ws.Range("I" & Summary_Table_Row).Value = Ticker
    ws.Range("J" & Summary_Table_Row).Value = YearlyChange
    
    ' conditonal format the positive and negative change
    If YearlyChange > 0 Then
    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
    
    Else
    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
    
    End If
    ' print percent change in column K
    ws.Range("K" & Summary_Table_Row).Value = PercentChange
     
    ' print total stock volume in column L
    ws.Range("L" & Summary_Table_Row).Value = Volume
    ws.Range("L" & Summary_Table_Row).Style = "Comma"
     
    Summary_Table_Row = Summary_Table_Row + 1
     
    OpeningPrice = ws.Cells(i + 1, 3).Value
    Volume = 0
    PercentChange = 0
     
    Else
       Volume = Volume + ws.Cells(i, 7).Value
    End If
    
    Next i
    
    Next ws
    
    For Each ws In ActiveWorkbook.Worksheets
        ws.Columns.AutoFit
    Next ws
End Sub

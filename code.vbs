Sub AlphaTest()

    For Each ws In Worksheets
    ' Set a variable for ticker type
    Dim TickerType As String
    TickerType = 0
    ' Set total for Yearly Change
    Dim YearlyChange As Double
    YearlyChange = 0
    'Set Variable for Percent Change
    Dim PercentChange As Double
    PercentChange = 0
    'Set Variable for Total Stock Volume
    Dim TtlStockVol As Double
    TtlStockVol = 0
    'Set Variable for Open Price
    Dim OpenPrice As Double
    'Set Variable for Close Price
    Dim ClosePrice As Double
    'Set Variable for OpenStock
    Dim OpenStock As Double
    OpenStock = 2
    
    'Set Headings
    ws.Range("I1").Value = "Ticker Type"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"


    'Keep track of Ticker Type in a summary table
    Dim SummaryTableRow As Integer
    SummaryTableRow = 2
    
    'Finding the last row of the page
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Loop through the rows of columns
    For i = 2 To LastRow
    
   'Search for the values of the next cell is different than the current cell
   If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
   
   'Set to the Ticker Type
   TickerType = ws.Cells(i, 1).Value
   'Print the Ticker Type in summary table
   ws.Range("I" & SummaryTableRow).Value = TickerType
   
   'Set to Open Price
   OpenPrice = ws.Cells(OpenStock, 3).Value
   
   'Set to Close Price
   ClosePrice = ws.Cells(i, 6).Value
   
   'Set to yearly Change
   YearlyChange = ClosePrice - OpenPrice
   'YearlyChange = YearlyChange + ((Cells(i, 6).Value / Cells(i, 3).Value) / 100) *WRONG
   'Print the Yearly Ceange in summary table
   ws.Range("J" & SummaryTableRow).Value = YearlyChange
   
   'Set to Percent Change
   PercentChange = YearlyChange / OpenPrice
   'Print the Percent Change in summary table
   ws.Range("K" & SummaryTableRow).Value = PercentChange
   
   'Make Percent Changes a percent
   ws.Range("K" & SummaryTableRow).NumberFormat = "0.00%"
   
   'Set to Total Stock Volume
   TtlStockVolume = TtlStockVolume + ws.Cells(i, 7).Value
   
    'Print the Total Stock Value in summary table
   ws.Range("L" & SummaryTableRow).Value = TtlStockVolume
   
    'Reset the Total Stock Volume
   TtlStockVolume = 0
   
   'Add one to the summary table row
   SummaryTableRow = SummaryTableRow + 1
   
   
   'Add one to the open price
   OpenPrice = i + 1
   
   'Assign Color
   
        'Green is > 0
        If ws.Range("J" & SummaryTableRow - 1).Value > 0 Then
        ws.Range("J" & SummaryTableRow - 1).Interior.ColorIndex = 4
        
        Else
        
        'Red if < 0
        ws.Range("J" & SummaryTableRow - 1).Interior.ColorIndex = 3
        End If
    
   Else
   
      'Set to Total Stock Volume (This keeps the running total)
      TtlStockVolume = TtlStockVolume + ws.Cells(i, 7).Value
   
   
    End If
    
Next i

Next ws

End Sub


    

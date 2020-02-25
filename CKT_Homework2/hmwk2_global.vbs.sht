Sub Stocks()
For Each ws In Worksheets
  Worksheets(ws.Name).Activate
' Set an initial variable for holding the ticker name
    Dim TickerSymbol As String
    TickerSymbol = " "
' Set an initial variable for holding the total ticker volume
    Dim TotalTickerVol As Double
    TotalTickerVol = 0
' Set an initial variable for holding open variable, closed variable, price change
    Dim OpenPrice As Double
    OpenPrice = 0
    Dim ClosePrice As Double
    ClosePrice = 0
    Dim PriceChange As Double
    ChangePrice = 0
    Dim PercentChange As Double
    PercentChange = 0
' Keep track of the location ticker name in the summary table
    Dim SummaryTableRow As Long
    SummaryTableRow = 2
' Set row count for worksheet
    Dim LastRow As Long
    Dim i As Long
    LastRow = Cells(Rows.Count, "A").End(xlUp).Row
'Summary Table Titles
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
 ' Set initial value of Open Price for the first Ticker of sht
    OpenPrice = Cells(2, 3).Value
 ' Loop through all stock entries
    For i = 2 To LastRow
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
 ' Set the ticker symbol to insert name
            TickerSymbol = Cells(i, "A").Value
 'Retrieve ClosePrice
            ClosePrice = Cells(i, "F").Value
' Calculate Price Change
            PriceChange = ClosePrice - OpenPrice
' Calculate % Change
            PercentChange = (PriceChange / OpenPrice) * 100
' *Calculate Total Volume
            TotalTickerVol = TotalTickerVol + Cells(i, "G").Value
' Print TickerSymbol in the Summary Table (I)
            Range("I" & SummaryTableRow).Value = TickerSymbol
' Print the PriceChange to the Summary Table (J)
            Range("J" & SummaryTableRow).Value = PriceChange
            If (PriceChange > 0) Then
            Range("J" & SummaryTableRow).Interior.ColorIndex = 4
            ElseIf (PriceChange <= 0) Then
            Range("J" & SummaryTableRow).Interior.ColorIndex = 3
            End If
' Print the PercentChange to the Summary Table (K)
            Range("K" & SummaryTableRow).Value = (PercentChange)
' Print the TotalTickerVol to the Summary Table (L)
            Range("L" & SummaryTableRow).Value = TotalTickerVol

' Add one to the summary table row
            SummaryTableRow = SummaryTableRow + 1

' Reset PriceChange and PercentChange (for new Ticker)
            PriceChange = 0
            PercentChange = 0
            ClosePrice = 0
            TotalTickerVol = 0
' Next Ticker's OpenPrice (for new Ticker)
            OpenPrice = Cells(i + 1, 3).Value
        Else
' *Calculate Total Volume
            TotalTickerVol = TotalTickerVol + Cells(i, "G").Value
        End If
  Next i
 Next ws
End Sub

  
  
  


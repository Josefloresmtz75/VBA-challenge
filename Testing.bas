Attribute VB_Name = "Module1"
Sub wallStreet()
    Dim Sheets As Worksheet
    Application.ScreenUpdating = False
    For Each Sheets In Worksheets
        Sheets.Select
        Call VBAStocks
    Next
    Application.ScreenUpdating = True
End Sub
Sub VBAStocks()
'Put Titles
Range("I1") = "Ticker"
Range("J1") = "Open"
Range("K1") = "Close"
Range("L1") = "YearlyChange"
Range("M1") = "PercentChange"
Range("N1") = "TotalStockVolume"
Range("P2") = "Greatest%Increase"
Range("P3") = "Greatest%Decrease"
Range("P2") = "GreatestTotalVolume"
'Declare Variables
Dim Column, Current As Integer
Dim Rowdelimiter, Count, Row, Rowdelimiter_2  As Long
Dim tickerClose, tickerOpen, Volume, Max_increase, Min_increase, Max_Volume As Double
Rowdelimiter = Cells(Rows.Count, 1).End(xlUp).Row
Current = 2
'Loop for values
For Row = 2 To Rowdelimiter
    If Cells(Row, 1).Value = Cells(Row + 1, 1).Value Then
    Count = Count + Cells(Row, 1).Rows.Count
    Volume = Volume + Cells(Row, 7).Value
    Else
    Cells(Current, 9) = Cells(Row, 1)
    Count = Cells(Row, 1).Rows.Count + Count
    Volume = Cells(Row, 7).Value + Volume
    tickerClose = Cells(Row, 6).Value
    tickerOpen = Cells(Row - Count + 1, 3).Value
    Cells(Current, 10) = tickerOpen
    Cells(Current, 11) = tickerClose
    Cells(Current, 12) = tickerClose - tickerOpen
       If Cells(Current, 12) < 0 Then Cells(Current, 12).Interior.ColorIndex = 3 Else Cells(Current, 12).Interior.ColorIndex = 4
       If tickerOpen = 0 Then Cells(Current, 13) = 0 Else Cells(Current, 13) = tickerClose / tickerOpen - 1
    Cells(Current, 14) = Volume
    Current = Current + 1
    Volume = 0
    Count = 0
End If
Next Row
'Summary Table
Rowdelimiter_2 = Cells(Rows.Count, 9).End(xlUp).Row
Max_increase = Application.WorksheetFunction.Max(Range("M3:M" & Rowdelimiter_2))
Min_increase = Application.WorksheetFunction.Min(Range("M3:M" & Rowdelimiter_2))
Max_Volume = Application.WorksheetFunction.Max(Range("N3:N" & Rowdelimiter_2))
maxIncreaseRow = Range("M:M").Find(what:=Max_increase, lookat:=xlPart).Row
minIncreaseRow = Range("M:M").Find(what:=Min_increase, lookat:=xlPart).Row
MaxVolumeRow = Range("N:N").Find(what:=Max_Volume, lookat:=xlPart).Row
Cells(2, 17) = Cells(maxIncreaseRow, 9).Value
Cells(2, 18) = Max_increase
Cells(3, 17) = Cells(minIncreaseRow, 9).Value
Cells(3, 18) = Min_increase
Cells(4, 17) = Cells(MaxVolumeRow, 9).Value
Cells(4, 18) = Max_Volume
End Sub



    


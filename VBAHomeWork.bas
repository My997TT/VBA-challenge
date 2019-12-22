Attribute VB_Name = "VBAHomeWork"
Sub StockSummary()

Dim Ticker As String

Dim StOpen As Double
Dim StClose As Double

Dim YrChange As Double
Dim PercentChange As Double
Dim StockVolume As LongLong

Dim tblrow As Integer

addHeaders

tblrow = 2

'Count Rows
CountofRows = ActiveSheet.Range("A2", ActiveSheet.Cells(Rows.Count, 1).End(xlUp)).Rows.Count + 1


For Row = 2 To CountofRows

    Ticker = Cells(Row, 1).Value

 If Cells(Row + 1, 1).Value <> Cells(Row, 1).Value Then
  
    StockVolume = Cells(Row, 7) + StockVolume
    Cells(tblrow, 12) = StockVolume
    StockVolume = 0
    StClose = Cells(Row, 6).Value
    
    
    YrChange = StClose - StOpen
    Cells(tblrow, 10) = YrChange
    
        If (StOpen = 0 Or YrChange = 0) Then
              PercentChange = 0
        Else:  PercentChange = YrChange / StOpen
             Cells(tblrow, 11) = PercentChange
        End If

    
    Cells(tblrow, 9) = Ticker
    tblrow = tblrow + 1

 Else: StockVolume = Cells(Row, 7).Value + StockVolume
         Ticker = Cells(Row, 1).Value
     
 End If

 
 If Cells(Row - 1, 1).Value <> Cells(Row, 1).Value Then
    StOpen = Cells(Row, 3).Value
    
 End If

Next


ChangeGreen
ChangeRed
FindMaxMinTicker

 Columns("I:Q").Select
 Columns("I:Q").EntireColumn.AutoFit

End Sub

Sub addHeaders()
   [I1:L1] = [{"Ticker","Yearly Change","Percent Change","Total Stock Volume"}]
   Range("K:K").NumberFormat = "0.00%"
   Range("L:L").NumberFormat = "#,##0"
   
   [P1:Q1] = [{"Ticker","Value"}]
   [O2] = [{"Greatest % Increase"}]
   [O3] = [{"Greatest % Decrease"}]
   [o4] = [{"Greatest Total Volume"}]
   Range("q2:q3").NumberFormat = "0.00%"
   Range("q4:q4").NumberFormat = "#,##0"
   
   
End Sub


Sub ChangeGreen()
    Range("J2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Cells.FormatConditions.Delete
    'Range("J2:J289").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 5287936
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
End Sub
Sub ChangeRed()
    Range("J2").Select
    Range(Selection, Selection.End(xlDown)).Select
    'Range("J2:J289").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
End Sub

Sub FindMaxMinTicker()

Dim Maxvalue As Double
Dim MinValue As Double
Dim MaxVol As LongLong
Dim TickerName As String
Dim Rcnt As Long

Maxvalue = Application.WorksheetFunction.max(Range("k:k"))
Cells(2, 17).Value = Maxvalue

MinValue = Application.WorksheetFunction.Min(Range("k:k"))
Cells(3, 17).Value = MinValue

MaxVol = Application.WorksheetFunction.max(Range("l:l"))
Cells(4, 17).Value = MaxVol

Rcnt = ActiveSheet.Range("i2", ActiveSheet.Cells(Rows.Count, 9).End(xlUp)).Rows.Count

For r = 2 To Rcnt

If Cells(r, 11).Value = Maxvalue Then
       Maxvalue = Cells(r, 11).Value
       TickerName = Cells(r, 9).Value
       Cells(2, 16).Value = TickerName
    End If

If Cells(r, 11).Value = MinValue Then
       MinValue = Cells(r, 11).Value
       TickerName = Cells(r, 9).Value
       Cells(3, 16).Value = TickerName
    End If

If Cells(r, 12).Value = MaxVol Then
       MaxVol = Cells(r, 12).Value
       TickerName = Cells(r, 9).Value
       Cells(4, 16).Value = TickerName
    End If


Next

End Sub



Sub getStockVolume()

Application.DisplayAlerts = False
Application.ScreenUpdating = False

tStart = Now()

Set a = Application
Set w = WorksheetFunction

VolCol = "G"
startPxCol = "C"
closePxCol = "F"
tickerCol = "A"

nSheets = a.Sheets.Count

For i = nSheets To 1 Step -1
    If InStr(Sheets(i).Name, "Result") > 0 Then
        Sheets(i).Delete
    End If
Next i

nSheets = a.Sheets.Count
Set ws = Sheets.Add(After:=Sheets(nSheets))
ws.Name = "Result-Intermediate"

ws.Range("A1") = "Year"
ws.Range("B1") = "Ticker"
ws.Range("C1") = "SheetNo"




nSheets = a.Sheets.Count
For i = 1 To nSheets - 1
    
    Sheets(i).Columns("J:K").ClearContents
    
    Sheets(i).Range("J1:J" & w.CountA(Sheets(i).Columns("A"))).Value2 = _
    Sheets(i).Range("B1:B" & w.CountA(Sheets(i).Columns("A"))).Value2
    
    Sheets(i).Columns("J").TextToColumns Destination:=Range("J1"), DataType:=xlFixedWidth, _
            FieldInfo:=Array(Array(0, 1), Array(4, 1)), TrailingMinusNumbers:=True
    
    Sheets(i).Range("J1") = "Year"
    Sheets(i).Range("K1") = "Month/Day"
    
    'copy the tickers to result sheet
    numOfOccupiedRowsinResultSheet = w.CountA(ws.Columns("A"))
    
    ws.Range("B" & (numOfOccupiedRowsinResultSheet + 1) & ":B" & (numOfOccupiedRowsinResultSheet + w.CountA(Sheets(i).Columns("A")) - 1)).Value2 = _
    Sheets(i).Range("A2:A" & w.CountA(Sheets(i).Columns("A"))).Value2
    
    ws.Range("A" & (numOfOccupiedRowsinResultSheet + 1) & ":A" & (numOfOccupiedRowsinResultSheet + w.CountA(Sheets(i).Columns("A")) - 1)).Value2 = _
    Sheets(i).Range("J2:J" & w.CountA(Sheets(i).Columns("A"))).Value2
    
    ws.Range("C" & (numOfOccupiedRowsinResultSheet + 1) & ":C" & (numOfOccupiedRowsinResultSheet + w.CountA(Sheets(i).Columns("A")) - 1)).Value2 = _
    i
    
    numOfOccupiedRowsinResultSheet = w.CountA(ws.Columns("A"))
    
    ws.Range("A1:C" & numOfOccupiedRowsinResultSheet).RemoveDuplicates Columns:=Array(1, 2, 3), _
        Header:=xlYes
Next i




        
numOfOccupiedRowsinResultSheet = w.CountA(ws.Columns("A"))

year_arr = ws.Range("A1:A" & numOfOccupiedRowsinResultSheet).Value2
ticker_arr = ws.Range("B1:B" & numOfOccupiedRowsinResultSheet).Value2
sheetno_arr = ws.Range("C1:C" & numOfOccupiedRowsinResultSheet).Value2


ReDim volume_arr(1 To numOfOccupiedRowsinResultSheet, 1 To 1)

volume_arr(1, 1) = "Total Stock Volume"
For r = 2 To numOfOccupiedRowsinResultSheet
    volume_arr(r, 1) = a.SumIfs(Sheets(sheetno_arr(r, 1)).Columns(VolCol), _
                                    Sheets(sheetno_arr(r, 1)).Columns("A"), ticker_arr(r, 1), _
                                    Sheets(sheetno_arr(r, 1)).Columns("J"), year_arr(r, 1))
                                    
                                    
Next r

'sort all the sheets to get the min dates first
For i = 1 To nSheets - 1
    totalRowCount = w.CountA(Sheets(i).Columns("A"))
    Sheets(i).Range("A1:K" & totalRowCount).Sort _
    Key1:=Sheets(i).Range("A1"), Order1:=xlAscending, _
    Key2:=Sheets(i).Range("K1"), Order2:=xlAscending, _
    Header:=xlYes
Next i


ReDim startPx_arr(1 To numOfOccupiedRowsinResultSheet, 1 To 1)
startPx_arr(1, 1) = "OpenPx"
For r = 2 To numOfOccupiedRowsinResultSheet
    startPx_arr(r, 1) = a.Index(Sheets(sheetno_arr(r, 1)).Columns(startPxCol), a.Match(ticker_arr(r, 1), Sheets(sheetno_arr(r, 1)).Columns("A"), 0))
Next r


For i = 1 To nSheets - 1
    totalRowCount = w.CountA(Sheets(i).Columns("A"))
    Sheets(i).Range("A1:K" & totalRowCount).Sort _
    Key1:=Sheets(i).Range("A1"), Order1:=xlAscending, _
    Key2:=Sheets(i).Range("K1"), Order2:=xlDescending, _
    Header:=xlYes
Next i


ReDim closePx_arr(1 To numOfOccupiedRowsinResultSheet, 1 To 1)
ReDim yrChangePx_arr(1 To numOfOccupiedRowsinResultSheet, 1 To 1)
ReDim percChangePx_arr(1 To numOfOccupiedRowsinResultSheet, 1 To 1)

closePx_arr(1, 1) = "ClosePx"
yrChangePx_arr(1, 1) = "Yearly Change"
percChangePx_arr(1, 1) = "Percentage Change"

For r = 2 To numOfOccupiedRowsinResultSheet
    closePx_arr(r, 1) = a.Index(Sheets(sheetno_arr(r, 1)).Columns(closePxCol), a.Match(ticker_arr(r, 1), Sheets(sheetno_arr(r, 1)).Columns("A"), 0))
    yrChangePx_arr(r, 1) = closePx_arr(r, 1) - startPx_arr(r, 1)
    
    If startPx_arr(r, 1) <> 0 Then
        percChangePx_arr(r, 1) = closePx_arr(r, 1) / startPx_arr(r, 1) - 1
    Else
        percChangePx_arr(r, 1) = 0
    End If
Next r


Erase year_arr
Erase ticker_arr
Erase sheetno_arr
Erase startPx_arr
Erase closePx_arr

ws.Range("D1:D" & numOfOccupiedRowsinResultSheet).Value2 = yrChangePx_arr
ws.Range("E1:E" & numOfOccupiedRowsinResultSheet).Value2 = percChangePx_arr
ws.Range("F1:F" & numOfOccupiedRowsinResultSheet).Value2 = volume_arr


Erase volume_arr
Erase percChangePx_arr
Erase yrChangePx_arr

'clear temp data from each data sheet
For i = 1 To nSheets - 1
    Sheets(i).Columns("J:K").ClearContents
Next i


ws.Columns("C").Delete


'sort the output to ensure years are in order
ws.Range("A1:E" & numOfOccupiedRowsinResultSheet).Sort _
    Key1:=Sheets(i).Range("A1"), Order1:=xlDescending, _
    Key2:=Sheets(i).Range("B1"), Order2:=xlAscending, _
    Header:=xlYes


'split the data as per the year
ws.Range("J1") = "Unique Years"
ws.Range("J2:J" & numOfOccupiedRowsinResultSheet).Value2 = ws.Range("A2:A" & numOfOccupiedRowsinResultSheet).Value2
ws.Range("J1:J" & numOfOccupiedRowsinResultSheet).RemoveDuplicates Columns:=Array(1), _
        Header:=xlYes
noOfYears = w.CountA(ws.Columns("J"))

nSheets = a.Sheets.Count
For i = 2 To noOfYears
    yearValue = ws.Range("J" & i)
    Set ws2 = Sheets.Add(After:=Sheets(nSheets))
    ws2.Name = "Result-" & yearValue
    ws2.Range("A1:E1").Value2 = ws.Range("A1:E1").Value2
    
    rowCountYearValue = w.CountIf(ws.Columns("A"), yearValue)
    firstRowYear = a.Match(yearValue, ws.Columns("A"), 0)
    
    ws2.Range("A2:E" & rowCountYearValue + 1).Value2 = ws.Range("A" & firstRowYear & ":E" & firstRowYear + rowCountYearValue - 1).Value2
    
    'formatting data
    ws2.Range("C2:C" & rowCountYearValue + 1).Select
    ws2.Range("C2:C" & rowCountYearValue + 1).FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 65280
        .TintAndShade = 0
    End With
    
    ws2.Range("C2:C" & rowCountYearValue + 1).FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
    End With

    ws2.Range("A1").Select
    
    ws2.Columns("C").NumberFormat = "0.00000000"
    ws2.Columns("D").NumberFormat = "0.00%"
    ws2.Columns("E").NumberFormat = "#,##0"
    
    ws2.Range("H1") = yearValue
    ws2.Range("H2") = "Greatest % Increase"
    ws2.Range("H3") = "Greatest % Decrease"
    ws2.Range("H4") = "Greatest Total Volume"
    ws2.Range("I1") = "Ticker"
    ws2.Range("J1") = "Value"
    
    'finding max/min
    ws2.Range("J2") = a.Max(ws2.Columns("D"))
    ws2.Range("J3") = a.Min(ws2.Columns("D"))
    ws2.Range("J4") = a.Max(ws2.Columns("E"))
    ws2.Range("I2") = a.Index(ws2.Columns("B"), a.Match(a.Max(ws2.Columns("D")), ws2.Columns("D"), 0))
    ws2.Range("I3") = a.Index(ws2.Columns("B"), a.Match(a.Min(ws2.Columns("D")), ws2.Columns("D"), 0))
    ws2.Range("I4") = a.Index(ws2.Columns("B"), a.Match(a.Max(ws2.Columns("E")), ws2.Columns("E"), 0))
    
    ws2.Range("J2:J3").NumberFormat = "0.00%"
    ws2.Range("J4").NumberFormat = "#,##0"
    
    ws2.Cells.EntireColumn.AutoFit
    
    nSheets = a.Sheets.Count
Next i

ThisWorkbook.Save

tend = Now()
MsgBox "Successfully Completed! Time Taken : " & (tend - tStart) * 24 * 60 & " Mins."

Application.DisplayAlerts = True
Application.ScreenUpdating = True


End Sub
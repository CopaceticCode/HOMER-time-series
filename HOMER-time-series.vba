' accounts for multiple arrays (different solar orientations)

Sub DeleteAllSheetsExcept() ' utility for removing all sheets except 'time series'
    Dim ws As Worksheet
    Application.DisplayAlerts = False ' Prevents confirmation dialogs
    
    ' First delete all sheets except "time series"
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "time series" Then ' Keep this sheet
            ws.Delete
        End If
    Next ws
    
    ' Delete "time series" summaries
    With ThisWorkbook.Worksheets("time series")
        ' Delete specified rows
        .Rows("8763:8795").Delete
        
        ' Unhide all rows
        .Rows.Hidden = False
    End With
    
    Application.DisplayAlerts = True
End Sub

Sub ReformatTimeSeriesDataExpanded()
    ' Define column numbers
    Dim colDateTime As Integer
    Dim colPVs() As Integer  ' Array to hold multiple PV column indices
    Dim numPVs As Integer    ' Number of PV arrays found
    Dim colLoad As Integer
    Dim colGrid As Integer
    Dim colBatteryDischarge As Integer
    Dim colBatteryCharge As Integer
    Dim colBatterySOC As Integer
    Dim colExcessProduction As Integer
    Dim colInverterPowerOutput As Integer
    
    ' Set column numbers based on headers
    colDateTime = getColumnIndex("Time")
    
    ' Find all PV columns
    numPVs = 0
    ReDim colPVs(0)  ' Initialize array
    
    ' Check for base PV column
    If getColumnIndex("Generic flat plate PV Power Output") > 0 Then
        numPVs = numPVs + 1
        ReDim Preserve colPVs(numPVs - 1)
        colPVs(numPVs - 1) = getColumnIndex("Generic flat plate PV Power Output")
    End If
    
    ' Check for numbered PV columns
    Dim i As Integer
    i = 1
    Do While getColumnIndex("Generic flat plate PV (" & i & ") Power Output") > 0
        numPVs = numPVs + 1
        ReDim Preserve colPVs(numPVs - 1)
        colPVs(numPVs - 1) = getColumnIndex("Generic flat plate PV (" & i & ") Power Output")
        i = i + 1
    Loop
    
    ' Only create dailyPVs array if we found PV columns
    If numPVs > 0 Then
        ReDim dailyPVs(numPVs - 1)
        ReDim wsPVs(numPVs - 1)
    Else
        MsgBox "No PV columns found in the data!", vbExclamation
        Exit Sub
    End If
    
    colLoad = getColumnIndex("Total Electrical Load Served")
    colGrid = getColumnIndex("Grid Purchases")
    colExcessProduction = getColumnIndex("Excess Electrical Production")
    colInverterPowerOutput = getColumnIndex("Inverter Power Output")
    colBatteryCharge = getColumnIndex("Generic 1kWh Li-Ion Charge Power")
    colBatteryDischarge = getColumnIndex("Generic 1kWh Li-Ion Discharge Power")
    colBatterySOC = getColumnIndex("Generic 1kWh Li-Ion State of Charge")
    
    ' Debug
    Debug.Print "colDateTime: " & colDateTime & vbCrLf & _
            "Number of PV Arrays: " & numPVs & vbCrLf & _
            "colLoad: " & colLoad & vbCrLf & _
            "colGrid: " & colGrid & vbCrLf & _
            "colBatteryDischarge: " & colBatteryDischarge & vbCrLf & _
            "colBatteryCharge: " & colBatteryCharge & vbCrLf & _
            "colBatterySOC: " & colBatterySOC & vbCrLf & _
            "colExcessProduction: " & colExcessProduction & vbCrLf & _
            "colInverterPowerOutput: " & colInverterPowerOutput

    Dim wsSource As Worksheet
    Dim wsLoad As Worksheet
    Dim wsGrid As Worksheet
    Dim wsBatterySOC As Worksheet
    Dim wsBatteryCharge As Worksheet
    Dim wsBatteryDischarge As Worksheet
    Dim wsSummary As Worksheet
    Dim wsHourlyAvg As Worksheet
    Dim wsInverter As Worksheet
    
    Dim lastRow As Long
    Dim currentDate As Date
    Dim previousDate As Date
    Dim currentHour As Integer
    Dim targetRow As Long
    Dim dailyLoad As Double
    Dim dailyBatterySupply As Double
    Dim dailyGrid As Double
    Dim dailyExcess As Double
    Dim dailyInverter As Double
    Dim hourlyTotals(0 To 23, 1 To 8) As Double ' 8 noted parameters
    Dim hourlyCounts(0 To 23) As Long
    
    ' Define date format
    Dim dateFormat As String
    dateFormat = "d mmm"
    
    ' Set the source worksheet (assuming it's the active sheet)
    Set wsSource = ActiveSheet
    
    ' Create PV worksheets first
    ReDim wsPVs(numPVs - 1)
    For i = 0 To numPVs - 1
        If Not SheetExists(IIf(i = 0 And getColumnIndex("Generic flat plate PV Power Output") > 0, _
                              "PV Hourly Output", _
                              "PV" & i & " Hourly Output")) Then
            Set wsPVs(i) = ThisWorkbook.Worksheets.Add
            If i = 0 And getColumnIndex("Generic flat plate PV Power Output") > 0 Then
                wsPVs(i).Name = "PV Hourly Output"
            Else
                wsPVs(i).Name = "PV" & i & " Hourly Output"
            End If
            ' Set up headers for each PV worksheet
            wsPVs(i).Cells(1, 1).Value = "Date"
            For j = 0 To 23
                wsPVs(i).Cells(1, j + 2).NumberFormat = "@"
                wsPVs(i).Cells(1, j + 2).Value = Format$(j, "00") & ":00"
            Next j
        Else
            Set wsPVs(i) = ThisWorkbook.Worksheets(IIf(i = 0 And getColumnIndex("Generic flat plate PV Power Output") > 0, _
                                                      "PV Hourly Output", _
                                                      "PV" & i & " Hourly Output"))
        End If
    Next i
    
    ' Create or get references to other worksheets (only once)
    If Not SheetExists("Inverter Output") Then
        Set wsInverter = ThisWorkbook.Worksheets.Add
        wsInverter.Name = "Inverter Output"
    Else
        Set wsInverter = ThisWorkbook.Worksheets("Inverter Output")
    End If
    
    If Not SheetExists("Load") Then
        Set wsLoad = ThisWorkbook.Worksheets.Add
        wsLoad.Name = "Load"
    Else
        Set wsLoad = ThisWorkbook.Worksheets("Load")
    End If
    
    If Not SheetExists("Grid Consumption") Then
        Set wsGrid = ThisWorkbook.Worksheets.Add
        wsGrid.Name = "Grid Consumption"
    Else
        Set wsGrid = ThisWorkbook.Worksheets("Grid Consumption")
    End If
    
    If Not SheetExists("Battery SOC") Then
        Set wsBatterySOC = ThisWorkbook.Worksheets.Add
        wsBatterySOC.Name = "Battery SOC"
    Else
        Set wsBatterySOC = ThisWorkbook.Worksheets("Battery SOC")
    End If
    
    If Not SheetExists("Battery Charge") Then
        Set wsBatteryCharge = ThisWorkbook.Worksheets.Add
        wsBatteryCharge.Name = "Battery Charge"
    Else
        Set wsBatteryCharge = ThisWorkbook.Worksheets("Battery Charge")
    End If
    
    If Not SheetExists("Battery Discharge") Then
        Set wsBatteryDischarge = ThisWorkbook.Worksheets.Add
        wsBatteryDischarge.Name = "Battery Discharge"
    Else
        Set wsBatteryDischarge = ThisWorkbook.Worksheets("Battery Discharge")
    End If
    
    If Not SheetExists("Daily Summary") Then
        Set wsSummary = ThisWorkbook.Worksheets.Add
        wsSummary.Name = "Daily Summary"
    Else
        Set wsSummary = ThisWorkbook.Worksheets("Daily Summary")
    End If
    
    If Not SheetExists("Hourly Averages") Then
        Set wsHourlyAvg = ThisWorkbook.Worksheets.Add
        wsHourlyAvg.Name = "Hourly Averages"
    Else
        Set wsHourlyAvg = ThisWorkbook.Worksheets("Hourly Averages")
    End If
    
    ' Set up headers for time-based sheets
    Dim sheetsToFormat As Variant
    sheetsToFormat = Array(wsInverter, wsLoad, wsGrid, wsBatterySOC, wsBatteryCharge, wsBatteryDischarge)
    
    For Each ws In sheetsToFormat
        ws.Cells(1, 1).Value = "Date"
        For i = 0 To 23
            With ws.Cells(1, i + 2)
                .NumberFormat = "@"
                .Value = Format$(i, "00") & ":00"
            End With
        Next i
    Next ws
    
    ' Set up headers for Summary sheet
    wsSummary.Cells(1, 1).Value = "Date"
    wsSummary.Cells(1, 2).Value = "Total Load (kWh)"
    wsSummary.Cells(1, 3).Value = "PV Production (kWh)"
    wsSummary.Cells(1, 4).Value = "Total Inverter Output (kWh)"
    wsSummary.Cells(1, 5).Value = "Battery Supply (kWh)"
    wsSummary.Cells(1, 6).Value = "Grid Purchases (kWh)"
    wsSummary.Cells(1, 7).Value = "Excess Production (kWh)"
    
    ' Set up headers for Hourly Averages
    wsHourlyAvg.Cells(1, 1).Value = "Hour"
    wsHourlyAvg.Cells(1, 2).Value = "Average Load"
    wsHourlyAvg.Cells(1, 3).Value = "Average PV Production"
    wsHourlyAvg.Cells(1, 4).Value = "Average Excess Production"
    wsHourlyAvg.Cells(1, 5).Value = "Average Inverter Output"
    wsHourlyAvg.Cells(1, 6).Value = "Average Grid Consumption"
    wsHourlyAvg.Cells(1, 7).Value = "Average Battery Discharge"
    wsHourlyAvg.Cells(1, 8).Value = "Average Battery Charging"
    wsHourlyAvg.Cells(1, 9).Value = "Average Battery SOC"
    
    For i = 0 To 23
        With wsHourlyAvg.Cells(i + 2, 1)
            .NumberFormat = "@"
            .Value = Format$(i, "00") & ":00"
        End With
    Next i
    
    ' Initialize arrays
    ReDim dailyPVs(numPVs - 1)
    
    ' Find the last row with data
    lastRow = wsSource.Cells(wsSource.Rows.Count, colDateTime).End(xlUp).Row
    
    ' Initialize target row
    targetRow = 2
    
    ' Loop through source data
    For i = 3 To lastRow
        currentDate = DateValue(wsSource.Cells(i, colDateTime).Value)
        currentHour = Hour(CDate(wsSource.Cells(i, colDateTime).Value))
        
        ' If it's a new day, start new rows and reset daily totals
        If i = 3 Or currentDate <> previousDate Then
            If i > 3 Then
                ' Write daily summary for previous day
                wsSummary.Cells(targetRow - 1, 2).Value = dailyLoad
                
                ' Calculate total PV production
                Dim totalDailyPV As Double
                totalDailyPV = 0
                For j = 0 To numPVs - 1
                    totalDailyPV = totalDailyPV + dailyPVs(j)
                Next j
                wsSummary.Cells(targetRow - 1, 3).Value = totalDailyPV
                
                wsSummary.Cells(targetRow - 1, 4).Value = dailyInverter
                wsSummary.Cells(targetRow - 1, 5).Value = dailyBatterySupply
                wsSummary.Cells(targetRow - 1, 6).Value = dailyGrid
                wsSummary.Cells(targetRow - 1, 7).Value = dailyExcess
            End If
            
            ' Write date to all sheets
            For j = 0 To numPVs - 1
                wsPVs(j).Cells(targetRow, 1).Value = Format(currentDate, "d mmm")
            Next j
            wsLoad.Cells(targetRow, 1).Value = Format(currentDate, "d mmm")
            wsGrid.Cells(targetRow, 1).Value = Format(currentDate, "d mmm")
            wsInverter.Cells(targetRow, 1).Value = Format(currentDate, "d mmm")
            If colBatterySOC > 0 Then wsBatterySOC.Cells(targetRow, 1).Value = Format(currentDate, "d mmm")
            If colBatteryCharge > 0 Then wsBatteryCharge.Cells(targetRow, 1).Value = Format(currentDate, "d mmm")
            If colBatteryDischarge > 0 Then wsBatteryDischarge.Cells(targetRow, 1).Value = Format(currentDate, "d mmm")
            wsSummary.Cells(targetRow, 1).Value = Format(currentDate, "d mmm")
            
            ' Reset daily totals
            dailyLoad = 0
            For j = 0 To numPVs - 1
                dailyPVs(j) = 0
            Next j
            dailyBatterySupply = 0
            dailyGrid = 0
            dailyExcess = 0
            dailyInverter = 0
            
            targetRow = targetRow + 1
        End If
        
        ' Write hourly data to sheets
        For j = 0 To numPVs - 1
            If colPVs(j) > 0 Then
                wsPVs(j).Cells(targetRow - 1, currentHour + 2).Value = wsSource.Cells(i, colPVs(j)).Value
                dailyPVs(j) = dailyPVs(j) + wsSource.Cells(i, colPVs(j)).Value
            End If
        Next j
        
        wsLoad.Cells(targetRow - 1, currentHour + 2).Value = wsSource.Cells(i, colLoad).Value
        wsGrid.Cells(targetRow - 1, currentHour + 2).Value = wsSource.Cells(i, colGrid).Value
        wsInverter.Cells(targetRow - 1, currentHour + 2).Value = wsSource.Cells(i, colInverterPowerOutput).Value
        
        If colBatterySOC > 0 Then wsBatterySOC.Cells(targetRow - 1, currentHour + 2).Value = wsSource.Cells(i, colBatterySOC).Value
        If colBatteryCharge > 0 Then wsBatteryCharge.Cells(targetRow - 1, currentHour + 2).Value = wsSource.Cells(i, colBatteryCharge).Value
        If colBatteryDischarge > 0 Then wsBatteryDischarge.Cells(targetRow - 1, currentHour + 2).Value = wsSource.Cells(i, colBatteryDischarge).Value
        
        ' Update daily totals
        dailyLoad = dailyLoad + wsSource.Cells(i, colLoad).Value
        dailyGrid = dailyGrid + wsSource.Cells(i, colGrid).Value
        If colBatteryDischarge > 0 Then
            dailyBatterySupply = dailyBatterySupply + wsSource.Cells(i, colBatteryDischarge).Value
        End If
        dailyExcess = dailyExcess + wsSource.Cells(i, colExcessProduction).Value
        dailyInverter = dailyInverter + wsSource.Cells(i, colInverterPowerOutput).Value
        
        ' Update hourly totals for averages
        hourlyTotals(currentHour, 1) = hourlyTotals(currentHour, 1) + wsSource.Cells(i, colLoad).Value
        
        ' Calculate total PV for hourly averages
        Dim totalHourlyPV As Double
        totalHourlyPV = 0
        For j = 0 To numPVs - 1
            If colPVs(j) > 0 Then
                totalHourlyPV = totalHourlyPV + wsSource.Cells(i, colPVs(j)).Value
            End If
        Next j
        hourlyTotals(currentHour, 2) = hourlyTotals(currentHour, 2) + totalHourlyPV
        
        hourlyTotals(currentHour, 3) = hourlyTotals(currentHour, 3) + wsSource.Cells(i, colExcessProduction).Value
        hourlyTotals(currentHour, 4) = hourlyTotals(currentHour, 4) + wsSource.Cells(i, colInverterPowerOutput).Value
        hourlyTotals(currentHour, 5) = hourlyTotals(currentHour, 5) + wsSource.Cells(i, colGrid).Value
        If colBatteryDischarge > 0 Then hourlyTotals(currentHour, 6) = hourlyTotals(currentHour, 6) + wsSource.Cells(i, colBatteryDischarge).Value
        If colBatteryCharge > 0 Then hourlyTotals(currentHour, 7) = hourlyTotals(currentHour, 7) + wsSource.Cells(i, colBatteryCharge).Value
        If colBatterySOC > 0 Then hourlyTotals(currentHour, 8) = hourlyTotals(currentHour, 8) + wsSource.Cells(i, colBatterySOC).Value
        
        hourlyCounts(currentHour) = hourlyCounts(currentHour) + 1
        previousDate = currentDate
    Next i
    
    ' Write final day's summary
    wsSummary.Cells(targetRow - 1, 2).Value = dailyLoad
    
    ' Calculate final day's total PV
    totalDailyPV = 0
    For j = 0 To numPVs - 1
        totalDailyPV = totalDailyPV + dailyPVs(j)
    Next j
    wsSummary.Cells(targetRow - 1, 3).Value = totalDailyPV
    
    wsSummary.Cells(targetRow - 1, 4).Value = dailyInverter
    wsSummary.Cells(targetRow - 1, 5).Value = dailyBatterySupply
    wsSummary.Cells(targetRow - 1, 6).Value = dailyGrid
    wsSummary.Cells(targetRow - 1, 7).Value = dailyExcess
    
    ' Calculate and write hourly averages
    For i = 0 To 23
        If hourlyCounts(i) > 0 Then
            For j = 1 To 8
                wsHourlyAvg.Cells(i + 2, j + 1).Value = hourlyTotals(i, j) / hourlyCounts(i)
            Next j
        End If
    Next i
    
    ' Format sheets
    Dim wsSheet As Worksheet
    For Each wsSheet In ThisWorkbook.Worksheets
        Select Case wsSheet.Name
            Case "PV Hourly Output", "Inverter Output", "Load", "Grid Consumption"
                Call FormatWorksheet(wsSheet, dateFormat)
            Case "Daily Summary", "Hourly Averages"
                Call FormatWorksheet(wsSheet, dateFormat)
            Case "Battery SOC"
                If colBatterySOC > 0 Then Call FormatWorksheet(wsSheet, dateFormat)
            Case "Battery Charge"
                If colBatteryCharge > 0 Then Call FormatWorksheet(wsSheet, dateFormat)
            Case "Battery Discharge"
                If colBatteryDischarge > 0 Then Call FormatWorksheet(wsSheet, dateFormat)
        End Select
    Next wsSheet
    
    ' Create monthly totals
    CreateMonthlyTotals
    
    'Draw charts
    Call CreateEnergyCharts
    
    ' Process the time series sheet to create summary and diagram
    Call ProcessTimeSeriesSheet
    
    MsgBox "Data has been reformatted into separate analysis sheets.", vbInformation
End Sub


Function getColumnIndex(header As String) As Integer
    Dim ws As Worksheet
    Dim headerRange As Range
    Dim result As Variant
    
    Set ws = ActiveSheet
    Set headerRange = ws.Rows(1) ' This should be a range object
    
    ' Use Match to find the header
    On Error Resume Next ' Ignore errors if the header is not found
    result = Application.Match(header, headerRange, 0)
    On Error GoTo 0 ' Resume normal error handling
    
    ' Check if the result is a number
    If IsNumeric(result) Then
        getColumnIndex = result
    Else
        getColumnIndex = 0 ' or handle it as you like
        ' MsgBox "Header '" & header & "' not found in the first row."
    End If
End Function

Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    
    SheetExists = False
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = sheetName Then
            SheetExists = True
            Exit Function
        End If
    Next ws
End Function

Private Sub FormatWorksheet(ByRef ws As Worksheet, ByVal dateFormat As String)
    With ws
        Select Case ws.Name
        Case "Daily Summary"
            .Columns("A").NumberFormat = dateFormat
        Case "Load", "Grid Consumption", "Battery SOC", _
             "Battery Charge", "Battery Discharge"
                .Columns("A").NumberFormat = dateFormat
        End Select
        .Range("B:ZZ").NumberFormat = "0.00"
        .Columns.AutoFit
        .Range("A1").EntireRow.Font.Bold = True
    End With
End Sub

Sub PushItem(ByRef arr As Variant, item As Variant)
    Dim tempArr() As Worksheet
    Dim i As Long
    
    If Not IsArray(arr) Then Exit Sub
    
    ' Create temp array and copy existing items
    ReDim tempArr(0 To UBound(arr) + 1)
    For i = 0 To UBound(arr)
        Set tempArr(i) = arr(i)
    Next i
    
    ' Add new item
    Set tempArr(UBound(tempArr)) = item
    
    ' Point original array to temp array
    arr = tempArr
End Sub

Sub CreateMonthlyTotals()
    Dim wsSummary As Worksheet
    Dim wsMonthly As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim currentMonth As Integer
    Dim monthlyPV(1 To 12) As Double
    Dim monthlyLoad(1 To 12) As Double
    Dim monthlyGrid(1 To 12) As Double
    Dim monthlyBatterySupply(1 To 12) As Double
    Dim monthlyExcess(1 To 12) As Double
    Dim monthlyInverter(1 To 12) As Double
    Dim daysInMonth(1 To 12) As Integer
    
    Set wsSummary = ThisWorkbook.Worksheets("Daily Summary")
    Set wsMonthly = ThisWorkbook.Worksheets.Add
    wsMonthly.Name = "Monthly Totals"
    
    lastRow = wsSummary.Cells(wsSummary.Rows.Count, 1).End(xlUp).Row
    
    daysInMonth(1) = 31: daysInMonth(2) = 28: daysInMonth(3) = 31: daysInMonth(4) = 30
    daysInMonth(5) = 31: daysInMonth(6) = 30: daysInMonth(7) = 31: daysInMonth(8) = 31
    daysInMonth(9) = 30: daysInMonth(10) = 31: daysInMonth(11) = 30: daysInMonth(12) = 31
    
    For i = 1 To 12
        monthlyPV(i) = 0
        monthlyLoad(i) = 0
        monthlyGrid(i) = 0
        monthlyBatterySupply(i) = 0
        monthlyExcess(i) = 0
        monthlyInverter(i) = 0
    Next i
    
    For i = 2 To lastRow
        currentMonth = Month(wsSummary.Cells(i, 1).Value)
        monthlyLoad(currentMonth) = monthlyLoad(currentMonth) + wsSummary.Cells(i, 2).Value
        monthlyPV(currentMonth) = monthlyPV(currentMonth) + wsSummary.Cells(i, 3).Value
        monthlyInverter(currentMonth) = monthlyInverter(currentMonth) + wsSummary.Cells(i, 4).Value
        If wsSummary.Cells(i, 5).Value > 0 Then
            monthlyBatterySupply(currentMonth) = monthlyBatterySupply(currentMonth) + wsSummary.Cells(i, 5).Value
        End If
        monthlyGrid(currentMonth) = monthlyGrid(currentMonth) + wsSummary.Cells(i, 6).Value
        monthlyExcess(currentMonth) = monthlyExcess(currentMonth) + wsSummary.Cells(i, 7).Value
    Next i
    
    ' Headers (reordered)
    wsMonthly.Cells(1, 1).Value = "Month"
    wsMonthly.Cells(1, 2).Value = "Total Load (kWh)"
    wsMonthly.Cells(1, 3).Value = "Total Inverter Output (kWh)"
    wsMonthly.Cells(1, 4).Value = "Total PV (kWh)"
    wsMonthly.Cells(1, 5).Value = "Total Battery Discharge (kWh)"
    wsMonthly.Cells(1, 6).Value = "Total Grid (kWh)"
    wsMonthly.Cells(1, 7).Value = "Total Excess (kWh)"
    
    ' Monthly data (reordered)
    For i = 1 To 12
        wsMonthly.Cells(i + 1, 1).Value = MonthName(i)
        wsMonthly.Cells(i + 1, 2).Value = monthlyLoad(i)
        wsMonthly.Cells(i + 1, 3).Value = monthlyInverter(i)
        wsMonthly.Cells(i + 1, 4).Value = monthlyPV(i)
        wsMonthly.Cells(i + 1, 5).Value = monthlyBatterySupply(i)
        wsMonthly.Cells(i + 1, 6).Value = monthlyGrid(i)
        wsMonthly.Cells(i + 1, 7).Value = monthlyExcess(i)
    Next i
    
    ' Add sums with spacing
    wsMonthly.Cells(14, 1).Value = "Annual Totals"
    For i = 2 To 7
        wsMonthly.Cells(14, i).Formula = "=SUM(" & wsMonthly.Cells(2, i).Address & ":" & wsMonthly.Cells(13, i).Address & ")"
    Next i
    
    ' Add empty row for spacing
    wsMonthly.Rows(15).Interior.Color = xlNone
    
    ' System Specifications and Financial Metrics (shifted down one row)
    wsMonthly.Cells(16, 1).Value = "Array Nominal Size (STC)"
    wsMonthly.Cells(16, 2).Value = "" ' Input field
    
    ' Add NPV Rate input cell
    wsMonthly.Cells(13, 9).Value = "NPV Discount Rate"
    wsMonthly.Cells(13, 10).Value = 0.05 ' Default 5%
    wsMonthly.Cells(13, 10).NumberFormat = "0.0%"
    
    ' Side table headers
    wsMonthly.Cells(1, 9).Value = "Year"
    wsMonthly.Cells(1, 10).Value = "Annual Output (kWh)"
    wsMonthly.Cells(1, 11).Value = "Cash Flow"
    wsMonthly.Cells(1, 12).Value = "Net Present Value"
    wsMonthly.Cells(1, 13).Value = "Discount Factor"
    
    ' Calculate 10-year projections with discount factors
    Dim baseOutput As String
    baseOutput = wsMonthly.Cells(20, 2).Address ' Adjusted for new row spacing
    
    For i = 1 To 10
        wsMonthly.Cells(i + 1, 9).Value = i
        If i = 1 Then
            wsMonthly.Cells(i + 1, 10).Formula = "=" & baseOutput
        ElseIf i = 2 Then
            wsMonthly.Cells(i + 1, 10).Formula = "=" & baseOutput & "*0.98" ' 2% degradation
        Else
            wsMonthly.Cells(i + 1, 10).Formula = "=" & wsMonthly.Cells(i, 10).Address & "*0.995" ' 0.5% degradation
        End If
        
        ' Discount Factor formula
        wsMonthly.Cells(i + 1, 13).Formula = "=1/((1+$J$13)^" & i & ")"
        wsMonthly.Cells(i + 1, 13).NumberFormat = "0.000"
    Next i

    wsMonthly.Cells(25, 1).Value = "Grid Electricity Cost ($/kWh)"
    wsMonthly.Cells(25, 2).Value = "" ' Input field
    
    ' Calculate cash flows
    For i = 1 To 10
        wsMonthly.Cells(i + 1, 11).Formula = "=" & wsMonthly.Cells(i + 1, 10).Address & "*$B$26"
        wsMonthly.Cells(i + 1, 12).Formula = "=" & wsMonthly.Cells(i + 1, 11).Address & "/((1+$J$13)^" & i & ")"
    Next i
    
    ' System Specifications and Financial Metrics
    wsMonthly.Cells(16, 1).Value = "Array Nominal Size (STC)"
    wsMonthly.Cells(16, 2).Value = "" ' Input field
    
    wsMonthly.Cells(17, 1).Value = "Renewable Energy Coverage"
    wsMonthly.Cells(17, 2).Formula = "=" & wsMonthly.Cells(14, 3).Address & "/" & wsMonthly.Cells(14, 2).Address
    
    wsMonthly.Cells(18, 1).Value = "Estimated Annual Energy Production"
    wsMonthly.Cells(18, 2).Formula = "=" & wsMonthly.Cells(14, 4).Address
    
    wsMonthly.Cells(19, 1).Value = "Estimated Annual Load"
    wsMonthly.Cells(19, 2).Formula = "=" & wsMonthly.Cells(14, 2).Address
    
    wsMonthly.Cells(20, 1).Value = "Estimated Annual Inverter Output"
    wsMonthly.Cells(20, 2).Formula = "=" & wsMonthly.Cells(14, 3).Address
    
    wsMonthly.Cells(21, 1).Value = "Estimated Annual Battery Discharge"
    wsMonthly.Cells(21, 2).Formula = "=" & wsMonthly.Cells(14, 5).Address
    
    wsMonthly.Cells(22, 1).Value = "Estimated Annual Peak Shaving"
    wsMonthly.Cells(22, 2).Formula = "=" & wsMonthly.Cells(14, 7).Address
    
    wsMonthly.Cells(24, 1).Value = "System Cost"
    wsMonthly.Cells(24, 2).Value = "" ' Input field
    
    wsMonthly.Cells(25, 1).Value = "Cost per Installed Watt"
    wsMonthly.Cells(25, 2).Formula = "=IF(" & wsMonthly.Cells(16, 2).Address & "<>0," & wsMonthly.Cells(24, 2).Address & "/" & wsMonthly.Cells(16, 2).Address & "/1000,"""")"
    
    wsMonthly.Cells(26, 1).Value = "Grid Electricity Cost ($/kWh)"
    wsMonthly.Cells(26, 2).Value = "" ' Input field
    
    wsMonthly.Cells(27, 1).Value = "Simple Payback Period (Years)"
    wsMonthly.Cells(27, 2).Formula = "=IF(" & wsMonthly.Cells(24, 2).Address & "<>0,(" & wsMonthly.Cells(24, 2).Address & "/AVERAGE(K2:K11)),"""")"
    
    wsMonthly.Cells(28, 1).Value = "Return on Investment"
    wsMonthly.Cells(28, 2).Formula = "=IF(" & wsMonthly.Cells(24, 2).Address & "<>0,(SUM(L2:L11)/" & wsMonthly.Cells(24, 2).Address & ")-1,"""")"
        
        
        
        
    wsMonthly.Cells(29, 1).Value = "Internal Rate of Return"

    Dim ws As Worksheet
    Dim sysCell As Range
    Dim cashFlowRange As Range
    Dim runningTotal As Double

    Set ws = ActiveSheet
    Set sysCell = ws.Range("B24")        ' System cost cell
    Set cashFlowRange = ws.Range("K2:K11") ' Cash flow range

    ' Enter headers
    ws.Range("I15").Value = "Cash Flows"
    ws.Range("L15").Value = "Net Cash Position"

    ' Set formula for negative system cost in J15
    ws.Range("J15").Formula = "=-B24"

    ' Set formulas for cash flows from K2:K11
    For i = 1 To 10
        ws.Cells(15 + i, "J").Formula = "=K" & (i + 1)
    Next i

    ' Enter IRR formula in B29
    ws.Range("B29").Formula = "=IRR(J15:J25)"

    ' Set formulas for Net Cash Position
    ' First year (M15) = -System Cost + First Year Cash Flow
    ws.Range("M15").Formula = "=-B24+K2"

    ' Subsequent years: Previous balance + current year's cash flow
    For i = 1 To 9
        ws.Cells(15 + i, "M").Formula = "=M" & (14 + i) & "+K" & (i + 2)
    Next i
    
    
    
    
    wsMonthly.Cells(30, 1).Value = "Net Present Value"
    wsMonthly.Cells(30, 2).Formula = "=IF(" & wsMonthly.Cells(24, 2).Address & "<>0,SUM(L2:L11)-" & wsMonthly.Cells(24, 2).Address & ","""")"
    
    ' Formatting
    With wsMonthly
        ' Format all numeric cells to 2 decimal places
        .Range("B2:G14").NumberFormat = "#,##0.00"
        .Range("B16:B30").NumberFormat = "#,##0.00"
        .Range("J2:L11").NumberFormat = "#,##0.00"
        
        ' Format percentages
        .Range("B17").NumberFormat = "0.00%"
        .Range("B28,B29").NumberFormat = "0.00%"
        .Range("J13").NumberFormat = "0.00%"
        .Range("M2:M11").NumberFormat = "0.000" ' Keep 3 decimals for discount factors
        
        ' Format currency
        .Range("B24,B25,B26,B30").NumberFormat = "$#,##0.00"
        
        ' Bold the Annual Totals row
        .Range("A14:G14").Font.Bold = True
        
        ' Format input cells with light gray background and thin border
        Dim inputCells As Range
        Set inputCells = .Range("B16,B24,B26,J13")
        With inputCells
            .Interior.Color = RGB(245, 245, 245)
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Borders.Color = RGB(200, 200, 200)
        End With
        
        ' Add input field note with subtle formatting
        .Cells(31, 1).Value = "Note: Fields with gray background require user input"
        With .Cells(31, 1)
            .Font.Italic = True
            .Font.Size = 8
            .Font.Color = RGB(100, 100, 100)
        End With
        
        ' Format headers and totals
        .Range("A1:G1,A14:G14,I1:M1").Font.Bold = True
        
        ' AutoFit columns
        .Columns("A:M").AutoFit
        
        ' Add alternating row colors in the main table
        With .Range("A2:G13")
            .Interior.Color = RGB(252, 252, 252)
        End With
    End With
    
    ' Modified Net Cash Position section
    wsMonthly.Cells(14, 12).Value = "Net Cash Position"
    For i = 1 To 10
        wsMonthly.Cells(14 + i, 12).Value = "Year " & i
    Next i
    
    ' Add Net Cash Position Chart
    Dim cashChart As Chart
    Set cashChart = wsMonthly.Shapes.AddChart2(201, xlColumnClustered).Chart
    
    With cashChart
        .SetSourceData wsMonthly.Range("L14:M24")
        .HasTitle = True
        .ChartTitle.Text = "Net Cash Position"
        
        With .Axes(xlValue)
            .HasTitle = False
            .TickLabels.NumberFormat = "$#,##0"
        End With
        
        With .Axes(xlCategory)
            .TickLabelPosition = xlTickLabelPositionLow
        End With
        
        .HasLegend = False
        .Parent.Top = wsMonthly.Range("D16").Top
        .Parent.Left = wsMonthly.Range("D16").Left
        .Parent.Width = 400
        .Parent.Height = 250
    End With
    
    ' Format B16 for kWp
    wsMonthly.Range("B16").NumberFormat = "#,##0.00 ""kWp"""
    
    ' Add Levelized Cost of Energy
    wsMonthly.Cells(23, 1).Value = "Levelized Cost of Energy"
    wsMonthly.Cells(23, 2).Formula = "=IF(SUM(J2:J11)>0,B24/SUM(J2:J11),0)"
    wsMonthly.Cells(23, 2).NumberFormat = "$#,##0.00"
    
    ' Set default grid electricity cost
    wsMonthly.Cells(26, 2).Value = 0.17
End Sub

Sub ProcessTimeSeriesSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("time series")
    
    ' Add sums row
    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Add sums in row 8763, each column summing its own data
    For i = 2 To lastCol
        Dim colLetter As String
        colLetter = Split(ws.Cells(1, i).Address, "$")(1)
        ws.Cells(8763, i).Formula = "=ROUND(SUM(" & colLetter & "3:" & colLetter & "8762),2)"
        ws.Cells(8763, i).NumberFormat = "#,##0.00"
    Next i
    
    ' Hide data rows
    ws.Rows("3:8762").Hidden = True
    
    ' Copy and transpose headers and values
    ws.Range(ws.Cells(1, 2), ws.Cells(1, lastCol)).Copy
    ws.Range("A8765").PasteSpecial Transpose:=True
    
    ' Copy values only for sums (not formulas)
    ws.Range(ws.Cells(8763, 2), ws.Cells(8763, lastCol)).Copy
    ws.Range("B8765").PasteSpecial Paste:=xlPasteValues, Transpose:=True
    
    ' Create new sheet for Mermaid diagram
    Dim mermaidSheet As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("Energy Flow Diagram").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set mermaidSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets("Charts"))
    mermaidSheet.Name = "Energy Flow Diagram"
    
    ' Add Mermaid instructions and diagram
    mermaidSheet.Range("B2").Value = "To view energy flow diagram: Copy the code below and paste into a Mermaid editor (e.g. https://mermaid.live)"
    mermaidSheet.Range("B3").Value = CreateMermaidDiagram(ws)
    
    ' Format diagram cell for better viewing
    With mermaidSheet.Range("B3")
        .WrapText = True
        .RowHeight = 400
        .ColumnWidth = 100
    End With
    
    ' Format instruction cell
    With mermaidSheet.Range("B2")
        .Font.Bold = True
        .Font.Size = 11
    End With
End Sub

Private Function FindColumnByHeader(ws As Worksheet, headerName As String) As Long
    Dim headerRange As Range
    Set headerRange = ws.Rows(1).Find(headerName, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not headerRange Is Nothing Then
        FindColumnByHeader = headerRange.Column
    Else
        FindColumnByHeader = 0
    End If
End Function

Function CreateMermaidDiagram(ws As Worksheet) As String
    ' Find relevant columns
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Get column numbers for each metric
    Dim colPVTotal As Long: colPVTotal = FindColumnByHeader(ws, "Total Renewable Power Output")
    Dim colExcess As Long: colExcess = FindColumnByHeader(ws, "Excess Electrical Production")
    Dim colInverterInput As Long: colInverterInput = FindColumnByHeader(ws, "Inverter Power Input")
    Dim colInverterOutput As Long: colInverterOutput = FindColumnByHeader(ws, "Inverter Power Output")
    
    ' Battery columns are optional
    Dim hasBattery As Boolean
    Dim colBatteryIn As Long: colBatteryIn = FindColumnByHeader(ws, "Generic 1kWh Li-Ion Charge Power")
    Dim colBatteryOut As Long: colBatteryOut = FindColumnByHeader(ws, "Generic 1kWh Li-Ion Discharge Power")
    hasBattery = (colBatteryIn > 0 And colBatteryOut > 0)
    
    ' Calculate totals using found columns
    Dim pvTotal, inverterInput, excess, inverterOutput As Double
    Dim batteryIn, batteryOut, batteryLosses As Double
    Dim load, gridPurchases As Double
    
    pvTotal = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(2, colPVTotal), ws.Cells(lastRow, colPVTotal)))
    inverterInput = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(2, colInverterInput), ws.Cells(lastRow, colInverterInput)))
    excess = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(2, colExcess), ws.Cells(lastRow, colExcess)))
    inverterOutput = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(2, colInverterOutput), ws.Cells(lastRow, colInverterOutput)))
    
    ' Battery calculations only if present
    If hasBattery Then
        batteryIn = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(2, colBatteryIn), ws.Cells(lastRow, colBatteryIn)))
        batteryOut = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(2, colBatteryOut), ws.Cells(lastRow, colBatteryOut)))
        batteryLosses = batteryIn - batteryOut
    Else
        batteryIn = 0
        batteryOut = 0
        batteryLosses = 0
    End If
    
    ' Calculate derived values
    Dim inverterLosses As Double
    inverterLosses = inverterInput - inverterOutput
    batteryLosses = batteryIn - batteryOut
    load = inverterOutput + gridPurchases
    
    ' Build diagram in parts
    Dim diagram As String
    Dim part1 As String, part2 As String, part3 As String
    
    ' Part 1
    part1 = "graph LR" & vbNewLine & _
            "    A[PV Output" & vbNewLine & _
            "    " & Format(pvTotal, "#,##0.00") & " kWh" & vbNewLine & _
            "    100%] --> B[Inverter Input" & vbNewLine & _
            "    " & Format(inverterInput, "#,##0.00") & " kWh" & vbNewLine & _
            "    " & Format(inverterInput / pvTotal * 100, "#0.0") & "%]" & vbNewLine & _
            "    A --> C[Excess" & vbNewLine & _
            "    " & Format(excess, "#,##0.00") & " kWh" & vbNewLine
            
    ' Part 2
    part2 = "    " & Format(excess / pvTotal * 100, "#0.0") & "%]" & vbNewLine & _
            "    B --> D[Inverter Output" & vbNewLine & _
            "    " & Format(inverterOutput, "#,##0.00") & " kWh" & vbNewLine & _
            "    " & Format(inverterOutput / pvTotal * 100, "#0.0") & "%]" & vbNewLine & _
            "    B --> IL[Inverter Losses" & vbNewLine & _
            "    " & Format(inverterLosses, "#,##0.00") & " kWh" & vbNewLine & _
            "    " & Format(inverterLosses / pvTotal * 100, "#0.0") & "%]"
            
    If hasBattery Then
        part2 = part2 & vbNewLine & _
                "    A --> BAT[Battery" & vbNewLine
    End If
    
    ' Part 3 - conditionally include battery nodes
    If hasBattery Then
        part3 = "    " & Format(batteryIn, "#,##0.00") & " kWh in" & vbNewLine & _
                "    " & Format(batteryOut, "#,##0.00") & " kWh out]" & vbNewLine & _
                "    BAT --> BL[Battery Losses" & vbNewLine & _
                "    " & Format(batteryLosses, "#,##0.00") & " kWh" & vbNewLine & _
                "    " & Format(batteryLosses / pvTotal * 100, "#0.0") & "%]" & vbNewLine & _
                "    BAT --> B" & vbNewLine
    Else
        part3 = vbNewLine
    End If
    
    part3 = part3 & _
            "    D --> E[Load" & vbNewLine & _
            "    " & Format(load, "#,##0.00") & " kWh" & vbNewLine & _
            "    " & Format(load / pvTotal * 100, "#0.0") & "%]" & vbNewLine & _
            "    F[Grid Purchases" & vbNewLine & _
            "    " & Format(gridPurchases, "#,##0.00") & " kWh" & vbNewLine & _
            "    " & Format(gridPurchases / pvTotal * 100, "#0.0") & "%] --> E"
    
    ' Combine all parts
    diagram = part1 & part2 & part3
    
    CreateMermaidDiagram = diagram
End Function

Sub CreateEnergyCharts()
    Dim ws As Worksheet
    Dim chtSheet As Worksheet
    Dim monthlyChart As Chart
    Dim hourlyChart As Chart
    Dim totalLoad As Double
    Dim totalInverter As Double
    Dim rng As Range
    
    ' Set reference to Monthly Totals sheet
    Set ws = ThisWorkbook.Sheets("Monthly Totals")
    
    ' Calculate totals to determine if PV should be included
    totalLoad = Application.Sum(ws.Range("B2:B13"))
    totalInverter = Application.Sum(ws.Range("C2:C13"))
    
    ' Add new worksheet for charts
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("Charts").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set chtSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets("Hourly Averages"))
    chtSheet.Name = "Charts"
    
    ' Create Monthly Chart
    Set monthlyChart = chtSheet.Shapes.AddChart2(201, xlLine).Chart
    
    ' Set data range based on condition
    If totalInverter < (0.9 * totalLoad) Then
        Set rng = ws.Range("A1:D13")  ' Include PV column
    Else
        Set rng = ws.Range("A1:C13")  ' Exclude PV column
    End If
    
    With monthlyChart
        .SetSourceData rng
        .HasTitle = True
        .ChartTitle.Text = "Energy Production and Consumption"
        
        ' Format axis
        With .Axes(xlValue)
            .HasTitle = True
            .AxisTitle.Text = "Energy (kWh)"
            .TickLabels.NumberFormat = "0"
        End With
        
        With .Axes(xlCategory)
            .TickLabelPosition = xlTickLabelPositionLow
            .TickLabelSpacing = 1
        End With
        
        ' Show Legend
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        
        ' Size and position (slightly taller to accommodate legend)
        .Parent.Top = 20
        .Parent.Left = 10
        .Parent.Width = 420
        .Parent.Height = 250  ' Increased from 210
    End With
    
    ' Create Hourly Chart
    Set hourlyChart = chtSheet.Shapes.AddChart2(201, xlLine).Chart
    
    With hourlyChart
        .SetSourceData Sheets("Hourly Averages").Range("A1:C25")
        .HasTitle = True
        .ChartTitle.Text = "Load and Solar Production"
        
        ' Format axis
        With .Axes(xlValue)
            .HasTitle = True
            .AxisTitle.Text = "Energy (kWh)"
            .TickLabels.NumberFormat = "0"
        End With
        
        ' Show Legend
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        
        ' Size and position (slightly taller to accommodate legend)
        .Parent.Top = 290  ' Adjusted for taller first chart
        .Parent.Left = 10
        .Parent.Width = 420
        .Parent.Height = 250 
    End With
End Sub

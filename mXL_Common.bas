Attribute VB_Name = "mXL_Common"
'*************************************************************************************************************************************************************************************************************************************************
'            COPYRIGHT NOTICE
'
' Copyright (C) David Briant 2009-2011 - All rights reserved
'
'*************************************************************************************************************************************************************************************************************************************************
 
Option Explicit

' error reporting
Private Const MODULE_NAME As String = "mXL_Common"
Private Const MODULE_VERSION As String = "0.0.0.1"

Private myTouchIDSeed As Long
Private myRangesByTouchID As New Dictionary
Private myValuesByTouchID As New Dictionary

' = get.workspace(10000)             - ?
' = get.workspace(44)
' = get.workspace(7124)

' ctrl F11 - brings up the macro editor


'*************************************************************************************************************************************************************************************************************************************************
' General utilities
'*************************************************************************************************************************************************************************************************************************************************

Function XLAddressOfCell(aCell As Range, Optional includeBookName As Boolean = False) As String
    If includeBookName Then
        XLAddressOfCell = "'[" & aCell.Worksheet.Parent.name() & "]" & aCell.Worksheet.name() & "'!" & aCell.Address
    Else
        XLAddressOfCell = "'" & aCell.Worksheet.name() & "'!" & aCell.Address
    End If
End Function

Function XLMakeVolatile(aVariant As Variant) As Variant
    Application.Volatile
    XLMakeVolatile = aVariant
End Function

Function XLTouch(aRange As Range) As Range
    Dim cellBeingTouched As Variant, cellValue As Variant
    For Each cellBeingTouched In aRange
        cellValue = cellBeingTouched.value
    Next
    Set XLTouch = aRange
End Function

Function XLIsFormula(aCell As Range) As Boolean
    XLIsFormula = Left$(aCell.Formula, 1) = "="
End Function

Function XLStripTicker(XLObjectTicker As String) As String
    Dim positionOfDot As Long
    positionOfDot = InStr(1, XLObjectTicker, ".")
    If positionOfDot > 0 Then
        XLStripTicker = Mid$(XLObjectTicker, 1, positionOfDot - 1)
    Else
        XLStripTicker = XLObjectTicker
    End If
End Function

Function XLSplit(text As String, delimiter As String) As Variant
    XLSplit = Split(text, delimiter)
End Function


'*********************************************************************************************************************************************************************
' XL Async Touch
'*********************************************************************************************************************************************************************

Function XLTw(Optional startTime As Variant, Optional aValueToReturn As Variant = Empty, Optional endTime As Variant, Optional cellToTouchTrue As Range = Nothing, _
    Optional cellToTouchFalse As Range = Nothing, Optional cellToTouchTime As Range, Optional cellToPasteResults As Range) As Variant
    ', Optional tickers As Range, Optional fields As String = ""
    XLTw = aValueToReturn
    If Not cellToTouchTrue Is Nothing Then XLQueueTouch cellToTouchTrue, True
    If Not cellToTouchFalse Is Nothing Then XLQueueTouch cellToTouchFalse, False
    If Not cellToTouchTime Is Nothing Then XLQueueTouch cellToTouchTime, endTime - startTime
    If Not cellToPasteResults Is Nothing Then XLQueueTouch cellToPasteResults, aValueToReturn
    'If Not tickers Is Nothing Then XLQueueTouch cellToPasteResults, fields, tickers
End Function

'*************************************************************************************************************************************************************************************************************************************************

Sub XLQueueTouch(cellToTouch As Range, aValue As Variant, Optional tickers As Range)
    Dim timerID As Long
    myTouchIDSeed = myTouchIDSeed + 1
    Set myRangesByTouchID(myTouchIDSeed) = cellToTouch
    myValuesByTouchID(myTouchIDSeed) = aValue
    'Set myTickersByTouchID(myTouchIDSeed) = tickers
    timerID = apiSetTimer(0, 0, 1, AddressOf XLTimerCallbackTouch)
End Sub

Sub XLTimerCallbackTouch(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal lngSysTime As Long)
    On Error Resume Next
    Application.OnTime Now() + TimeValue("00:00:01"), "XLTouchToSomething"
    If Err.Number = 0 Then apiKillTimer 0, idEvent
End Sub

Sub XLTouchToSomething()
    Dim keys As Variant, i As Long, rangeToTouch As Range, aValue As Variant, qt As QueryTable, tickers As Range, cell As Range, dataDestination As Range, fields As String
    Dim tickerCount As Long, query As String, width As Long
    keys = myRangesByTouchID.keys()
    For i = LBound(keys) To UBound(keys)
        Set rangeToTouch = myRangesByTouchID(keys(i))
        aValue = myValuesByTouchID(keys(i))
        rangeToTouch.value = aValue
    Next
    myRangesByTouchID.RemoveAll
    myValuesByTouchID.RemoveAll
    'myTickersByTouchID.RemoveAll
End Sub

' code to get data from a web query - in this case yahoo finance
'        If Not myTickersByTouchID(keys(i)) Is Nothing Then
'            rangeToTouch.ClearContents
'            On Error Resume Next
'            For Each qt In rangeToTouch.Parent.QueryTables
'                If IsError(qt.Destination) Then
'                    qt.Delete
'                End If
'            Next
'            On Error GoTo 0
'            For Each qt In rangeToTouch.Parent.QueryTables
'                If Not Application.Intersect(qt.Destination, rangeToTouch) Is Nothing Then
'                    qt.Delete
'                End If
'            Next
'            fields = myValuesByTouchID(keys(i))
'            Set tickers = myTickersByTouchID(keys(i))
'            For Each cell In tickers
'                tickerCount = tickerCount + 1
'                If tickerCount = 1 Then
'                    query = "http://download.finance.yahoo.com/d/quotes.csv?s=" & cell.value
'                Else
'                    query = query & "+" & cell.value
'                End If
'            Next
'            query = query & "&f=" & fields
'            Set dataDestination = Application.Range(rangeToTouch, rangeToTouch.Cells(tickerCount + 1))
'            dataDestination.ClearContents
'            width = dataDestination.ColumnWidth
'            Set qt = rangeToTouch.Parent.QueryTables.Add(Connection:="URL;" & query, Destination:=rangeToTouch)
'            qt.BackgroundQuery = True
'            qt.TablesOnlyFromHTML = False
'            qt.Refresh BackgroundQuery:=False
'            qt.SaveData = True
'            qt.Delete
'            Application.DisplayAlerts = False
'            dataDestination.TextToColumns Destination:=rangeToTouch, DataType:=xlDelimited, _
'                TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
'                Semicolon:=False, Comma:=True, Space:=False, other:=False
'            dataDestination.ColumnWidth = width
'            Application.DisplayAlerts = True
'        Else
'            rangeToTouch.value = aValue
'        End If


'*************************************************************************************************************************************************************************************************************************************************
' Module summary
'*************************************************************************************************************************************************************************************************************************************************

Private Function ModuleSummary() As Variant()
    ModuleSummary = Array(1, GLOBAL_PROJECT_NAME, MODULE_NAME, MODULE_VERSION)
End Function



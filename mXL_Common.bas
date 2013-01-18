Attribute VB_Name = "mXL_Common"
'************************************************************************************************************************************************
'
'    Copyright (c) 2009-2012 David Briant - see https://github.com/DangerMouseB
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Lesser General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Lesser General Public License for more details.
'
'    You should have received a copy of the GNU Lesser General Public License
'    along with this program.  If not, see <http://www.gnu.org/licenses/>.
'
'************************************************************************************************************************************************
 
Option Explicit

' error reporting
Private Const MODULE_NAME As String = "mXL_Common"
Private Const MODULE_VERSION As String = "0.0.0.1"

Private myPasteIDSeed As Long
Private myRangesByPasteID As New Dictionary
Private myValuesByPasteID As New Dictionary

Public myDoubleClickHandlerHolder As New Dictionary

' = get.workspace(10000)             - ?
' = get.workspace(44)
' = get.workspace(7124)

' ctrl F11 - brings up the macro editor


'*************************************************************************************************************************************************************************************************************************************************
' General utilities
'*************************************************************************************************************************************************************************************************************************************************

Function XLAddressOfCell(aCell As Range, Optional includeBookName As Boolean = False) As String
    If includeBookName Then
        XLAddressOfCell = "'[" & aCell.Worksheet.Parent.name() & "]" & aCell.Worksheet.name() & "'!" & aCell.address
    Else
        XLAddressOfCell = "'" & aCell.Worksheet.name() & "'!" & aCell.address
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

Function XLDropBlankRows(aRange As Variant, Optional columnToSelect As Long = 1, Optional errorsAsBlank As Boolean = False) As Variant
    Dim answer2D() As Variant, i As Long, i1 As Long, i2 As Long, j As Long, j1 As Long, j2 As Long, nextRow As Long
    DBCWhateverAs2DArray aRange, answer2D, i1, i2, j1, j2
    nextRow = i1
    For i = i1 To i2
        If IsError(answer2D(i, columnToSelect)) Then
            If Not errorsAsBlank Then XLDropBlankRows = CVErr(xlErrNA): Exit Function
        Else
            If answer2D(i, columnToSelect) <> "" Then
                For j = j1 To j2
                    answer2D(nextRow, j) = answer2D(i, j)
                Next
                nextRow = nextRow + 1
            End If
        End If
    Next
    If nextRow <= i2 Then
        XLDropBlankRows = DBSubArray(answer2D, 1, nextRow - 1, j1, j2)
    Else
        XLDropBlankRows = answer2D
    End If
End Function

Function XLSubArray(arrayFromXL As Variant, Optional ByVal r1 As Variant, Optional ByVal r2 As Variant, Optional ByVal c1 As Variant, Optional ByVal c2 As Variant) As Variant
    Dim arrayFromXL2D() As Variant
    DBCWhateverAs2DArray arrayFromXL, arrayFromXL2D
    XLSubArray = DBSubArray(arrayFromXL2D, r1, r2, c1, c2)
End Function


'*************************************************************************************************************************************************************************************************************************************************
' interpolation
'*************************************************************************************************************************************************************************************************************************************************

Function XLCellInterp(x1 As Range, x2 As Range) As Variant
    Dim delta As Double, answerD() As Double, numRows As Long, numCols As Long, i As Long
    If x1.Row <> x2.Row Then
        numRows = x2.Row - x1.Row + 1
        delta = (x2 - x1) / (numRows - 1)
        ReDim answer2D(1 To numRows - 2, 1 To 1)
        For i = 1 To numRows - 2
            answer2D(i, 1) = x1.value + i * delta
        Next
        XLCellInterp = answer2D
    ElseIf x1.Column <> x2.Column Then
        numCols = x2.Column - x1.Column + 1
        delta = (x2 - x1) / (numCols - 1)
        ReDim answer2D(1 To 1, 1 To numCols - 2)
        For i = 1 To numCols - 2
            answer2D(1, i) = x1.value + i * delta
        Next
        XLCellInterp = answer2D
    Else
        XLCellInterp = "Can only interpolate a horrizontal or vertical 1D array"
    End If
End Function


'*********************************************************************************************************************************************************************
' XL Paste Results
'*********************************************************************************************************************************************************************

Function XLPR(Optional startTime As Variant, Optional aValueToReturn As Variant = Empty, Optional endTime As Variant, Optional cellToPasteTrue As Range = Nothing, _
    Optional cellToPasteFalse As Range = Nothing, Optional cellToPasteTime As Range, Optional cellToPasteResults As Range) As Variant
    ', Optional tickers As Range, Optional fields As String = ""
    XLPR = aValueToReturn
    If Not cellToPasteTrue Is Nothing Then XLQueuePaste cellToPasteTrue, True
    If Not cellToPasteFalse Is Nothing Then XLQueuePaste cellToPasteFalse, False
    If Not cellToPasteTime Is Nothing Then XLQueuePaste cellToPasteTime, endTime - startTime
    If Not cellToPasteResults Is Nothing Then XLQueuePaste cellToPasteResults, aValueToReturn
    'If Not tickers Is Nothing Then XLQueuePaste cellToPasteResults, fields, tickers
End Function

'*************************************************************************************************************************************************************************************************************************************************

Sub XLQueuePaste(cellToPaste As Range, aValue As Variant, Optional tickers As Range)
    Dim timerID As Long
    myPasteIDSeed = myPasteIDSeed + 1
    Set myRangesByPasteID(myPasteIDSeed) = cellToPaste
    myValuesByPasteID(myPasteIDSeed) = aValue
    'Set myTickersByPasteID(myPasteIDSeed) = tickers
    timerID = apiSetTimer(0, 0, 1, AddressOf XLTimerCallbackPaste)
End Sub

Sub XLTimerCallbackPaste(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal lngSysTime As Long)
    Const DELAY As Date = 1# / 24# / 60# / 60# / 10#
    On Error Resume Next
    Application.OnTime Now() + DELAY, "XLApplicationCallbackPaste"
    If Err.Number = 0 Then apiKillTimer 0, idEvent
End Sub

Sub XLApplicationCallbackPaste()
    Dim keys As Variant, i As Long, rangeToPaste As Range, aValue As Variant, qt As QueryTable, tickers As Range, cell As Range, dataDestination As Range, fields As String
    Dim tickerCount As Long, query As String, width As Long
    keys = myRangesByPasteID.keys()
    For i = LBound(keys) To UBound(keys)
        Set rangeToPaste = myRangesByPasteID(keys(i))
        aValue = myValuesByPasteID(keys(i))
        rangeToPaste.value = aValue
    Next
    myRangesByPasteID.RemoveAll
    myValuesByPasteID.RemoveAll
    'myTickersByPasteID.RemoveAll
End Sub

' code to get data from a web query - in this case yahoo finance
'        If Not myTickersByPasteID(keys(i)) Is Nothing Then
'            rangeToPaste.ClearContents
'            On Error Resume Next
'            For Each qt In rangeToPaste.Parent.QueryTables
'                If IsError(qt.Destination) Then
'                    qt.Delete
'                End If
'            Next
'            On Error GoTo 0
'            For Each qt In rangeToPaste.Parent.QueryTables
'                If Not Application.Intersect(qt.Destination, rangeToPaste) Is Nothing Then
'                    qt.Delete
'                End If
'            Next
'            fields = myValuesByPasteID(keys(i))
'            Set tickers = myTickersByPasteID(keys(i))
'            For Each cell In tickers
'                tickerCount = tickerCount + 1
'                If tickerCount = 1 Then
'                    query = "http://download.finance.yahoo.com/d/quotes.csv?s=" & cell.value
'                Else
'                    query = query & "+" & cell.value
'                End If
'            Next
'            query = query & "&f=" & fields
'            Set dataDestination = Application.Range(rangeToPaste, rangeToPaste.Cells(tickerCount + 1))
'            dataDestination.ClearContents
'            width = dataDestination.ColumnWidth
'            Set qt = rangeToPaste.Parent.QueryTables.Add(Connection:="URL;" & query, Destination:=rangeToPaste)
'            qt.BackgroundQuery = True
'            qt.TablesOnlyFromHTML = False
'            qt.Refresh BackgroundQuery:=False
'            qt.SaveData = True
'            qt.Delete
'            Application.DisplayAlerts = False
'            dataDestination.TextToColumns Destination:=rangeToPaste, DataType:=xlDelimited, _
'                TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
'                Semicolon:=False, Comma:=True, Space:=False, other:=False
'            dataDestination.ColumnWidth = width
'            Application.DisplayAlerts = True
'        Else
'            rangeToPaste.value = aValue
'        End If


'*************************************************************************************************************************************************************************************************************************************************
' Register Double Click
'*************************************************************************************************************************************************************************************************************************************************

Function XLRegisterDoubleClick(monitoredRange As Range, action As String, ParamArray Arguments() As Variant) As String
    Dim handler As cXL_DCHandler, addressID As String, oldHandler As cXL_DCHandler
    Application.Volatile
    Set handler = New cXL_DCHandler
    handler.initialize monitoredRange, action, CVar(Arguments)
    addressID = XLAddressOfCell(monitoredRange)
    If myDoubleClickHandlerHolder.Exists(addressID) Then
        Set oldHandler = myDoubleClickHandlerHolder(addressID)
        oldHandler.disable
    End If
    Set myDoubleClickHandlerHolder(XLAddressOfCell(monitoredRange)) = handler
    XLRegisterDoubleClick = "[DC]"  'Registered """ & action & """ to " & XLAddressOfCell(monitoredRange)
End Function


'*************************************************************************************************************************************************************************************************************************************************
' Double Click Handler
'*************************************************************************************************************************************************************************************************************************************************

Sub XLDoubleClickHandler(target As Range, monitoredRange As Range, action As String, Arguments As Variant, Cancel As Boolean)
    On Error GoTo exceptionHandler
    Select Case UCase(action)
        Case "OPENFILENAME2CELL"
            XLOpenFileName2Cell target, action, Arguments
            Cancel = True
        Case "SAVEASFILENAME2CELL"
            XLSaveAsFileName2Cell target, action, Arguments
            Cancel = True
        Case "ADD"
            target.value = target.value + Arguments(0)
            Cancel = True
        Case "TODAY"
            target.value = Date
            Cancel = True
        Case "NOW"
            target.value = Now()
            Cancel = True
        Case "RANDOM"
            target.value = Rnd()
            Cancel = True
        Case "TRUE"
            target.value = True
            Cancel = True
        Case "FALSE"
            target.value = False
            Cancel = True
        Case "TOGGLE"
            target.value = Not (CBool(target.value))
            Cancel = True
    End Select
Exit Sub
exceptionHandler:
End Sub


'*************************************************************************************************************************************************************************************************************************************************
' Specific behaviours
'*************************************************************************************************************************************************************************************************************************************************

Private Sub XLOpenFileName2Cell(target As Range, action As String, Arguments As Variant)
    Dim filename As String, i As Long, i1 As Long, i2 As Long, pathRange As Range, fileNameRange As Range, fileFilter As Variant
    Set pathRange = Arguments(0)
    Set fileNameRange = Arguments(1)
    DBGetArrayBounds Arguments, 1, i1, i2
    fileFilter = "All Files,*.*"
    If i2 > 1 Then fileFilter = Arguments(2) & "," & fileFilter
    filename = Application.GetOpenFilename(fileFilter)
    If UCase(filename) = "FALSE" Then Exit Sub
    For i = Len(filename) To 1 Step -1
        If Mid$(filename, i, 1) = "\" Then
            pathRange.value = Mid$(filename, 1, i - 1)
            fileNameRange.value = Mid$(filename, i + 1)
            Exit Sub
        End If
    Next
    pathRange.value = ""
    fileNameRange.value = filename
End Sub

Private Sub XLSaveAsFileName2Cell(target As Range, action As String, Arguments As Variant)
    Dim filename As String, i As Long, i1 As Long, i2 As Long, pathRange As Range, fileNameRange As Range, fileFilter As Variant
    Set pathRange = Arguments(0)
    Set fileNameRange = Arguments(1)
    DBGetArrayBounds Arguments, 1, i1, i2
    fileFilter = "All Files,*.*"
    If i2 > 1 Then fileFilter = Arguments(2) & "," & fileFilter
    filename = Application.GetSaveAsFilename(fileNameRange.value, fileFilter)
    If UCase(filename) = "FALSE" Then Exit Sub
    For i = Len(filename) To 1 Step -1
        If Mid$(filename, i, 1) = "\" Then
            pathRange.value = Mid$(filename, 1, i - 1)
            fileNameRange.value = Mid$(filename, i + 1)
            Exit Sub
        End If
    Next
    pathRange.value = ""
    fileNameRange.value = filename
End Sub


'*************************************************************************************************************************************************************************************************************************************************
' Environment Variables
'*************************************************************************************************************************************************************************************************************************************************

Function XLVBGetEnv(name As String) As String
    XLVBGetEnv = VBGetEnv(name)
End Function

Function XLSetEnv(name As String, value As String) As String
    XLSetEnv = setEnv(name, value)
End Function

Function XLGetEnv(name As String, Optional maxSize As Long = 2048) As String
    XLGetEnv = getEnv(name, maxSize)
End Function

Function XLGetEnvs(names As Variant) As String()
    Dim answer2D() As String, names1D() As Variant, i As Long, i1 As Long, i2 As Long
    DBCWhateverAs1DArray names, names1D, i1, i2
    DBCreateNewArrayOfStrings answer2D, i1, i2, 1, 1
    For i = i1 To i2
        answer2D(i, 1) = getEnv(CStr(names1D(i)))
    Next
    XLGetEnvs = answer2D
End Function

Function XLSetEnvs(namesAndValues As Variant) As String
    Dim answer2D() As String, namesAndValues2D() As Variant, i As Long, i1 As Long, i2 As Long
    DBCWhateverAs2DArray namesAndValues, namesAndValues2D, i1, i2
    For i = i1 To i2
        setEnv CStr(namesAndValues2D(i, 1)), CStr(namesAndValues2D(i, 2))
    Next
    XLSetEnvs = "All set"
End Function


'*************************************************************************************************************************************************************************************************************************************************
' error reporting utilities
'*************************************************************************************************************************************************************************************************************************************************

Private Function ModuleSummary() As Variant()
    ModuleSummary = Array(1, GLOBAL_PROJECT_NAME, MODULE_NAME, MODULE_VERSION)
End Function



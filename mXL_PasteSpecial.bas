Attribute VB_Name = "mXL_PasteSpecial"
'************************************************************************************************************************************************
'
'    Copyright (c) 2009-2011 David Briant - see https://github.com/DangerMouseB
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

Private myLastCopiedAddress As String


'*************************************************************************************************************************************************************************************************************************************************
' paste special utilities
'*************************************************************************************************************************************************************************************************************************************************

Sub XLPasteSpecialValues()
    On Error GoTo exceptionHandler
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Exit Sub
exceptionHandler:
    XLPasteSpecialText
End Sub

Sub XLPasteSpecialFormulas()
    Selection.PasteSpecial Paste:=xlFormulas, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
End Sub

Sub XLPasteSpecialFormats()
    Selection.PasteSpecial Paste:=xlFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
End Sub

Sub XLPasteSpecialTranspose()
    Selection.PasteSpecial Paste:=xlAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
End Sub

Sub XLPasteSpecialText()
    On Error GoTo exceptionHandler
    ActiveSheet.PasteSpecial format:="Text", Link:=False, DisplayAsIcon:=False
Exit Sub
exceptionHandler:
    Stop
End Sub

Sub XLPasteSpecialUnicodeText()
    Selection.PasteSpecial format:="Unicode Text", Link:=False, DisplayAsIcon:=False
End Sub

Sub XLPasteSpecialHTML()
    Selection.PasteSpecial format:="HTML", Link:=False, DisplayAsIcon:=False
End Sub

Sub XLNoteAddressAndContinueCopying()
    If TypeName(Application.Selection) = "Range" Then myLastCopiedAddress = IIf(InStr(1, Application.Selection.Parent.name, " ") > 0, "'" & Application.Selection.Parent.name & "'", Application.Selection.Parent.name) & "!" & Application.Selection.Address
    Selection.Copy
End Sub

Sub XLPasteSpecialAddress()
    If myLastCopiedAddress = "" Then Exit Sub
    Selection.value = "'=" & myLastCopiedAddress
End Sub

Sub XLPasteSpecialGrids()
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone

    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeLeft).Weight = xlThin
    Selection.Borders(xlEdgeLeft).ColorIndex = xlAutomatic

    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).Weight = xlThin
    Selection.Borders(xlEdgeTop).ColorIndex = xlAutomatic

    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).Weight = xlThin
    Selection.Borders(xlEdgeBottom).ColorIndex = xlAutomatic

    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).Weight = xlThin
    Selection.Borders(xlEdgeRight).ColorIndex = xlAutomatic

    On Error Resume Next
    Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
'    Selection.Borders(xlInsideVertical).LineStyle = xlDash
'    Selection.Borders(xlInsideVertical).LineStyle = xlDashDot
'    Selection.Borders(xlInsideVertical).LineStyle = xlDashDotDot
'    Selection.Borders(xlInsideVertical).LineStyle = xlDot
'    Selection.Borders(xlInsideVertical).LineStyle = xlDouble
'    Selection.Borders(xlInsideVertical).LineStyle = xlSlantDashDot
'    Selection.Borders(xlInsideVertical).LineStyle = xlLineStyleNone
    Selection.Borders(xlInsideVertical).Weight = xlThin
    Selection.Borders(xlInsideVertical).ColorIndex = 15

    Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
    Selection.Borders(xlInsideHorizontal).Weight = xlThin
    Selection.Borders(xlInsideHorizontal).ColorIndex = 15
End Sub

Sub XLPasteSpecialGridBlankInside()
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone

    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeLeft).Weight = xlThin
    Selection.Borders(xlEdgeLeft).ColorIndex = xlAutomatic

    Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).Weight = xlThin
    Selection.Borders(xlEdgeTop).ColorIndex = xlAutomatic

    Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
    Selection.Borders(xlEdgeBottom).Weight = xlThin
    Selection.Borders(xlEdgeBottom).ColorIndex = xlAutomatic

    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).Weight = xlThin
    Selection.Borders(xlEdgeRight).ColorIndex = xlAutomatic

    On Error Resume Next
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
End Sub

Sub XLPasteSpecialNow()
    Dim oldFormat As Variant
    If Selection.rows.count <> 1 Or Selection.columns.count <> 1 Then Exit Sub
    oldFormat = Selection.NumberFormat
    Selection.value = Now()
    Selection.NumberFormat = oldFormat
End Sub


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cXL_DCHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Private WithEvents myWorksheet As Worksheet
Attribute myWorksheet.VB_VarHelpID = -1
Private myMonitoredRange As Range
Private myAction As String
Private myArguments As Variant

Sub initialize(monitoredRange As Range, action As String, Arguments As Variant)
    Set myWorksheet = monitoredRange.Parent
    Set myMonitoredRange = monitoredRange
    myAction = action
    myArguments = Arguments
End Sub

Private Sub myWorksheet_BeforeDoubleClick(ByVal target As Range, Cancel As Boolean)
    Dim address As String, handler As cXL_DCHandler
    If myWorksheet Is Nothing Then Exit Sub
    If target.Cells.Count <> 1 Then Exit Sub
    address = XLAddressOfCell(myMonitoredRange)
    If myDoubleClickHandlerHolder.Exists(address) Then
        Set handler = myDoubleClickHandlerHolder(address)
        If Not handler Is Me Then
            Me.disable
            Exit Sub
        End If
    End If
    If Not Application.Intersect(target, myMonitoredRange) Is Nothing Then XLDoubleClickHandler target, myMonitoredRange, myAction, myArguments, Cancel
End Sub

Sub disable()
    Set myWorksheet = Nothing
    Set myMonitoredRange = Nothing
    myAction = ""
    myArguments = Empty
End Sub

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
'    Copyright (c) 2009-2011, David Briant. All rights reserved.
'    Licensed under BSD 3-Clause License - see https://github.com/DangerMouseB
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

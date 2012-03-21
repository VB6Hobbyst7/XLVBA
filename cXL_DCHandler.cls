VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cXL_DCHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*************************************************************************************************************************************************************************************************************************************************
'
' Copyright (c) David Briant 2009-2011 - All rights reserved
'
'*************************************************************************************************************************************************************************************************************************************************
 
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
    If target.Cells.count <> 1 Then Exit Sub
    If Not Application.Intersect(target, myMonitoredRange) Is Nothing Then XLDoubleClickHandler target, myMonitoredRange, myAction, myArguments, Cancel
End Sub


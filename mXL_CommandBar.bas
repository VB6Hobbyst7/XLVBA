Attribute VB_Name = "mXL_CommandBar"
'*************************************************************************************************************************************************************************************************************************************************
'
' Copyright (c) David Briant 2008-2011 - All rights reserved
'
'*************************************************************************************************************************************************************************************************************************************************
 
Option Explicit


'*************************************************************************************************************************************************************************************************************************************************
' button bar utilities
'*************************************************************************************************************************************************************************************************************************************************

Function XLAddPopup(controls As CommandBarControls, ID As Integer, Optional startGroup As Boolean = False) As CommandBarPopup
    Set XLAddPopup = controls.Add(msoControlSplitButtonPopup, ID)
    XLAddPopup.BeginGroup = startGroup
End Function

Function XLAddComboBox(controls As CommandBarControls, ID As Integer, Optional startGroup As Boolean = False) As CommandBarComboBox
    Set XLAddComboBox = controls.Add(msoControlComboBox, ID)
    XLAddComboBox.BeginGroup = startGroup
End Function

Function XLAddSplitDropdown(controls As CommandBarControls, ID As Integer, Optional startGroup As Boolean = False) As CommandBarComboBox
    Set XLAddSplitDropdown = controls.Add(msoControlSplitDropdown, ID)
    XLAddSplitDropdown.BeginGroup = startGroup
End Function

Function XLAddButton(controls As CommandBarControls, ID As Integer, Optional startGroup As Boolean = False) As CommandBarButton
    Set XLAddButton = controls.Add(msoControlButton, ID)
    XLAddButton.BeginGroup = startGroup
End Function

Function XLGetCommandBar(name As String) As CommandBar
    On Error Resume Next
    Set XLGetCommandBar = Application.CommandBars(name)
End Function

Function XLGetCommandBarOrNew(name As String, Optional position As Variant, Optional temporary As Boolean = False)
    ' msoBarLeft, msoBarTop, msoBarRight, msoBarBottom    Indicate the left, top, right, and bottom coordinates of the new command bar
    ' msoBarFloating  Indicates that the new command bar won't be docked
    ' msoBarPopup Indicates that the new command bar will be a shortcut menu
    ' msoBarMenuBar   Indicates that the new command bar will replace the system menu bar on the Macintosh
    Dim bar As CommandBar, commandBarControl As Object, localPosition As Integer
    Set bar = XLGetCommandBar(name)
    If bar Is Nothing Then
        If IsMissing(position) Then
            localPosition = msoBarTop
        Else
            localPosition = position
        End If
        Set bar = Application.CommandBars.Add(name, localPosition, , temporary)
        bar.Visible = True
    Else
        If temporary Then
            If IsMissing(position) Then
                localPosition = bar.position
            Else
                localPosition = position
            End If
            bar.Delete
            Set bar = Application.CommandBars.Add(name, localPosition, , temporary)
            bar.Visible = True
        Else
            For Each commandBarControl In bar.controls
                If Not commandBarControl Is Nothing Then commandBarControl.Delete
            Next
        End If
    End If
    Set XLGetCommandBarOrNew = bar
End Function


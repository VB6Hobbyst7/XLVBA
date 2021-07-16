Attribute VB_Name = "mXL_Errors"
'************************************************************************************************************************************************
'
'    Copyright (c) 2009-2012, David Briant. All rights reserved.
'    Licensed under BSD 3-Clause License - see https://github.com/DangerMouseB
'
'************************************************************************************************************************************************

Option Explicit

' error reporting
Private Const MODULE_NAME As String = "mXL_Errors"
Private Const MODULE_VERSION As String = "1.0.0"   ' semver convention


Function xlMissing() As Variant
    xlMissing = CVErr(xlErrNA)
End Function

Function xlNull() As Variant
    xlNull = CVErr(xlErrNull)
End Function

Function xlNaN() As Variant
    xlNaN = CVErr(xlErrNum)
End Function

Function xlInf() As Variant
    xlInf = CVErr(xlErrDiv0)
End Function

Function xlErr() As Variant
    xlErr = CVErr(xlErrValue)
End Function


Function xlIsMissing(x As Variant) As Boolean
    xlIsMissing = (x = CVErr(xlErrNA))
End Function

Function xlIsNull(x As Variant) As Boolean
    xlIsNull = (x = CVErr(xlErrNull))
End Function

Function xlIsNaN(x As Variant) As Boolean
    xlIsNaN = (x = CVErr(xlErrNum))
End Function

Function xlIsInf(x As Variant) As Boolean
    xlIsInf = (x = CVErr(xlErrDiv0))
End Function

Function xlIsErr(x As Variant) As Boolean
    xlIsErr = (x = CVErr(xlErrValue))
End Function

Function xlIsVbErr(x As Variant) As Boolean
    xlIsVbErr = (VarType(x) = vbString) And (InStr(1, x, "#VB{") > 0)
End Function
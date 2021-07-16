Attribute VB_Name = "mXL_RandomNumbers"
'************************************************************************************************************************************************
'
'    Copyright (c) 2009-2011, David Briant. All rights reserved.
'    Licensed under BSD 3-Clause License - see https://github.com/DangerMouseB
'
'************************************************************************************************************************************************

Option Explicit

Private Declare Function uNRRan0 Lib "QUtils" (ByRef seed As Long) As Double
Private Declare Function uNRRan1 Lib "QUtils" (ByRef seed As Long) As Double
Private Declare Function uNRRan1_alternate Lib "QUtils" (seed As Long) As Double
Private Declare Function uNRGaussianRan1 Lib "QUtils" (seed As Long) As Double
Private Declare Function uNRGaussianRan1_Alternate Lib "QUtils" (seed As Long) As Double
Private Declare Function uNRLogNormalFromNormal Lib "QUtils" (zeroOneGaussian As Double, mean As Double, sd As Double) As Double

' error reporting
Private Const MODULE_NAME As String = "mXL_RandomNumbers"
Private Const MODULE_VERSION As String = "1.0.0"   ' semver convention


'*************************************************************************************************************************************************************************************************************************************************
' block generators
'*************************************************************************************************************************************************************************************************************************************************

Function XLRan1(Optional seed As Long = -1, Optional numRows As Long = 1, Optional numCols As Long = 1) As Double()
    Dim answer2D() As Double, i As Long, j As Long
    ReDim answer2D(1 To numRows, 1 To numCols)
    For i = 1 To numRows
        For j = 1 To numCols
            answer2D(i, j) = uNRRan1(seed)
        Next
    Next
    XLRan1 = answer2D
End Function

Function XLGaussianRan1(Optional seed As Long = -1, Optional numRows As Long = 1, Optional numCols As Long = 1) As Double()
    Dim answer2D() As Double, i As Long, j As Long
    ReDim answer2D(1 To numRows, 1 To numCols)
    For i = 1 To numRows
        For j = 1 To numCols
            answer2D(i, j) = uNRGaussianRan1(seed)
        Next
    Next
    XLGaussianRan1 = answer2D
End Function


'*************************************************************************************************************************************************************************************************************************************************
' error reporting utilities
'*************************************************************************************************************************************************************************************************************************************************

Private Function ModuleSummary() As Variant()
    ModuleSummary = Array(1, GLOBAL_PROJECT_NAME, MODULE_NAME, MODULE_VERSION)
End Function


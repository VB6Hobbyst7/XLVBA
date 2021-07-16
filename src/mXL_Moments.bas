Attribute VB_Name = "mXL_Moments"
'************************************************************************************************************************************************
'
'    Copyright (c) 2009-2011, David Briant. All rights reserved.
'    Licensed under BSD 3-Clause License - see https://github.com/DangerMouseB
'
'************************************************************************************************************************************************

Option Explicit

Private Declare Function uDBMoments1D Lib "QUtils" (ByRef anArray2D() As Double) As Double()

Function XLMoments(data As Range) As Double()
    Dim data2D() As Double, i As Long, j As Long
    ReDim data2D(1 To data.Rows.Count, 1 To data.Columns.Count)
    For i = 1 To data.Rows.Count
        For j = 1 To data.Columns.Count
            data2D(i, j) = data(i, j).Value
        Next
    Next
    XLMoments = uDBMoments1D(data2D)
End Function


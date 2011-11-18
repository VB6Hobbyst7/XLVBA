Attribute VB_Name = "mXL_Moments"
'*************************************************************************************************************************************************************************************************************************************************
'            COPYRIGHT NOTICE
'
' Copyright (c) David Briant 2009-2011 - All rights reserved
'
'*************************************************************************************************************************************************************************************************************************************************
 
Option Explicit

Private Declare Function uTRAMoments1D Lib "QUtils" (ByRef anArray2D() As Double) As Double()

Function XLMoments(data As Range) As Double()
    Dim data2D() As Double, i As Long, j As Long
    ReDim data2D(1 To data.Rows.Count, 1 To data.Columns.Count)
    For i = 1 To data.Rows.Count
        For j = 1 To data.Columns.Count
            data2D(i, j) = data(i, j).Value
        Next
    Next
    XLMoments = uTRAMoments1D(data2D)
End Function


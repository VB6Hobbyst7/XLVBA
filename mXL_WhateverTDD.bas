Attribute VB_Name = "mXL_WhateverTDD"
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

Function XLWhatever2DMissingTest(Optional fred As Variant) As String
    Dim fred2D() As Variant, i1 As Long, i2 As Long, j1 As Long, j2 As Long, retval As Long
    retval = DBCWhateverAs2DArray(fred, fred2D, i1, i2, j1, j2)
    If retval = mXL_VBInterface.WHATEVER_IS_MISSING Then XLWhatever2DMissingTest = "Passed": Exit Function
    XLWhatever2DMissingTest = "Failed"
End Function

Function XLWhatever2DEmptyTest(Optional fred As Variant = Empty) As String
    Dim fred2D() As Variant, i1 As Long, i2 As Long, j1 As Long, j2 As Long, retval As Long
    retval = DBCWhateverAs2DArray(fred, fred2D, i1, i2, j1, j2)
    If retval = mXL_VBInterface.WHATEVER_IS_EMPTY Then XLWhatever2DEmptyTest = "Passed": Exit Function
    XLWhatever2DEmptyTest = "Failed"
End Function

Function XLWhatever2DEmptyTestWrapper() As String
    XLWhatever2DEmptyTestWrapper = XLWhatever2DEmptyTest(Empty)
End Function

Function XLWhatever2DErrorTest(Optional fred As Variant = Empty) As String
    Dim fred2D() As Variant, i1 As Long, i2 As Long, j1 As Long, j2 As Long, retval As Long
    retval = DBCWhateverAs2DArray(fred, fred2D, i1, i2, j1, j2)
    If retval = mXL_VBInterface.WHATEVER_IS_ERROR Then XLWhatever2DErrorTest = "Passed": Exit Function
    XLWhatever2DErrorTest = "Failed"
End Function

Function XLWhatever2DTest(fred As Variant, ei1 As Long, ei2 As Long, ej1 As Long, ej2 As Long, Optional defaultHorizontal = True) As String
    Dim fred2D() As Variant, i1 As Long, i2 As Long, j1 As Long, j2 As Long, retval As Long
    XLWhatever2DTest = "Passed"
    retval = DBCWhateverAs2DArray(fred, fred2D, i1, i2, j1, j2, defaultHorizontal)
    If retval <> 0 Then XLWhatever2DTest = "Failed": Exit Function
    If i1 <> ei1 Or i2 <> ei2 Or j1 <> ej1 Or j2 <> ej2 Then XLWhatever2DTest = "Failed": Exit Function
End Function


Function XLWhatever1DMissingTest(Optional fred As Variant) As String
    Dim fred1D() As Variant, i1 As Long, i2 As Long, retval As Long
    retval = DBCWhateverAs1DArray(fred, fred1D, i1, i2)
    If retval = mXL_VBInterface.WHATEVER_IS_MISSING Then XLWhatever1DMissingTest = "Passed": Exit Function
    XLWhatever1DMissingTest = "Failed"
End Function

Function XLWhatever1DEmptyTest(Optional fred As Variant = Empty) As String
    Dim fred1D() As Variant, i1 As Long, i2 As Long, retval As Long
    retval = DBCWhateverAs1DArray(fred, fred1D, i1, i2)
    If retval = mXL_VBInterface.WHATEVER_IS_EMPTY Then XLWhatever1DEmptyTest = "Passed": Exit Function
    XLWhatever1DEmptyTest = "Failed"
End Function

Function XLWhatever1DEmptyTestWrapper() As String
    XLWhatever1DEmptyTestWrapper = XLWhatever1DEmptyTest(Empty)
End Function

Function XLWhatever1DErrorTest(Optional fred As Variant = Empty) As String
    Dim fred1D() As Variant, i1 As Long, i2 As Long, retval As Long
    retval = DBCWhateverAs1DArray(fred, fred1D, i1, i2)
    If retval = mXL_VBInterface.WHATEVER_IS_ERROR Then XLWhatever1DErrorTest = "Passed": Exit Function
    XLWhatever1DErrorTest = "Failed"
End Function

Function XLWhatever1DTest(fred As Variant, ei1 As Long, ei2 As Long) As String
    Dim fred1D() As Variant, i1 As Long, i2 As Long, retval As Long
    XLWhatever1DTest = "Passed"
    retval = DBCWhateverAs1DArray(fred, fred1D, i1, i2)
    If retval <> 0 Then XLWhatever1DTest = "Failed": Exit Function
    If i1 <> ei1 Or i2 <> ei2 Then XLWhatever1DTest = "Failed": Exit Function
End Function


Function XLWhatever2DDMissingTest(Optional fred As Variant) As String
    Dim fred2D() As Double, i1 As Long, i2 As Long, j1 As Long, j2 As Long, retval As Long
    retval = DBCWhateverAs2DArrayD(fred, fred2D, i1, i2, j1, j2)
    If retval = mXL_VBInterface.WHATEVER_IS_MISSING Then XLWhatever2DDMissingTest = "Passed": Exit Function
    XLWhatever2DDMissingTest = "Failed"
End Function

Function XLWhatever2DDEmptyTest(Optional fred As Variant = Empty) As String
    Dim fred2D() As Double, i1 As Long, i2 As Long, j1 As Long, j2 As Long, retval As Long
    retval = DBCWhateverAs2DArrayD(fred, fred2D, i1, i2, j1, j2)
    If retval = mXL_VBInterface.WHATEVER_IS_EMPTY Then XLWhatever2DDEmptyTest = "Passed": Exit Function
    XLWhatever2DDEmptyTest = "Failed"
End Function

Function XLWhatever2DDEmptyTestWrapper() As String
    XLWhatever2DDEmptyTestWrapper = XLWhatever2DDEmptyTest(Empty)
End Function

Function XLWhatever2DDErrorTest(Optional fred As Variant = Empty) As String
    Dim fred2D() As Double, i1 As Long, i2 As Long, j1 As Long, j2 As Long, retval As Long
    retval = DBCWhateverAs2DArrayD(fred, fred2D, i1, i2, j1, j2)
    If retval = mXL_VBInterface.WHATEVER_IS_ERROR Then XLWhatever2DDErrorTest = "Passed": Exit Function
    XLWhatever2DDErrorTest = "Failed"
End Function

Function XLWhatever2DDTest(fred As Variant, ei1 As Long, ei2 As Long, ej1 As Long, ej2 As Long, Optional defaultHorizontal = True) As String
    Dim fred2D() As Double, i1 As Long, i2 As Long, j1 As Long, j2 As Long, retval As Long
    XLWhatever2DDTest = "Passed"
    retval = DBCWhateverAs2DArrayD(fred, fred2D, i1, i2, j1, j2, defaultHorizontal)
    If retval <> 0 Then XLWhatever2DDTest = "Failed": Exit Function
    If i1 <> ei1 Or i2 <> ei2 Or j1 <> ej1 Or j2 <> ej2 Then XLWhatever2DDTest = "Failed": Exit Function
End Function


Function XLWhatever1DDMissingTest(Optional fred As Variant) As String
    Dim fred1D() As Double, i1 As Long, i2 As Long, retval As Long
    retval = DBCWhateverAs1DArrayD(fred, fred1D, i1, i2)
    If retval = mXL_VBInterface.WHATEVER_IS_MISSING Then XLWhatever1DDMissingTest = "Passed": Exit Function
    XLWhatever1DDMissingTest = "Failed"
End Function

Function XLWhatever1DDEmptyTest(Optional fred As Variant = Empty) As String
    Dim fred1D() As Double, i1 As Long, i2 As Long, retval As Long
    retval = DBCWhateverAs1DArrayD(fred, fred1D, i1, i2)
    If retval = mXL_VBInterface.WHATEVER_IS_EMPTY Then XLWhatever1DDEmptyTest = "Passed": Exit Function
    XLWhatever1DDEmptyTest = "Failed"
End Function

Function XLWhatever1DDEmptyTestWrapper() As String
    XLWhatever1DDEmptyTestWrapper = XLWhatever1DDEmptyTest(Empty)
End Function

Function XLWhatever1DDErrorTest(Optional fred As Variant = Empty) As String
    Dim fred1D() As Double, i1 As Long, i2 As Long, retval As Long
    retval = DBCWhateverAs1DArrayD(fred, fred1D, i1, i2)
    If retval = mXL_VBInterface.WHATEVER_IS_ERROR Then XLWhatever1DDErrorTest = "Passed": Exit Function
    XLWhatever1DDErrorTest = "Failed"
End Function

Function XLWhatever1DDTest(fred As Variant, ei1 As Long, ei2 As Long) As String
    Dim fred1D() As Double, i1 As Long, i2 As Long, retval As Long
    XLWhatever1DDTest = "Passed"
    retval = DBCWhateverAs1DArrayD(fred, fred1D, i1, i2)
    If retval <> 0 Then XLWhatever1DDTest = "Failed": Exit Function
    If i1 <> ei1 Or i2 <> ei2 Then XLWhatever1DDTest = "Failed": Exit Function
End Function



Function XLWhateverTestString() As String
    XLWhateverTestString = "hello"
End Function

Function XLWhateverTestArrayOfStrings() As Variant
    XLWhateverTestArrayOfStrings = Array("hello", "there")
End Function

Function XLWhateverTestArrayOfDoubles() As Variant
    XLWhateverTestArrayOfDoubles = Array(3#, 1#)
End Function

Function XLWhateverTestEmpty() As Variant
End Function

Function XLWhateverTestReturnOptional(Optional whatever As Variant) As Variant
    XLWhateverTestReturnOptional = whatever
End Function


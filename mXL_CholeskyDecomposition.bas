Attribute VB_Name = "mXL_CholeskyDecomposition"
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

Private Declare Function uNRBRF_newToken Lib "QUtils" (tolerance As Double, lower As Double, fLower As Double, UPPER As Double, fUpper As Double) As Double()

Private Declare Function uNRCholeskyDecomposition Lib "QUtils" (io_aMatrix() As Double, n As Long, o_pVector2D() As Double) As HRESULT
Private Declare Function uNRCholeskySolve Lib "QUtils" (aMatrix() As Double, n As Long, pVector2D() As Double, bVector2D() As Double, ioXVector2D() As Double) As HRESULT

Private Const A_IS_NOT_PD_MATRIX = vbObjectError + 30001     '  Matrix a, with rounding errors, is not positive definite

Function XLCholeskyDecomposition(aMatrix As Variant) As Variant
    Dim aMatrix2D() As Double, answer2D() As Double, n As Long, pVector2D() As Double, i As Long, j As Long, retVal As Long
    retVal = DBCWhateverAs2DArrayD(aMatrix, aMatrix2D, , n)
    If retVal <> 0 Then XLCholeskyDecomposition = "#Invalid matrix!": Exit Function
    DBCreateNewArrayOfDoubles pVector2D, 1, n, 1, 1
    DBCreateNewArrayOfDoubles answer2D, 1, n, 1, n + 1
    If uNRCholeskyDecomposition(aMatrix2D, n, pVector2D).HRESULT <> S_OK Then XLCholeskyDecomposition = answer2D: Exit Function
    For i = 1 To n
        For j = 1 To n
            answer2D(i, j) = aMatrix2D(i, j)
        Next
        answer2D(i, n + 1) = pVector2D(i, 1)
    Next
    XLCholeskyDecomposition = answer2D
End Function

Function XLCholeskySolve(aMatrix As Variant, pVector As Variant, yVector As Variant) As Variant
    Dim aMatrix2D() As Double, pVector2D() As Double, yVector2D() As Double, xVector2D() As Double, n As Long, retVal As Long
    retVal = DBCWhateverAs2DArrayD(aMatrix, aMatrix2D, , n)
    If retVal <> 0 Then XLCholeskySolve = "#Invalid matrix!": Exit Function
    retVal = DBCWhateverAs2DArrayD(pVector, pVector2D)
    If retVal <> 0 Then XLCholeskySolve = "#Invalid pVector!": Exit Function
    retVal = DBCWhateverAs2DArrayD(yVector, yVector2D)
    If retVal <> 0 Then XLCholeskySolve = "#Invalid yVector!": Exit Function
    DBCreateNewArrayOfDoubles xVector2D, 1, n, 1, 1
    uNRCholeskySolve aMatrix2D, n, pVector2D, yVector2D, xVector2D
    XLCholeskySolve = xVector2D
End Function



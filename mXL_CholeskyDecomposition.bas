Attribute VB_Name = "mXL_CholeskyDecomposition"
Option Explicit

Private Const A_IS_NOT_PD_MATRIX = vbObjectError + 30001     '  Matrix a, with rounding errors, is not positive definite

Function XLCholeskyDecomposition(aMatrix As Variant) As Variant
    Dim aMatrix2D() As Double, answer2D() As Double, n As Long, pVector2D() As Double, i As Long, j As Long, retval As Long
    retval = DBCWhateverAs2DArrayD(aMatrix, aMatrix2D, , n)
    If retval <> 0 Then XLCholeskyDecomposition = "#Invalid matrix!": Exit Function
    DBCreateNewArrayOfDoubles pVector2D, 1, n, 1, 1
    DBCreateNewArrayOfDoubles answer2D, 1, n, 1, n + 1
    If NRCholeskyDecomposition(aMatrix2D, n, pVector2D).HRESULT <> S_OK Then XLCholeskyDecomposition = answer2D: Exit Function
    For i = 1 To n
        For j = 1 To n
            answer2D(i, j) = aMatrix2D(i, j)
        Next
        answer2D(i, n + 1) = pVector2D(i, 1)
    Next
    XLCholeskyDecomposition = answer2D
End Function

Function XLCholeskySolve(aMatrix As Variant, pVector As Variant, yVector As Variant) As Variant
    Dim aMatrix2D() As Double, pVector2D() As Double, yVector2D() As Double, xVector2D() As Double, n As Long, retval As Long
    retval = DBCWhateverAs2DArrayD(aMatrix, aMatrix2D, , n)
    If retval <> 0 Then XLCholeskySolve = "#Invalid matrix!": Exit Function
    retval = DBCWhateverAs2DArrayD(pVector, pVector2D)
    If retval <> 0 Then XLCholeskySolve = "#Invalid pVector!": Exit Function
    retval = DBCWhateverAs2DArrayD(yVector, yVector2D)
    If retval <> 0 Then XLCholeskySolve = "#Invalid yVector!": Exit Function
    DBCreateNewArrayOfDoubles xVector2D, 1, n, 1, 1
    NRCholeskySolve aMatrix2D, n, pVector2D, yVector2D, xVector2D
    XLCholeskySolve = xVector2D
End Function



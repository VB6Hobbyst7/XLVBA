Attribute VB_Name = "mXL_VBInterface"
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
Option Private Module

' error reporting
Private Const MODULE_NAME As String = "mXL_VBInterface"
Private Const MODULE_VERSION As String = "0.0.0.1"

Public Const WHATEVER_IS_MISSING As Long = 1
Public Const WHATEVER_IS_EMPTY As Long = 2
Public Const WHATEVER_IS_ERROR As Long = 3
Public Const WHATEVER_IS_BADLY_DIMENSIONED As Long = 4
Public Const WHATEVER_IS_NOT_VECTOR As Long = 5
Public Const WHATEVER_IS_NOT_NUMERIC As Long = 6


'*************************************************************************************************************************************************************************************************************************************************
' Coerce stuff from XL into nice VB arrays
'*************************************************************************************************************************************************************************************************************************************************

Function DBCWhateverAs2DArray(whatever As Variant, oAnswer2D() As Variant, Optional i1 As Long, Optional i2 As Long, Optional j1 As Long, Optional j2 As Long, Optional defaultHorrizontal = True) As Long
    Dim arrayOrVariant As Variant, nDimensions As Long, i As Long, vt As Long
    
    If IsMissing(whatever) Then DBCWhateverAs2DArray = WHATEVER_IS_MISSING: Exit Function
    If IsEmpty(whatever) Then DBCWhateverAs2DArray = WHATEVER_IS_EMPTY: Exit Function
    If IsError(whatever) Then DBCWhateverAs2DArray = WHATEVER_IS_ERROR: Exit Function

    If TypeName(whatever) = "Range" Then
        arrayOrVariant = whatever.value
    Else
        arrayOrVariant = whatever
    End If

    vt = varType(arrayOrVariant)
    If vt < vbArray Then
        DBCreateNewArrayOfVariants oAnswer2D, 1, 1, 1, 1
        oAnswer2D(1, 1) = arrayOrVariant
        i1 = 1
        i2 = 1
        j1 = 1
        j2 = 1
        Exit Function
    End If
    
    DBGetArrayDetails arrayOrVariant, nDimensions
    DBGetArrayBounds arrayOrVariant, 1, i1, i2
    Select Case nDimensions
        Case 1
            If defaultHorrizontal Then
                DBCreateNewArrayOfVariants oAnswer2D, 1, 1, i1, i2
                For i = i1 To i2
                    oAnswer2D(1, i) = arrayOrVariant(i)
                Next
                j1 = i1
                j2 = i2
                i1 = 1
                i2 = 1
            Else
                DBCreateNewArrayOfVariants oAnswer2D, i1, i2, 1, 1
                For i = i1 To i2
                    oAnswer2D(i, 1) = arrayOrVariant(i)
                Next
                j1 = 1
                j2 = 1
            End If
        Case 2
            DBGetArrayBounds arrayOrVariant, 2, j1, j2
            oAnswer2D = arrayOrVariant
        Case Else
            DBCWhateverAs2DArray = WHATEVER_IS_BADLY_DIMENSIONED: Exit Function
    End Select
    
End Function



Function DBCWhateverAs1DArray(whatever As Variant, oAnswer1D() As Variant, Optional i1 As Long, Optional i2 As Long) As Long
    Dim arrayOrVariant As Variant, nDimensions As Long, i As Long, j1 As Long, j2 As Long, j As Long, vt As Long
    
    If IsMissing(whatever) Then DBCWhateverAs1DArray = WHATEVER_IS_MISSING: Exit Function
    If IsEmpty(whatever) Then DBCWhateverAs1DArray = WHATEVER_IS_EMPTY: Exit Function
    If IsError(whatever) Then DBCWhateverAs1DArray = WHATEVER_IS_ERROR: Exit Function

    If TypeName(whatever) = "Range" Then
        arrayOrVariant = whatever.value
    Else
        arrayOrVariant = whatever
    End If

    vt = varType(arrayOrVariant)
    If vt < vbArray Then
        DBCreateNewArrayOfVariants oAnswer1D, 1, 1
        oAnswer1D(1) = arrayOrVariant
        i1 = 1
        i2 = 1
        Exit Function
    End If
    
    DBGetArrayDetails arrayOrVariant, nDimensions
    DBGetArrayBounds arrayOrVariant, 1, i1, i2
    Select Case nDimensions
        Case 1
            oAnswer1D = arrayOrVariant
        Case 2
            DBGetArrayBounds arrayOrVariant, 2, j1, j2
            Select Case True
                Case i1 = 1 And i2 = 1
                    DBCreateNewArrayOfVariants oAnswer1D, j1, j2
                    For j = j1 To j2
                        oAnswer1D(j) = arrayOrVariant(1, j)
                    Next
                    i1 = j1
                    i2 = j2
                Case j1 = 1 And j2 = 1
                    DBCreateNewArrayOfVariants oAnswer1D, i1, i2
                    For i = i1 To i2
                        oAnswer1D(i) = arrayOrVariant(i, 1)
                    Next
                Case Else
                    DBCWhateverAs1DArray = WHATEVER_IS_NOT_VECTOR: Exit Function
            End Select
        Case Else
            DBCWhateverAs1DArray = WHATEVER_IS_BADLY_DIMENSIONED: Exit Function
    End Select
    
End Function



Function DBCWhateverAs2DArrayD(whatever As Variant, oAnswer2D() As Double, Optional i1 As Long, Optional i2 As Long, Optional j1 As Long, Optional j2 As Long, Optional defaultHorrizontal = True) As Long
    Dim arrayOrVariant As Variant, nDimensions As Long, i As Long, j As Long, vt As Long
    
    If IsMissing(whatever) Then DBCWhateverAs2DArrayD = WHATEVER_IS_MISSING: Exit Function
    If IsEmpty(whatever) Then DBCWhateverAs2DArrayD = WHATEVER_IS_EMPTY: Exit Function
    If IsError(whatever) Then DBCWhateverAs2DArrayD = WHATEVER_IS_ERROR: Exit Function

    If TypeName(whatever) = "Range" Then
        arrayOrVariant = whatever.value
    Else
        arrayOrVariant = whatever
    End If

    vt = varType(arrayOrVariant)
    If vt < vbArray Then
        DBCreateNewArrayOfDoubles oAnswer2D, 1, 1, 1, 1
        On Error GoTo exceptionHandler
        oAnswer2D(1, 1) = CDbl(arrayOrVariant)
        i1 = 1
        i2 = 1
        j1 = 1
        j2 = 1
        Exit Function
    End If
    
    DBGetArrayDetails arrayOrVariant, nDimensions
    If nDimensions = 2 And vt = (vbArray Or vbDouble) Then
        On Error GoTo exceptionHandler
        oAnswer2D = arrayOrVariant
        DBGetArrayBounds arrayOrVariant, 1, i1, i2
        DBGetArrayBounds arrayOrVariant, 2, j1, j2
        Exit Function
    End If
    
    DBGetArrayBounds arrayOrVariant, 1, i1, i2
    Select Case nDimensions
        Case 1
            If defaultHorrizontal Then
                DBCreateNewArrayOfDoubles oAnswer2D, 1, 1, i1, i2
                On Error GoTo exceptionHandler
                For i = i1 To i2
                    oAnswer2D(1, i) = CDbl(arrayOrVariant(i))
                Next
                j1 = i1
                j2 = i2
                i1 = 1
                i2 = 1
                Exit Function
            Else
                DBCreateNewArrayOfDoubles oAnswer2D, i1, i2, 1, 1
                On Error GoTo exceptionHandler
                For i = i1 To i2
                    oAnswer2D(i, 1) = CDbl(arrayOrVariant(i))
                Next
                j1 = 1
                j2 = 1
                Exit Function
            End If
        Case 2
            DBGetArrayBounds arrayOrVariant, 2, j1, j2
            DBCreateNewArrayOfDoubles oAnswer2D, i1, i2, j1, j2
            On Error GoTo exceptionHandler
            For i = i1 To i2
                For j = j1 To j2
                    oAnswer2D(i, j) = CDbl(arrayOrVariant(i, j))
                Next
            Next
            Exit Function
        Case Else
            DBCWhateverAs2DArrayD = WHATEVER_IS_BADLY_DIMENSIONED: Exit Function
    End Select

exceptionHandler:
    DBCWhateverAs2DArrayD = WHATEVER_IS_NOT_NUMERIC
End Function



Function DBCWhateverAs1DArrayD(whatever As Variant, oAnswer1D() As Double, Optional i1 As Long, Optional i2 As Long) As Long
    Dim arrayOrVariant As Variant, nDimensions As Long, i As Long, j1 As Long, j2 As Long, j As Long, vt As Long
    
    If IsMissing(whatever) Then DBCWhateverAs1DArrayD = WHATEVER_IS_MISSING: Exit Function
    If IsEmpty(whatever) Then DBCWhateverAs1DArrayD = WHATEVER_IS_EMPTY: Exit Function
    If IsError(whatever) Then DBCWhateverAs1DArrayD = WHATEVER_IS_ERROR: Exit Function

    If TypeName(whatever) = "Range" Then
        arrayOrVariant = whatever.value
    Else
        arrayOrVariant = whatever
    End If

    vt = varType(arrayOrVariant)
    If vt < vbArray Then
        DBCreateNewArrayOfDoubles oAnswer1D, 1, 1
        On Error GoTo exceptionHandler
        oAnswer1D(1) = CDbl(arrayOrVariant)
        i1 = 1
        i2 = 1
        Exit Function
    End If
    
    DBGetArrayDetails arrayOrVariant, nDimensions
    If nDimensions = 1 And vt = (vbArray Or vbDouble) Then
        On Error GoTo exceptionHandler
        oAnswer1D = arrayOrVariant
        DBGetArrayBounds arrayOrVariant, 1, i1, i2
        Exit Function
    End If
        
    DBGetArrayDetails arrayOrVariant, nDimensions
    DBGetArrayBounds arrayOrVariant, 1, i1, i2
    Select Case nDimensions
        Case 1
            DBCreateNewArrayOfDoubles oAnswer1D, i1, i2
            On Error GoTo exceptionHandler
            For i = i1 To i2
                oAnswer1D(i) = CDbl(arrayOrVariant(i))
            Next
            Exit Function
        Case 2
            DBGetArrayBounds arrayOrVariant, 2, j1, j2
            Select Case True
                Case i1 = 1 And i2 = 1
                    DBCreateNewArrayOfDoubles oAnswer1D, j1, j2
                    On Error GoTo exceptionHandler
                    For j = j1 To j2
                        oAnswer1D(j) = CDbl(arrayOrVariant(1, j))
                    Next
                    i1 = j1
                    i2 = j2
                    Exit Function
                Case j1 = 1 And j2 = 1
                    DBCreateNewArrayOfDoubles oAnswer1D, i1, i2
                    On Error GoTo exceptionHandler
                    For i = i1 To i2
                        oAnswer1D(i) = CDbl(arrayOrVariant(i, 1))
                    Next
                    Exit Function
                Case Else
                    DBCWhateverAs1DArrayD = WHATEVER_IS_NOT_VECTOR: Exit Function
            End Select
        Case Else
            DBCWhateverAs1DArrayD = WHATEVER_IS_BADLY_DIMENSIONED: Exit Function
    End Select

Exit Function
exceptionHandler:
    DBCWhateverAs1DArrayD = WHATEVER_IS_NOT_NUMERIC
End Function


'*************************************************************************************************************************************************************************************************************************************************
' Utilities to help make the interface to functions being called by XL and VBA more uniform
'*************************************************************************************************************************************************************************************************************************************************

Function DBRows(variantOrRange2D As Variant) As Variant
    If TypeName(variantOrRange2D) = "Range" Then
        DBRows = variantOrRange2D.rows.count
    Else
        DBRows = UBound(variantOrRange2D, 1)
    End If
End Function

Function DBCols(variantOrRange2D As Variant) As Variant
    If TypeName(variantOrRange2D) = "Range" Then
        DBCols = variantOrRange2D.columns.count
    Else
        DBCols = UBound(variantOrRange2D, 2)
    End If
End Function


'*************************************************************************************************************************************************************************************************************************************************
' module summary
'*************************************************************************************************************************************************************************************************************************************************

Private Function ModuleSummary() As Variant()
    ModuleSummary = Array(1, GLOBAL_PROJECT_NAME, MODULE_NAME, MODULE_VERSION)
End Function



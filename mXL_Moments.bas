Attribute VB_Name = "mXL_Moments"
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


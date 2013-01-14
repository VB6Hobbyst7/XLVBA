Attribute VB_Name = "mBoiler_AutoOpen"
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

'Sub auto_open()
    'Dim hModule As Long
    'initializeSpecialShortCutKeys
    'hModule = DLLHandle("QUtils", "c:\dev\bin\QUtils.dll")
    'If hModule = 0 Then MsgBox "Couldn't load QUtils!"
'End Sub


'*************************************************************************************************************************************************************************************************************************************************
' initialise - short cut keys
'*************************************************************************************************************************************************************************************************************************************************

Sub initializeSpecialShortCutKeys()
    Application.OnKey "^+v", "XLPasteSpecialValues"
    Application.OnKey "^+f", "XLPasteSpecialFormulas"
    Application.OnKey "^+b", "XLPasteSpecialFormats"
    Application.OnKey "^c", "XLNoteAddressAndContinueCopying"
    Application.OnKey "^+a", "XLPasteSpecialAddress"
    Application.OnKey "^+w", "XLPasteSpecialGrids"
    Application.OnKey "^+e", "XLPasteSpecialGridBlankInside"
    Application.OnKey "^+n", "XLPasteSpecialNow"
End Sub

Sub uninitializeSpecialShortCutKeys()
    Application.OnKey "^+v"
    Application.OnKey "^+f"
    Application.OnKey "^+b"
    Application.OnKey "^c"
    Application.OnKey "^+a"
    Application.OnKey "^+w"
    Application.OnKey "^+e"
    Application.OnKey "^+n"
End Sub



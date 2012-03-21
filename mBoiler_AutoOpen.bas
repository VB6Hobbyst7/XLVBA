Attribute VB_Name = "mBoiler_AutoOpen"
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



Attribute VB_Name = "mBoiler_ProjectNameVBA"
Option Explicit
Option Private Module

Function GLOBAL_PROJECT_NAME() As String
    GLOBAL_PROJECT_NAME = Application.ThisWorkbook.name
End Function


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTyper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : cTyper
' Author    : Michiel van der Blonk (blonkm@gmail.com)
' Date      : 3/18/2015
' Purpose   : find out what type an office file is (word, excel, powerpoint)
'---------------------------------------------------------------------------------------

Option Explicit

Private ctlFiletypes As Collection

Public Enum eFileType
    [_first] = 1
    tWord = 1
    tExcel = 2
    tPowerpoint = 3
    tAll = 4
    [_last] = 4
End Enum

Const sFiletypeWord As String = "Word 2007 Document"
Const sFiletypeExcel As String = "Microsoft Office Excel 2007 Workbook"
Const sFiletypePowerpoint As String = "Microsoft Office PowerPoint 2007 Presentation"
Const sFiletypeAll As String = ""

Public Function FileTypeName(Filetype As eFileType) As String
    FileTypeName = ctlFiletypes(Filetype)
End Function

Public Function getType(objFile As file) As eFileType
    Dim n As eFileType
    
    getType = tAll
    For n = eFileType.[_first] To eFileType.[_last]
        If objFile.Type = ctlFiletypes(n) Then
            getType = n
        End If
    Next
End Function

Private Sub Class_Initialize()
    Set ctlFiletypes = New Collection
    ctlFiletypes.Add sFiletypeWord, "Word"
    ctlFiletypes.Add sFiletypeExcel, "Excel"
    ctlFiletypes.Add sFiletypePowerpoint, "Powerpoint"
    ctlFiletypes.Add sFiletypeAll, ""
End Sub

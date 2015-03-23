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

Private colExtensions As Collection

Private objWordDocument As cOfficeWordDocument
Private objExcelDocument As cOfficeExcelDocument
Private objPowerpointDocument As cOfficePowerPointDocument

Public Enum eFileType
    [_first] = 1
    tWord = 1
    tExcel = 2
    tPowerpoint = 3
    tAll = 4
    [_last] = 4
End Enum

Const sFiletypeWord As String = "docx"
Const sFiletypeExcel As String = "xlsx"
Const sFiletypePowerpoint As String = "pptx"
Const sFiletypeAll As String = ""


'---------------------------------------------------------------------------------------
' Procedure : Extension
' Author    : Michiel van der Blonk (blonkm@gmail.com)
' Date      : 3/22/2015
' Purpose   : get the extension for filetype, e.g. tWord => "docx"
'---------------------------------------------------------------------------------------
'
Public Function Extension(Filetype As eFileType) As String
    Extension = colExtensions(Filetype)
End Function


'---------------------------------------------------------------------------------------
' Procedure : getType
' Author    : Michiel van der Blonk (blonkm@gmail.com)
' Date      : 3/22/2015
' Purpose   : get file type for a file. E.g. "file.docx" => tWord
'---------------------------------------------------------------------------------------
'
Public Function getType(ByVal objFile As Object) As eFileType
    Dim n As eFileType
    Dim objFS ' As FileSystemObject
    
    Set objFS = CreateObject("Scripting.FileSystemObject")
    getType = tAll
    For n = eFileType.[_first] To eFileType.[_last]
        If objFS.GetExtensionName(objFile) = colExtensions(n) Then
            getType = n
        End If
    Next
End Function

'---------------------------------------------------------------------------------------
' Procedure : IsOfficeDocument
' Author    : Michiel van der Blonk (blonkm@gmail.com)
' Date      : 3/22/2015
' Purpose   : check if a file is a word, excel or powerpoint file
'             preconditions: last filetype is "*.*"
'---------------------------------------------------------------------------------------
'
Public Function IsOfficeDocument(objFile As Object) As Boolean
    IsOfficeDocument = getType(objFile) < eFileType.[_last]
End Function


'---------------------------------------------------------------------------------------
' Procedure : OfficeDocument
' Author    : Michiel van der Blonk (blonkm@gmail.com)
' Date      : 3/22/2015
' Purpose   : get object corresponding to file type
'---------------------------------------------------------------------------------------
'
Public Function OfficeDocument(Filetype As eFileType) As IOfficeDocument
    Select Case Filetype
        Case tWord
            Set OfficeDocument = objWordDocument
        Case tExcel
            Set OfficeDocument = objExcelDocument
        Case tPowerpoint
            Set OfficeDocument = objPowerpointDocument
    End Select
End Function

Private Sub Class_Initialize()
    Set colExtensions = New Collection
    colExtensions.Add sFiletypeWord, "Word"
    colExtensions.Add sFiletypeExcel, "Excel"
    colExtensions.Add sFiletypePowerpoint, "Powerpoint"
    colExtensions.Add sFiletypeAll, ""
    
    ' prepare office document objects
    Set objWordDocument = New cOfficeWordDocument
    Set objExcelDocument = New cOfficeExcelDocument
    Set objPowerpointDocument = New cOfficePowerPointDocument
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : cFolder
' Author    : Michiel van der Blonk (blonkm@gmail.com)
' Date      : 3/18/2015
' Purpose   : interface to files collection in a folder
'---------------------------------------------------------------------------------------

Option Explicit

Public Filetype As eFileType
Private m_path As String
Private m_files As Collection

'---------------------------------------------------------------------------------------
' Procedure : Files
' Author    : Michiel van der Blonk (blonkm@gmail.com)
' Date      : 3/18/2015
' Purpose   : retrieve files in a folder of a certain type (word, excel, ppt)
'---------------------------------------------------------------------------------------
Public Property Get Files() As Collection
    Dim objFS ' As FileSystemObject
    Dim objFolder As Folder
    Dim objFile As File
    Dim objTyper As cTyper
    Dim bAddFile As Boolean
    
    If m_files Is Nothing Then
        Set objFS = CreateObject("Scripting.FileSystemObject")
        Set objTyper = New cTyper
        Set m_files = New Collection

        If objFS.FolderExists(Path) Then
            Set objFolder = objFS.GetFolder(Path)
            For Each objFile In objFolder.Files
                ' ensure this is an office document or a selected type
                If Filetype = tAll Then
                    bAddFile = objTyper.IsOfficeDocument(objFile)
                Else
                    bAddFile = objTyper.getType(objFile) = Filetype
                End If
                If bAddFile Then
                    m_files.Add objFile
                End If
            Next
        End If
    End If
    Set Files = m_files
End Property

Public Property Get Path() As Variant
    Path = m_path
End Property


'---------------------------------------------------------------------------------------
' Procedure : Path
' Author    : Michiel van der Blonk (blonkm@gmail.com)
' Date      : 3/22/2015
' Purpose   : change path, and repopulate file list if necessary
'---------------------------------------------------------------------------------------
'
Public Property Let Path(ByVal vNewValue As Variant)
    ' when path changes, repopulate the file list
    If vNewValue <> m_path Then
        Dim temp As Object
        m_path = vNewValue
        Set m_files = Nothing
        Set temp = Files ' repopulate
    End If
End Property

Private Sub Class_Initialize()
    Set m_files = Nothing
    Filetype = tAll

    If m_path = "" Then
        m_path = ActiveDocument.Path
    End If
End Sub

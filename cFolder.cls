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
Public m_path As String

Private m_files As Collection

'---------------------------------------------------------------------------------------
' Procedure : Files
' Author    : Michiel van der Blonk (blonkm@gmail.com)
' Date      : 3/18/2015
' Purpose   : retrieve files in a folder of a certain type (word, excel, ppt)
'---------------------------------------------------------------------------------------
Public Property Get Files() As Collection
    Dim objFS As New FileSystemObject
    Dim objFolder As Folder
    Dim objFile As file
    Dim objTyper As cTyper
    
    If m_files Is Nothing Then

        Set objTyper = New cTyper
        Set m_files = New Collection

        If objFS.FolderExists(Path) Then
            Set objFolder = objFS.GetFolder(Path)
            For Each objFile In objFolder.Files
                If objTyper.getType(objFile) = Filetype Then
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

Public Property Let Path(ByVal vNewValue As Variant)
    If vNewValue <> m_path Then
        Dim temp As Object
        m_path = vNewValue
        Set temp = Files
    End If
End Property

Private Sub Class_Initialize()
    Set m_files = Nothing
    Filetype = tAll

    If m_path = "" Then
        m_path = ActiveDocument.Path
    End If
End Sub


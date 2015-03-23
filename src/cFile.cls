VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : cFile
' Author    : Michiel van der Blonk (blonkm@gmail.com)
' Date      : 3/18/2015
' Purpose   : Interface to getting fonts in use in a file (word, excel, ppt)
'---------------------------------------------------------------------------------------

Option Explicit

Private m_filetype As eFileType
Private m_path As String
Private m_fonts As Collection
Private objTyper As cTyper

Public includeInstalled As Boolean
Public oFile ' As Scripting.File
Public objDocument As IOfficeDocument

'---------------------------------------------------------------------------------------
' Procedure : PrintFonts
' Author    : Michiel van der Blonk (blonkm@gmail.com)
' Date      : 3/22/2015
' Purpose   : print fonts used in this file
'---------------------------------------------------------------------------------------
'
Public Sub PrintFonts()
    Dim vFont As Variant
    Dim colFonts As Collection
    
    Set colFonts = Fonts
    For Each vFont In colFonts
        Debug.Print vFont
    Next
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Filetype
' Author    : Michiel van der Blonk (blonkm@gmail.com)
' Date      : 3/22/2015
' Purpose   : get the filetype of an office document (tWord, tExcel, tPowerpoint)
'---------------------------------------------------------------------------------------
'
Public Property Get Filetype() As eFileType
    Filetype = objTyper.getType(oFile)
End Property

'---------------------------------------------------------------------------------------
' Procedure : Fonts
' Author    : Michiel van der Blonk (blonkm@gmail.com)
' Date      : 3/22/2015
' Purpose   : collect all fonts from the office document in member oFile
'---------------------------------------------------------------------------------------
'
Public Property Get Fonts() As Collection
    ' check type
    If oFile Is Nothing Then
        Set m_fonts = Nothing
    Else
        Set objDocument = objTyper.OfficeDocument(Filetype)
        objDocument.includeInstalled = includeInstalled
        ' use singleton
        If m_fonts Is Nothing Then
            Set m_fonts = New Collection
            Set m_fonts = objDocument.GetFonts(oFile)
        End If
    End If
    Set Fonts = m_fonts
End Property

Private Sub Class_Initialize()
    Set objTyper = New cTyper
    includeInstalled = False
End Sub

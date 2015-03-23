VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cOfficeWordDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : cOfficeWordDocument
' Author    : Michiel van der Blonk (blonkm@gmail.com)
' Date      : 3/22/2015
' Purpose   : list all fonts in a Word document
'---------------------------------------------------------------------------------------

Option Explicit

Implements IOfficeDocument

Private m_includeInstalled As Boolean

'---------------------------------------------------------------------------------------
' Procedure : GetFonts
' Author    : Michiel van der Blonk (blonkm@gmail.com)
' Date      : 3/18/2015
' Purpose   : Get all fonts in a Word document file
'---------------------------------------------------------------------------------------
'
Public Function IOfficeDocument_GetFonts(ByVal oFile As Object) As Collection
    Dim objParagraph    ' As Paragraph
    Dim sFontType       As String
    Dim isInstalled     As Boolean
    Dim objShape        ' As Word.Shape
    Dim objDocument     ' As Word.Document
    Dim i               As Integer
    
    Dim colFonts As Collection
    Set colFonts = New Collection
        
    On Error GoTo GetFonts_Error
    
    Set objDocument = GetObject(oFile.Path)
    'regular text
    For Each objParagraph In objDocument.Paragraphs
        For i = 1 To objParagraph.Range.Characters.Count
            sFontType = GetFontName(objParagraph.Range.Characters(i).Font)
            If IsValidFont(sFontType, m_includeInstalled) Then
                colFonts.Add sFontType, sFontType
            End If
        Next
    Next
    
    'shapes
    For Each objShape In objDocument.Shapes
        If objShape.Type = msoAutoShape Or objShape.Type = msoGroup Or objShape.Type = msoTextBox Then
            If objShape.TextFrame.HasText Then
                sFontType = GetFontName(objShape.TextFrame.TextRange.Font)
                If objShape.TextFrame.HasText Then
                    If IsValidFont(sFontType, m_includeInstalled) Then
                        colFonts.Add sFontType, sFontType
                    End If
                End If
            End If
        End If
    Next
    
    objDocument.Close SaveChanges:=False
    Set objDocument = Nothing
    Set IOfficeDocument_GetFonts = colFonts
    
    Exit Function
GetFonts_Error:
    If Err.Number = 457 Then ' key already exists
        Resume Next
    Else
        MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetFonts of Class Module cOfficeWordDocument"
    End If
End Function

Public Property Get IOfficeDocument_includeInstalled() As Boolean
    IOfficeDocument_includeInstalled = m_includeInstalled
End Property

Public Property Let IOfficeDocument_includeInstalled(ByVal bNewValue As Boolean)
    m_includeInstalled = bNewValue
End Property

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cOfficePowerPointDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : cOfficePowerPointDocument
' Author    : Michiel van der Blonk (blonkm@gmail.com)
' Date      : 3/22/2015
' Purpose   : list all fonts in a Powerpoint document
'---------------------------------------------------------------------------------------

Option Explicit

Implements IOfficeDocument

Private m_includeInstalled As Boolean

'---------------------------------------------------------------------------------------
' Procedure : GetFonts
' Author    : Michiel van der Blonk (blonkm@gmail.com)
' Date      : 3/18/2015
' Purpose   : Get all fonts in a PowerPoint file
'---------------------------------------------------------------------------------------
'
Public Function IOfficeDocument_GetFonts(ByVal oFile As Object) As Collection
    Dim objSlide ' as slide
    Dim objShape ' As Shape
    Dim objPresentation ' As PowerPoint.Presentation
    Dim sFontType As String
    Dim isInstalled As Boolean
    Dim colFonts As Collection
    
    On Error GoTo GetFonts_Error
    
    Set colFonts = New Collection
    Set objPresentation = GetObject(oFile.Path)
    
    For Each objSlide In objPresentation.Slides
        For Each objShape In objSlide.Shapes
            If objShape.HasTextFrame Then
                sFontType = GetFontName(objShape.TextFrame.TextRange.Font.Name)
                If objShape.TextFrame.HasText Then
                    If IsValidFont(sFontType, m_includeInstalled) Then
                        colFonts.Add sFontType, sFontType
                    End If
                End If
            End If
        Next
    Next
    objPresentation.Close
    Set objPresentation = Nothing
    Set IOfficeDocument_GetFonts = colFonts
    
    Exit Function
GetFonts_Error:
    If Err.Number = 457 Then ' key already exists
        Resume Next
    Else
        MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetFonts of Class Module cOfficePowerPointDocument"
    End If
End Function


Public Property Get IOfficeDocument_includeInstalled() As Boolean
    IOfficeDocument_includeInstalled = m_includeInstalled
End Property

Public Property Let IOfficeDocument_includeInstalled(ByVal bNewValue As Boolean)
    m_includeInstalled = bNewValue
End Property

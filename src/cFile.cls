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

Public IncludeInstalled As Boolean
Public oFile As Scripting.file

Public Sub PrintFonts()
    Dim vFont As Variant
    Dim colFonts As Collection
    
    Set colFonts = Fonts
    For Each vFont In colFonts
        Debug.Print vFont
    Next
End Sub

Public Property Get Filetype() As eFileType
    Filetype = objTyper.getType(oFile)
End Property

Public Property Get Fonts() As Collection
    ' use cached member
    If m_fonts Is Nothing Then
        Set m_fonts = New Collection
        
        Select Case Filetype
            Case tWord
                Set m_fonts = GetWordFonts(oFile)
            Case tExcel
                Set m_fonts = GetXLFonts(oFile)
            Case tPowerpoint
                Set m_fonts = GetPPTFonts(oFile)
            Case tAll
        End Select
    End If
    Set Fonts = m_fonts
End Property

Private Function GetWordFonts(oFile As file) As Collection
    Dim objParagraph    As Paragraph
    Dim sFontType       As String
    Dim isInstalled     As Boolean
    Dim objShape        As Word.Shape
    Dim objDocument     As Word.Document
    Dim i               As Integer
    
    Dim colFonts As Collection
    Set colFonts = New Collection
        
    On Error GoTo errHandler
    
    Set objDocument = GetObject(oFile.Path)
    'regular text
    For Each objParagraph In objDocument.Paragraphs
        For i = 1 To objParagraph.Range.Characters.Count
            sFontType = GetFontName(objParagraph.Range.Characters(i).Font)
            If IsValidFont(sFontType) Then
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
                    If IsValidFont(sFontType) Then
                        colFonts.Add sFontType, sFontType
                    End If
                End If
            End If
        End If
    Next
    
    objDocument.Close saveChanges:=False
    Set objDocument = Nothing
    Set GetWordFonts = colFonts
    Exit Function
errHandler:
    If Err.Number = 457 Then ' key already exists
        Resume Next
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetXLFonts
' Author    : Michiel van der Blonk (blonkm@gmail.com)
' Date      : 3/18/2015
' Purpose   : Get all fonts in an Excel file
' Caveat    : this is slow since it loops through all cells
'---------------------------------------------------------------------------------------
'
Private Function GetXLFonts(objFile As file) As Collection

    On Error GoTo GetXLFonts_Error

    Dim x As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorkSheet As Excel.Worksheet
    Dim sFontType As String
    Dim isInstalled As Boolean
    Dim colFonts As Collection
    Dim objCell As Variant

    Set colFonts = New Collection
    Set GetXLFonts = colFonts

    Set objWorkbook = GetObject(oFile.Path)

    For Each objWorkSheet In objWorkbook.Worksheets
        For Each objCell In objWorkSheet.UsedRange.Cells
            sFontType = objCell.Font.Name
            If IsValidFont(sFontType) Then
                colFonts.Add sFontType, sFontType
            End If
        Next
    Next

    On Error GoTo 0
    Exit Function

GetXLFonts_Error:
    If Err.Number = 457 Then    ' key already exists
        Resume Next
    Else
        MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetXLFonts of Class Module cFile"
    End If
End Function

Private Function GetPPTFonts(oFile As file) As Collection
    Dim objSlide ' as slide
    Dim objShape ' As Shape
    Dim p ' As PowerPoint.Presentation
    Dim sFontType As String
    Dim isInstalled As Boolean
    Dim colFonts As Collection
    
    On Error GoTo errHandler
    
    Set colFonts = New Collection
    Set p = GetObject(oFile.Path)
    
    For Each objSlide In p.Slides
        For Each objShape In objSlide.Shapes
            If objShape.HasTextFrame Then
                sFontType = GetFontName(objShape.TextFrame.TextRange.Font.Name)
                If objShape.TextFrame.HasText Then
                    If IsValidFont(sFontType) Then
                        colFonts.Add sFontType, sFontType
                    End If
                End If
            End If
        Next
    Next
    p.Close
    Set p = Nothing
    Set GetPPTFonts = colFonts
    
    Exit Function
errHandler:
    If Err.Number = 457 Then ' key already exists
        Resume Next
    End If
End Function

Private Function GetFontName(objFont As Font) As String
    GetFontName = ""
    If objFont.NameOther <> "" Then
        GetFontName = objFont.NameOther
    End If
    If objFont.NameFarEast <> "" Then
        GetFontName = objFont.NameFarEast
    End If
    If objFont.NameBi <> "" Then
        GetFontName = objFont.NameBi
    End If
    If objFont.NameAscii <> "" Then
        GetFontName = objFont.NameAscii
    End If
    If objFont.Name <> "" Then
        GetFontName = objFont.Name
    End If
End Function

Private Function IsValidFont(sFont As String) As Boolean
    Dim isInstalled As Boolean
    
    IsValidFont = False
    isInstalled = IsFontInstalled(sFont)
    
    ' this happens sometimes
    ' don't know why yet
    If sFont = "" Then
        IsValidFont = False
        Exit Function
    End If
    ' include installed means this is valid
    If (isInstalled And IncludeInstalled) Then
        IsValidFont = True
    End If
    ' uninstalled fonts should always be in the list
    If Not isInstalled Then
        IsValidFont = True
    End If
End Function

Private Sub Class_Initialize()
    Set objTyper = New cTyper
    IncludeInstalled = False
End Sub

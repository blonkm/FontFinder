Attribute VB_Name = "modFonts"
'---------------------------------------------------------------------------------------
' Module    : modFonts
' Author    : Michiel van der Blonk (blonkm@gmail.com)
' Date      : 3/22/2015
' Purpose   : some font related routines
'---------------------------------------------------------------------------------------

Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : PrintInstalledFonts
' Author    : Michiel van der Blonk (blonkm@gmail.com)
' Date      : 3/22/2015
' Purpose   : print a list of installed fonts, using the MS Word FontNames collection
'---------------------------------------------------------------------------------------
'
Public Sub PrintInstalledFonts()
    Dim f

    For Each f In Application.FontNames
        Debug.Print f
    Next
End Sub

'---------------------------------------------------------------------------------------
' Procedure : IsFontInstalled
' Author    : Michiel van der Blonk (blonkm@gmail.com)
' Date      : 3/22/2015
' Purpose   : check if a font is installed in the Windows system fonts
'             this is done by name, so should be an exact match
'---------------------------------------------------------------------------------------
'
Public Function IsFontInstalled(fontname As String) As Boolean
    Dim f
    Dim ret As Boolean

    ret = False
    For Each f In Application.FontNames
        If f = fontname Then
            ret = True
            Exit For
        End If
    Next
    IsFontInstalled = ret
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetFontName
' Author    : Michiel van der Blonk (blonkm@gmail.com)
' Date      : 3/22/2015
' Purpose   : get the name of the font in a font object
'             note that this is a tricky subject
'             apparently the font naem can be located in many properties
'---------------------------------------------------------------------------------------
'
Public Function GetFontName(objFont As Font) As String
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

'---------------------------------------------------------------------------------------
' Procedure : IsValidFont
' Author    : Michiel van der Blonk (blonkm@gmail.com)
' Date      : 3/22/2015
' Purpose   : should this font be added to the list
'---------------------------------------------------------------------------------------
'
Public Function IsValidFont(sFont As String, includeInstalled As Boolean) As Boolean
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
    If (isInstalled And includeInstalled) Then
        IsValidFont = True
    End If
    ' uninstalled fonts should always be in the list
    If Not isInstalled Then
        IsValidFont = True
    End If
End Function

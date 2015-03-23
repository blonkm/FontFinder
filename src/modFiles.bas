Attribute VB_Name = "modFiles"
'---------------------------------------------------------------------------------------
' Module    : modFiles
' Author    : Michiel van der Blonk (blonkm@gmail.com)
' Date      : 3/18/2015
' Purpose   : print fonts of several files in a folder
'---------------------------------------------------------------------------------------

Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : PrintFontsInFolder
' Author    : Michiel van der Blonk (blonkm@gmail.com)
' Date      : 3/18/2015
' Purpose   : print fonts of several files in a folder
'---------------------------------------------------------------------------------------
'
Sub PrintFontsInFolder(Filetype As eFileType, Optional Path)
    Dim objFolder As cFolder
    Dim objFile As File
    Dim objCFile As cFile
    Dim vFont As Variant
    Dim colFonts As Collection
    
    Set objFolder = New cFolder
    objFolder.Filetype = Filetype
    If IsMissing(Path) Then
        objFolder.Path = ActiveDocument.Path
    Else
        objFolder.Path = Path
    End If
       
    ' collect files
    For Each objFile In objFolder.Files
        Set objCFile = New cFile
        Set objCFile.oFile = objFile
        objCFile.includeInstalled = True
        ' collect fonts
        Set colFonts = objCFile.Fonts
        For Each vFont In colFonts
            Debug.Print objFile.Name, vFont
        Next
    Next
End Sub

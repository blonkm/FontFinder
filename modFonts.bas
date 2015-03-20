Attribute VB_Name = "modFonts"
Option Explicit

Public Sub PrintInstalledFonts()
    Dim f

    For Each f In Application.FontNames
        Debug.Print f
    Next
End Sub

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

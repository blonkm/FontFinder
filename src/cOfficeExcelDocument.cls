VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cOfficeExcelDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : cOfficeExcelDocument
' Author    : Michiel van der Blonk (blonkm@gmail.com)
' Date      : 3/22/2015
' Purpose   : list fonts in an Excel document
'---------------------------------------------------------------------------------------

Option Explicit

Implements IOfficeDocument

Private m_includeInstalled As Boolean

'---------------------------------------------------------------------------------------
' Procedure : GetFonts
' Author    : Michiel van der Blonk (blonkm@gmail.com)
' Date      : 3/18/2015
' Purpose   : Get all fonts in an Excel file
' Caveat    : this is slow since it loops through all cells
'---------------------------------------------------------------------------------------
'
Public Function IOfficeDocument_GetFonts(ByVal objFile As Object) As Collection

    On Error GoTo GetFonts_Error

    Dim objWorkbook ' As Excel.Workbook
    Dim objWorkSheet ' As Excel.Worksheet
    Dim sFontType As String
    Dim isInstalled As Boolean
    Dim colFonts As Collection
    Dim objCell As Variant

    Set colFonts = New Collection

    Set objWorkbook = GetObject(objFile.Path)

    For Each objWorkSheet In objWorkbook.Worksheets
        For Each objCell In objWorkSheet.UsedRange.Cells
            sFontType = objCell.Font.Name
            If IsValidFont(sFontType, m_includeInstalled) Then
                colFonts.Add sFontType, sFontType
            End If
        Next
    Next
    
    objWorkbook.Close SaveChanges:=False
    Set objWorkbook = Nothing
    Set IOfficeDocument_GetFonts = colFonts
    
    On Error GoTo 0
    Exit Function

GetFonts_Error:
    If Err.Number = 457 Then    ' key already exists
        Resume Next
    Else
        MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetFonts of Class Module cOfficeExcelDocument"
    End If
End Function

Public Property Get IOfficeDocument_includeInstalled() As Boolean
    IOfficeDocument_includeInstalled = m_includeInstalled
End Property

Public Property Let IOfficeDocument_includeInstalled(ByVal bNewValue As Boolean)
    m_includeInstalled = bNewValue
End Property

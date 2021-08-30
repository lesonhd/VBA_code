Attribute VB_Name = "tao_muc_luc"
Option Explicit

Sub link_sheets()
    Dim m_sheet As Worksheet
    Dim i As Integer
    Dim j As Integer

Set m_sheet = ThisWorkbook.Worksheets(1)
j = 1
For i = 2 To ThisWorkbook.Worksheets.Count
    m_sheet.Hyperlinks.Add anchor:=m_sheet.Cells(j, 4), Address:="", SubAddress:= _
    ThisWorkbook.Sheets(i).Name & "!A1", TextToDisplay:=ThisWorkbook.Sheets(i).Name
    j = j + 1
Next i
Set m_sheet = Nothing
End Sub


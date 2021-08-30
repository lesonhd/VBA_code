VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_BeforeClose(Cancel As Boolean)
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Dim t As Integer
' Dat mot bien de kiem tra lua chon cua nguoi dung
   t = MsgBox(Prompt:="Do you want to create a daily report?", Buttons:=vbYesNo, Title:="Create Report")

   If t = 6 Then
   ThisWorkbook.Save
   ThisWorkbook.Sheets(3).Delete
   ThisWorkbook.Sheets(3).Delete
   ThisWorkbook.Sheets(3).Delete
    ThisWorkbook.SaveAs Filename:=ThisWorkbook.Path & "\" & _
    "CMS-Update-" & Format(Date, "dd-mm-yyyy") & ".xlsx", FileFormat:= _
    xlWorkbookDefault, CreateBackup:=False
   End If
Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub



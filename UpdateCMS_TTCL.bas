Attribute VB_Name = "UpdateCMS_TTCL"
Option Explicit

Sub update_cmsTTCL()
    Dim myData As Worksheet
    Dim TTCL_data As Worksheet
    Dim TTCL_path As String
    
    
' mo file cua ttcl
TTCL_path = "E:\6. Long Son\2.QC Data\DailyWeldRecord 4-Aug-20.xlsx"
Application.Workbooks.Open TTCL_path
' dat ten cho cac sheet
    Set myData = ThisWorkbook.Sheets(1)
    Set TTCL_data = ActiveWorkbook.Sheets(1)
    

    
    
End Sub

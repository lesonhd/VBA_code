Attribute VB_Name = "CreateVisual"
Option Explicit
Dim mReport As Worksheet
Dim mData As Worksheet
Dim lrData As Long
Dim lrReport_list As Integer
Dim i_data As Long
Dim k_report_list As Integer
Dim j_report As Integer
Dim mReport_path As String
Dim SelectedFile As Object
Dim j_an_dong As Integer
Dim report_name As String


Sub chon_file()
    Set SelectedFile = Application.FileDialog(msoFileDialogFilePicker)
    
    SelectedFile.Show
    On Error GoTo c
    mReport_path = SelectedFile.SelectedItems(1)
    
   ' MsgBox mReport_path
c:
Set SelectedFile = Nothing
End Sub


Sub create_Visual()
    Call chon_file
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.CalculateBeforeSave = True

Set mData = ThisWorkbook.Sheets(1)


' tim dong cuoi sheet data
lrData = mData.Range("I" & Columns("I").Rows.Count).End(xlUp).Row
lrReport_list = ThisWorkbook.Sheets(3).Range("B" & Columns("B").Rows.Count).End(xlUp).Row

For k_report_list = 2 To lrReport_list
' On Error GoTo k
Workbooks.Open mReport_path

Set mReport = ActiveWorkbook.Sheets(3)
j_report = 19

For i_data = 7 To lrData
    If mData.Cells(i_data, 27).Value = ThisWorkbook.Sheets(3).Cells(k_report_list, 2).Value Then
    ' them cac thong tin chinh
        ' so report
        mReport.Range("J7").Value = mData.Cells(i_data, 27).Value
        ' Ngay report
        mReport.Range("E15").Value = mData.Cells(i_data, 26).Value
        ' Area
        mReport.Range("A14").Value = "Area: " & mData.Cells(i_data, 2).Value
        
    ' noi dung report
        ' ban ve 2
    mReport.Cells(j_report, 2).Value = mData.Cells(i_data, 9).Value
        'Sheet 3
    mReport.Cells(j_report, 3).Value = mData.Cells(i_data, 11).Value
        'Rev 4
    mReport.Cells(j_report, 4).Value = mData.Cells(i_data, 10).Value
        ' Joint 5
    mReport.Cells(j_report, 5).Value = mData.Cells(i_data, 12).Value
        'Size 6
    mReport.Cells(j_report, 6).Value = mData.Cells(i_data, 15).Value
        'type 7
    mReport.Cells(j_report, 7).Value = mData.Cells(i_data, 19).Value
        'root 8
    mReport.Cells(j_report, 8).Value = mData.Cells(i_data, 24).Value
        ' Cover 9
    mReport.Cells(j_report, 9).Value = mData.Cells(i_data, 25).Value
        ' Spool 10
    mReport.Cells(j_report, 10).Value = mData.Cells(i_data, 8).Value
        ' shop field 11
    mReport.Cells(j_report, 11).Value = mData.Cells(i_data, 20).Value
        ' DB 12
    mReport.Cells(j_report, 12).Value = mData.Cells(i_data, 16).Value
    
    j_report = j_report + 1
    
    End If

Next i_data

' An di dong trong trong file report
For j_an_dong = 55 To 148
    If mReport.Cells(j_an_dong, 2).Value = "" Then
        mReport.Rows(j_an_dong).EntireRow.Hidden = True
    End If
Next j_an_dong

'Copy sang sheet welding report
'   mReport.Range("B19:L" & j_report).Copy ActiveWorkbook.Sheets(3).Range("B19")
 '  For j_an_dong = 55 To 148
 '   If ActiveWorkbook.Sheets(3).Cells(j_an_dong, 2).Value = "" Then
 '       ActiveWorkbook.Sheets(3).Rows(j_an_dong).EntireRow.Hidden = True
'    End If
'Next j_an_dong

' Luu thanh file moi
report_name = Right(ThisWorkbook.Sheets(3).Cells(k_report_list, 2).Value, 18)
ActiveWorkbook.SaveAs "E:\6. Long Son\2.QC Data\1.WEC\" & report_name & ".xlsx"
' Luu thanh file pdf
ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        "E:\6. Long Son\2.QC Data\1.WEC\" & report_name & ".pdf", Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:= _
        False
ActiveWorkbook.Close

Set mReport = Nothing

Next k_report_list

'k:
Set mData = Nothing

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.CalculateBeforeSave = False
MsgBox "Tao thanh cong " & lrReport_list - 1 & " bao cao Visual"
End Sub

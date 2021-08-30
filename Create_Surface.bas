Attribute VB_Name = "Create_Surface"
Option Explicit

Dim mReport As Worksheet
Dim mData As Worksheet
Dim lrData As Long
Dim lrReport_list As Integer
Dim i_data As Long
Dim k_report_list As Integer
Dim j_report As Integer
Dim mReport_path As String
Dim myFolder_path As String
Dim SelectedFile As Object
Dim j_an_dong As Integer
Dim report_name As String


Sub chon_file()
    MsgBox " Chon Form Mau Surface Preparation"
    Set SelectedFile = Application.FileDialog(msoFileDialogFilePicker)
    
    SelectedFile.Show
    On Error GoTo c
    mReport_path = SelectedFile.SelectedItems(1)
    
   ' MsgBox mReport_path
c:
Set SelectedFile = Nothing
End Sub
Sub Chon_Folder()
' Chon 1 folder va tra ve duong dan den folder do.
MsgBox "Chon Folder Luu File"
Application.FileDialog(msoFileDialogFolderPicker).Show
On Error GoTo kcfd
    myFolder_path = Application.FileDialog(msoFileDialogFolderPicker).SelectedItems(1)
    
kcfd:
End Sub


Sub creat_Surface()
    Call chon_file
    Call Chon_Folder
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.CalculateBeforeSave = True

Set mData = ThisWorkbook.Sheets(6)


' tim dong cuoi sheet data
lrData = mData.Range("E" & Columns("E").Rows.Count).End(xlUp).Row
lrReport_list = ThisWorkbook.Sheets(3).Range("C" & Columns("C").Rows.Count).End(xlUp).Row

For k_report_list = 2 To lrReport_list
' On Error GoTo k
Workbooks.Open mReport_path

Set mReport = ActiveWorkbook.Sheets(3)
j_report = 18

For i_data = 3 To lrData
    If mData.Cells(i_data, 17).Value = ThisWorkbook.Sheets(3).Cells(k_report_list, 3).Value Then
    ' them cac thong tin chinh
        mReport.Range("Y8").Value = mData.Cells(i_data, 17).Value 'So bao cao
        mReport.Range("S15").Value = mData.Cells(i_data, 16).Value 'Ngay bao cao
        
        
    ' noi dung report
    mReport.Cells(j_report, 4).Value = mData.Cells(i_data, 3).Value ' Drawing
    mReport.Cells(j_report, 17).Value = mData.Cells(i_data, 4).Value ' Rev
    mReport.Cells(j_report, 18).Value = mData.Cells(i_data, 7).Value ' Paint system
    mReport.Cells(j_report, 19).Value = mData.Cells(i_data, 6).Value ' Size
    mReport.Cells(j_report, 21).Value = mData.Cells(i_data, 8).Value / 1000 ' lenth
    mReport.Cells(j_report, 23).Value = mData.Cells(i_data, 5).Value ' Spool No
    mReport.Cells(j_report, 28).Value = mData.Cells(i_data, 9).Value ' Area m2
        
    j_report = j_report + 1
    
    End If

Next i_data

' An di dong trong trong file report
For j_an_dong = 58 To 200
    If mReport.Cells(j_an_dong, 4).Value = "" Then
        mReport.Rows(j_an_dong).EntireRow.Hidden = True
    End If
Next j_an_dong

' Luu thanh file moi
report_name = Right(ThisWorkbook.Sheets(3).Cells(k_report_list, 3).Value, 18)
ActiveWorkbook.SaveAs myFolder_path & "\" & report_name & ".xlsx"

' Luu thanh file pdf
ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        myFolder_path & "\" & report_name & ".pdf", Quality:=xlQualityStandard, _
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
MsgBox "Tao thanh cong " & lrReport_list - 1 & " bao cao Surface Preparation"
End Sub

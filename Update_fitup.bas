Attribute VB_Name = "Update_fitup"
Option Explicit
    Dim data_ws As Worksheet
    Dim file_nguon As Worksheet
    Dim file_nguon_path As String
    Dim file_nguon_folder As String
    Dim file_nguon_name As String
    Dim get_file_nguon As Office.FileDialog
    Dim report_num As String
    Dim report_text As String
    Dim report_date As Date
    Dim i_nguon As Integer
    Dim lr_nguon As Integer
    Dim j_data As Long
    Dim lr_data As Long
    Dim spool_nguon As String
    Dim spool_data As String
    Dim joint_nguon As String
    Dim joint_data As String
    Dim dwsheet_nguon As String
    Dim dwsheet_data As String
    Dim dw_nguon As String
    Dim dw_data As String
    Dim db_nguon As Long
    Dim db_data As Long
    Dim data_cell_rp As String
    Dim data_cell_date As String
    Dim joint_update_count As Long
    Dim trung_spool As New Collection
    Dim trung_joint As New Collection
    Dim trung_dwsheet As New Collection
    Dim trung_rp As New Collection
    Dim trung_date As New Collection
    Dim trung_count As Integer
    Dim joint_nguon_count As Integer
    Dim Cong_ty As String
    
Sub Chon_Folder()
   Set get_file_nguon = Application.FileDialog(msoFileDialogFolderPicker)
   ' Mo ra hop thoai chon folder chua file nguon
   get_file_nguon.Show
   
   ' Chon folder mac dinh
   get_file_nguon.InitialFileName = ThisWorkbook.Path & "\"
   
   ' dat duong dan la folder duoc chon dau tien
   On Error GoTo c
    file_nguon_folder = get_file_nguon.SelectedItems(1) & "\"
    
    ' Lay ten cac file co trong folder
    'file_nguon_name = Dir(file_nguon_folder, vbNormal)
c:
    Set get_file_nguon = Nothing
    
End Sub
Sub general_Fitup()
'On Error GoTo c
        ' dat duong link den file nguon
    file_nguon_path = file_nguon_folder & file_nguon_name
    
    ' mo file nguon
    Workbooks.Open (file_nguon_path)
    
    ' dat lai ten cho file nguon va file data
    Set data_ws = ThisWorkbook.Sheets(1)
    Set file_nguon = ActiveWorkbook.Sheets(1)
    
    ' Lay dong cuoi cua bang data
    lr_data = data_ws.Range("H" & Columns("H").Rows.Count).End(xlUp).Row
    
    ' Lay dong cuoi cua file nguon
    lr_nguon = file_nguon.Range("B" & Columns("B").Rows.Count).End(xlUp).Row
        
    ' Lay Ngay va So bao cao
    report_date = file_nguon.Range("E15").Value
    report_text = file_nguon.Range("I7").Value
    report_num = Trim(Right(report_text, Len(report_text) - 1))
'c:
End Sub
Sub DinhNghia_nguon_fitup()
    spool_nguon = file_nguon.Cells(i_nguon, 8).Value
    joint_nguon = file_nguon.Cells(i_nguon, 5).Value
    dwsheet_nguon = file_nguon.Cells(i_nguon, 3).Value
    dw_nguon = file_nguon.Cells(i_nguon, 2).Value
    db_nguon = file_nguon.Cells(i_nguon, 10).Value
End Sub
Sub DinhNghia_data_fitup()
    spool_data = data_ws.Cells(j_data, 8).Value
    joint_data = data_ws.Cells(j_data, 12).Value
    dwsheet_data = data_ws.Cells(j_data, 11).Value
    dw_data = data_ws.Cells(j_data, 9).Value
    db_data = data_ws.Cells(j_data, 16).Value
    data_cell_rp = data_ws.Cells(j_data, 22).Value
    data_cell_date = data_ws.Cells(j_data, 21).Value
End Sub

Sub Update_fitup()
    
Application.ScreenUpdating = False
Application.DisplayAlerts = False
' chon cong ty de update du lieu
    Cong_ty = Application.inputbox(" Nhap ten cong ty muon update", Type:=2)
' Neu chay rieng thi can goi phuong thuc Chon_folder
'Call Chon_folder
' Dat gia tri cho bien dem so joint
joint_update_count = 0
joint_nguon_count = 0

    ' Lay ten cac file co trong folder
    file_nguon_name = Dir(file_nguon_folder, vbNormal)
    
Do While Len(file_nguon_name) > 0
    Call general_Fitup
    
' Update bao cao
For i_nguon = 19 To lr_nguon
Call DinhNghia_nguon_fitup

    For j_data = 7 To lr_data
    Call DinhNghia_data_fitup
   
        If spool_data = spool_nguon And joint_data = joint_nguon And dw_data = dw_nguon And db_data = db_nguon Then
            data_ws.Cells(j_data, 21).Value = report_date
            data_ws.Cells(j_data, 22).Value = report_num
            data_ws.Cells(j_data, 28).Value = Cong_ty
            joint_update_count = joint_update_count + 1
        End If
        
    Next j_data
    joint_nguon_count = joint_nguon_count + 1
Next i_nguon

' dong file nguon
ActiveWorkbook.Close
' dat lai file nguon = 0
file_nguon_name = Dir
Loop

Set data_ws = Nothing
Set file_nguon = Nothing

Application.ScreenUpdating = True
Application.DisplayAlerts = True
MsgBox "Update thanh cong " & joint_update_count & "/" & joint_nguon_count & " moi"
End Sub

Sub KiemTraTrungDuLieu()
    
Application.ScreenUpdating = False
Application.DisplayAlerts = False
' Neu chay rieng thi can goi phuong thuc Chon_folder
'Call Chon_folder
   ' Lay ten cac file co trong folder
    file_nguon_name = Dir(file_nguon_folder, vbNormal)
    
'dat gia tri cho bien dem so moi bi trung
trung_count = 0

Do While Len(file_nguon_name) > 0
    Call general_Fitup
' Update bao cao
For i_nguon = 19 To lr_nguon
Call DinhNghia_nguon_fitup

    For j_data = 7 To lr_data
        Call DinhNghia_data_fitup
    
' Kiem tra cac dieu kien xem co bi trung khong
        If spool_data = spool_nguon And joint_data = joint_nguon And _
        dwsheet_data = dwsheet_nguon And dw_data = dw_nguon And _
        (data_cell_rp <> "" Or data_cell_date <> "") Then
        
' Lay cac du lieu trung vao cac bo suu tap
            trung_spool.Add Item:=spool_data
            trung_joint.Add Item:=joint_data
            trung_dwsheet.Add Item:=dwsheet_data
            trung_rp.Add Item:=data_cell_rp
            trung_date.Add Item:=data_cell_date
            trung_count = trung_count + 1
            
        End If
        
    Next j_data
    
Next i_nguon

' dong file nguon
ActiveWorkbook.Close

' dat lai file nguon = 0
file_nguon_name = Dir

Loop

' in du lieu trung vao workbook vua tao ra
Dim i As Variant

' Tao bao cao du lieu trung
If trung_count <> 0 Then
' Tao ra workbook moi de luu du lieu trung
Workbooks.Add
ActiveWorkbook.Sheets(1).Range("C1:C" & trung_dwsheet.Count).NumberFormat = "@"
    'in spool trung
    For i = 1 To trung_spool.Count
        ActiveWorkbook.Sheets(1).Cells(i, 1).Value = trung_spool.Item(i)
        
    Next i
    
    ' in join trung
    For i = 1 To trung_joint.Count
        ActiveWorkbook.Sheets(1).Cells(i, 2).Value = trung_joint.Item(i)
    Next i
    
    ' in drawing sheet
    For i = 1 To trung_dwsheet.Count
        ActiveWorkbook.Sheets(1).Cells(i, 3).Value = trung_dwsheet.Item(i)
    Next i
    
    'in so report
    For i = 1 To trung_rp.Count
        ActiveWorkbook.Sheets(1).Cells(i, 4).Value = trung_rp.Item(i)
    Next i
    
    ' in date report trung
    For i = 1 To trung_date.Count
        ActiveWorkbook.Sheets(1).Cells(i, 5).Value = CDate(trung_date.Item(i))
    Next i
    
    'luu file du lieu trung
    ActiveWorkbook.SaveAs ThisWorkbook.Path & "\" & "dulieuFit-uptrung.xlsx"
    ActiveWorkbook.Close
    
    MsgBox "Co : " & trung_count & " Joint bi trung"
    
    Else
    'MsgBox " Khong co bao cao bi trung"
End If
' xoa cac phan tu cua bo suu tap
    For i = 1 To trung_spool.Count
        trung_spool.Remove 1
    Next i
    
    For i = 1 To trung_joint.Count
        trung_joint.Remove 1
    Next i
    
    For i = 1 To trung_dwsheet.Count
        trung_dwsheet.Remove 1
    Next i
    
    For i = 1 To trung_rp.Count
        trung_rp.Remove 1
    Next i
    
    For i = 1 To trung_date.Count
        trung_date.Remove 1
    Next i
    
Set data_ws = Nothing
Set file_nguon = Nothing

Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub
Sub CapNhatFit_up()
   Set get_file_nguon = Application.FileDialog(msoFileDialogFolderPicker)
   ' Mo ra hop thoai chon folder chua file nguon
   get_file_nguon.Show
   
   ' Chon folder mac dinh
   get_file_nguon.InitialFileName = ThisWorkbook.Path & "\"
   
   ' dat duong dan la folder duoc chon dau tien
   On Error GoTo c
    file_nguon_folder = get_file_nguon.SelectedItems(1) & "\"

Call KiemTraTrungDuLieu
If trung_count = 0 Then
    Call Update_fitup
    Else
    MsgBox "chua update bao cao"
    Exit Sub
End If
c:
    Set get_file_nguon = Nothing
    file_nguon_folder = ""
End Sub

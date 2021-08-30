Attribute VB_Name = "Chon_file_va_folder"
Option Explicit
Public myrange As Range
Public mycell As Range
Public myFolder As String

Public Sub range_select()
' Chon 1 vung du lieu bang cua so inputbox, Type:=8 la tra ve gia tri la vung chon
On Error GoTo kcv
Set myrange = Application.InputBox(Prompt:="Nhap vung can chon", Type:=8)
kcv:
End Sub

Sub folder_select()
' Chon 1 folder va tra ve duong dan den folder do.
Application.FileDialog(msoFileDialogFolderPicker).Show
On Error GoTo kcfd
    myFolder = Application.FileDialog(msoFileDialogFolderPicker).SelectedItems(1)
kcfd:
End Sub

Sub colection_subfolder()
' tao bo suu tap cac sub folder trong folder chinh
' sau khi co bo suu tap roi ta có the loop qua tung folder vaf kiem tra cac file trong do
Call folder_select
    Dim sub_collection As New Collection
    Dim folder_name As String
    Dim sub_folder As Variant
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    
folder_name = Dir(myFolder & "\", vbDirectory)
    Do While folder_name <> ""
        Select Case folder_name
        Case Is = ".", ".."
        
        Case Else
            Select Case fso.GetExtensionName(folder_name)
            Case Is <> ""
            
            Case Else
            sub_collection.Add Item:=folder_name
            End Select
        End Select
     folder_name = Dir
    Loop
End Sub
Public Sub selectInputFile()
    Dim fd As Office.FileDialog
  
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
  
      .AllowMultiSelect = False
      .InitialFileName = Application.ActiveWorkbook.Path & "\"
      .Filters.Clear
      .Filters.Add "Excel 2007", "*.xlsx"
      .Filters.Add "All Files", "*.*"
  
      If .Show = True Then
        SelectedFile = .SelectedItems(1)
      End If
    End With
End Sub

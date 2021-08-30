Attribute VB_Name = "Aone_report"
Sub rfikhongtai()
    Dim num_of_rfi As Long
    Dim num_of_column As Long
    Dim i As Long, j As Long
    Dim rfi_khongtai As Object
    Dim t As Object
    Dim k As String
    Dim duong_dan As Variant
    
    
    num_of_column = 9
    num_of_rfi = Workbooks("data_khongtai.xlsx").Sheets("Sheet1").Cells(Rows.Count, "A").End(xlUp).Row - 1
    With CreateObject("word.application")
        .Visible = True
        
        For i = 1 To num_of_rfi
        Set rfi_khongtai = .documents.Open("E:\3.SDWP 1B\7.QC-Construction\Aone\Report\rfi_khongtai.docx")
        Set t = rfi_khongtai.Content
            For j = 1 To num_of_column
               t.Find.Execute _
               FindText:=Workbooks("data_khongtai.xlsx").Sheets("Sheet1").Cells(1, j).Value, _
               ReplaceWith:=Workbooks("data_khongtai.xlsx").Sheets("Sheet1").Cells(i + 1, j).Value, _
               Replace:=wdReplaceAll
               k = Workbooks("data_khongtai.xlsx").Sheets("Sheet1").Cells(i + 1, 1).Value
            Next
            duong_dan = "E:\3.SDWP 1B\7.QC-Construction\Aone\Report\khong_tai\"
            rfi_khongtai.SaveAs2 Filename:=duong_dan & k & ".docx"
        Next
        .Quit
    End With
    Set t = Nothing
    Set rfi_khongtai = Nothing
End Sub

Sub reportKhongtai()
Dim num_of_report As Long
    Dim num_of_column As Long
    Dim i As Long, j As Long
    Dim report_khongtai As Object
    Dim t As Object
    Dim k As String
    Dim duong_dan As Variant
    
    num_of_column = 8
    num_of_report = Workbooks("data_khongtai.xlsx").Sheets("Sheet1").Cells(Rows.Count, "A").End(xlUp).Row - 1
    With CreateObject("word.application")
        .Visible = True
        
        For i = 1 To num_of_report
        Set report_khongtai = .documents.Open("E:\3.SDWP 1B\7.QC-Construction\Aone\Report\report_khongtai.docx")
        Set t = report_khongtai.Content
            For j = 1 To num_of_column
               t.Find.Execute _
               FindText:=Workbooks("data_khongtai.xlsx").Sheets("Sheet1").Cells(1, j).Value, _
               ReplaceWith:=Workbooks("data_khongtai.xlsx").Sheets("Sheet1").Cells(i + 1, j).Value, _
               Replace:=wdReplaceAll
               k = Workbooks("data_khongtai.xlsx").Sheets("Sheet1").Cells(i + 1, 2).Value

            Next
            duong_dan = "E:\3.SDWP 1B\7.QC-Construction\Aone\Report\khong_tai\"
            report_khongtai.SaveAs2 Filename:=duong_dan & k & ".docx"
        Next
        .Quit
    End With
    Set t = Nothing
    Set report_khongtai = Nothing
End Sub
Sub Cotai_rfi()
    Dim num_of_rfi As Long
    Dim num_of_column As Long
    Dim i As Long, j As Long
    Dim Cotai_rfi As Object
    Dim t As Object
    Dim k As String
    Dim duong_dan As Variant
    
    num_of_column = 5
    num_of_rfi = Workbooks("data_khongtai.xlsx").Sheets("Sheet2").Cells(Rows.Count, "A").End(xlUp).Row - 1
    With CreateObject("word.application")
        .Visible = True
        
        For i = 32 To num_of_rfi
        Set Cotai_rfi = .documents.Open("E:\3.SDWP 1B\7.QC-Construction\Aone\Report\cotai_rfi.docx")
        Set t = Cotai_rfi.Content
            For j = 1 To num_of_column
               t.Find.Execute _
               FindText:=Workbooks("data_khongtai.xlsx").Sheets("Sheet2").Cells(1, j).Value, _
               ReplaceWith:=Workbooks("data_khongtai.xlsx").Sheets("Sheet2").Cells(i + 1, j).Value, _
               Replace:=wdReplaceAll
               k = Workbooks("data_khongtai.xlsx").Sheets("Sheet2").Cells(i + 1, 1).Value
            Next
            duong_dan = "E:\3.SDWP 1B\7.QC-Construction\Aone\Report\co_tai\"
            Cotai_rfi.SaveAs2 Filename:=duong_dan & k & ".docx"
        Next
        .Quit
    End With
    Set t = Nothing
    Set Cotai_rfi = Nothing
End Sub
Sub cotai_report()
Dim num_of_report As Long
    Dim num_of_column As Long
    Dim i As Long, j As Long
    Dim cotai_report As Object
    Dim t As Object
    Dim k As String
    Dim duong_dan As Variant
    
    num_of_column = 6
    num_of_report = Workbooks("data_khongtai.xlsx").Sheets("Sheet2").Cells(Rows.Count, "A").End(xlUp).Row - 1
    With CreateObject("word.application")
        .Visible = True
        
        For i = 32 To num_of_report
        Set cotai_report = .documents.Open("E:\3.SDWP 1B\7.QC-Construction\Aone\Report\cotai_report.docx")
        Set t = cotai_report.Content
            For j = 1 To num_of_column
               t.Find.Execute _
               FindText:=Workbooks("data_khongtai.xlsx").Sheets("Sheet2").Cells(1, j).Value, _
               ReplaceWith:=Workbooks("data_khongtai.xlsx").Sheets("Sheet2").Cells(i + 1, j).Value, _
               Replace:=wdReplaceAll
               k = Workbooks("data_khongtai.xlsx").Sheets("Sheet2").Cells(i + 1, 2).Value

            Next
            duong_dan = "E:\3.SDWP 1B\7.QC-Construction\Aone\Report\co_tai\"
            cotai_report.SaveAs2 Filename:=k & ".docx"
        Next
        .Quit
    End With
    Set t = Nothing
    Set cotai_report = Nothing
End Sub
Sub ld_cotai_rfi()
    Dim num_of_rfi As Long
    Dim num_of_column As Long
    Dim i As Long, j As Long
    Dim ld_cotai_rfi As Object
    Dim t As Object
    Dim k As String
    Dim duong_dan As Variant
     
    num_of_column = 9
    num_of_rfi = Workbooks("data_khongtai.xlsx").Sheets("Sheet3").Cells(Rows.Count, "A").End(xlUp).Row - 1
    With CreateObject("word.application")
        .Visible = True
        
        For i = 1 To num_of_rfi
        Set ld_cotai_rfi = .documents.Open("E:\3.SDWP 1B\7.QC-Construction\Aone\Report\ld_cotai_rfi.docx")
        Set t = ld_cotai_rfi.Content
            For j = 1 To num_of_column
               t.Find.Execute _
               FindText:=Workbooks("data_khongtai.xlsx").Sheets("Sheet3").Cells(1, j).Value, _
               ReplaceWith:=Workbooks("data_khongtai.xlsx").Sheets("Sheet3").Cells(i + 1, j).Value, _
               Replace:=wdReplaceAll
               k = Workbooks("data_khongtai.xlsx").Sheets("Sheet3").Cells(i + 1, 1).Value
            Next
            duong_dan = "E:\3.SDWP 1B\7.QC-Construction\Aone\Report\ld_cotai\"
            ld_cotai_rfi.SaveAs2 Filename:=duong_dan & k & ".docx"
        Next
        .Quit
    End With
    Set t = Nothing
    Set ld_cotai_rfi = Nothing
End Sub
Sub ld_cotai()
Dim num_of_report As Long
    Dim num_of_column As Long
    Dim i As Long, j As Long
    Dim ld_cotai As Object
    Dim t As Object
    Dim k As String
    Dim duong_dan As Variant
    
    num_of_column = 11
    num_of_report = Workbooks("data_khongtai.xlsx").Sheets("Sheet3").Cells(Rows.Count, "A").End(xlUp).Row - 1
    With CreateObject("word.application")
        .Visible = True
        
        For i = 1 To num_of_report
        Set ld_cotai = .documents.Open("E:\3.SDWP 1B\7.QC-Construction\Aone\Report\ld_cotai.docx")
        Set t = ld_cotai.Content
            For j = 1 To num_of_column
               t.Find.Execute _
               FindText:=Workbooks("data_khongtai.xlsx").Sheets("Sheet3").Cells(1, j).Value, _
               ReplaceWith:=Workbooks("data_khongtai.xlsx").Sheets("Sheet3").Cells(i + 1, j).Value, _
               Replace:=wdReplaceAll
               k = Workbooks("data_khongtai.xlsx").Sheets("Sheet3").Cells(i + 1, 2).Value

            Next
            duong_dan = "E:\3.SDWP 1B\7.QC-Construction\Aone\Report\ld_cotai\"
            ld_cotai.SaveAs2 Filename:=duong_dan & k & ".docx"
        Next
        .Quit
    End With
    Set t = Nothing
    Set ld_cotai = Nothing
End Sub
Sub ld_khongtai_rfi()
    Dim num_of_rfi As Long
    Dim num_of_column As Long
    Dim i As Long, j As Long
    Dim ld_khongtai_rfi As Object
    Dim t As Object
    Dim k As String
    Dim duong_dan As Variant
    
    num_of_column = 9
    num_of_rfi = Workbooks("data_khongtai.xlsx").Sheets("Sheet4").Cells(Rows.Count, "A").End(xlUp).Row - 1
    With CreateObject("word.application")
        .Visible = True
        
        For i = 1 To num_of_rfi
        Set ld_khongtai_rfi = .documents.Open("E:\3.SDWP 1B\7.QC-Construction\Aone\Report\ld_khongtai_rfi.docx")
        Set t = ld_khongtai_rfi.Content
            For j = 1 To num_of_column
               t.Find.Execute _
               FindText:=Workbooks("data_khongtai.xlsx").Sheets("Sheet4").Cells(1, j).Value, _
               ReplaceWith:=Workbooks("data_khongtai.xlsx").Sheets("Sheet4").Cells(i + 1, j).Value, _
               Replace:=wdReplaceAll
               k = Workbooks("data_khongtai.xlsx").Sheets("Sheet4").Cells(i + 1, 1).Value
            Next
            duong_dan = "E:\3.SDWP 1B\7.QC-Construction\Aone\Report\ld_khongtai\"
            ld_khongtai_rfi.SaveAs2 Filename:=duong_dan & k & ".docx"
        Next
        .Quit
    End With
    Set t = Nothing
    Set ld_khongtai_rfi = Nothing
End Sub
Sub ld_khongtai()
Dim num_of_report As Long
    Dim num_of_column As Long
    Dim i As Long, j As Long
    Dim ld_khongtai As Object
    Dim t As Object
    Dim k As String
    Dim duong_dan As Variant
    
    num_of_column = 11
    num_of_report = Workbooks("data_khongtai.xlsx").Sheets("Sheet4").Cells(Rows.Count, "A").End(xlUp).Row - 1
    With CreateObject("word.application")
        .Visible = True
        
        For i = 1 To num_of_report
        Set ld_khongtai = .documents.Open("E:\3.SDWP 1B\7.QC-Construction\Aone\Report\ld_khongtai.docx")
        Set t = ld_khongtai.Content
            For j = 1 To num_of_column
               t.Find.Execute _
               FindText:=Workbooks("data_khongtai.xlsx").Sheets("Sheet4").Cells(1, j).Value, _
               ReplaceWith:=Workbooks("data_khongtai.xlsx").Sheets("Sheet4").Cells(i + 1, j).Value, _
               Replace:=wdReplaceAll
               k = Workbooks("data_khongtai.xlsx").Sheets("Sheet4").Cells(i + 1, 2).Value

            Next
            duong_dan = "E:\3.SDWP 1B\7.QC-Construction\Aone\Report\ld_khongtai\"
            ld_khongtai.SaveAs2 Filename:=duong_dan & k & ".docx"
        Next
        .Quit
    End With
    Set t = Nothing
    Set ld_khongtai = Nothing
End Sub

Sub attach_dd_khongtai()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
    Dim i As Integer
    Dim k As String
    Dim duong_dan As String
    Dim templ As Object
    
    
    For i = 2 To 42
     Workbooks.Open Filename:= _
        "E:\3.SDWP 1B\7.QC-Construction\Aone\Report\Attach.xlsx"
     
     
    ' Attach 3- Bien ban nghiem thu lap dat
        Workbooks("Attach.xlsx").Sheets(1).Range("C18") = Workbooks("data_khongtai.xlsx").Sheets(1).Cells(i, 10) 'ten thiet bi
        Workbooks("Attach.xlsx").Sheets(1).Range("E18") = Workbooks("data_khongtai.xlsx").Sheets(1).Cells(i, 8) ' so bb nghiem thu 4B
        Workbooks("Attach.xlsx").Sheets(1).Range("G18") = Workbooks("data_khongtai.xlsx").Sheets(1).Cells(i, 13) ' Ngay nghiem thu
        
    ' Attach 2 - Bien ban nghiem thu vat tu
        Workbooks("Attach.xlsx").Sheets(2).Range("C18") = Workbooks("data_khongtai.xlsx").Sheets(1).Cells(i, 10) 'ten thiet bi
        Workbooks("attach.xlsx").Sheets(2).Range("E18") = Workbooks("data_khongtai.xlsx").Sheets(1).Cells(i, 7) ' so bb nghiem thu vat tu
        
    ' Attach 1 - So ban ve
        Workbooks("attach.xlsx").Sheets(3).Range("C18") = Workbooks("data_khongtai.xlsx").Sheets(1).Cells(i, 9) 'ten thiet bi
        
        
    ' Save thanh file moi
        k = Workbooks("data_khongtai.xlsx").Sheets(1).Cells(i, 2).Value
        duong_dan = "E:\3.SDWP 1B\7.QC-Construction\Aone\Report\khong_tai\"
        Windows("attach.xlsx").Activate
        ActiveWorkbook.SaveAs duong_dan & "attach for " & k & ".xlsx"
        ActiveWorkbook.Close
    Next i
Application.DisplayAlerts = True
Application.ScreenUpdating = True
Set templ = Nothing
End Sub

Sub checklist_form()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
    Dim i As Integer
    Dim k As String
    Dim duong_dan As String
    
    
    For i = 2 To 42
     Workbooks.Open Filename:= _
        "E:\3.SDWP 1B\7.QC-Construction\Aone\Report\Checklist_form.xlsx"
      
    ' Dien vao check list
        Workbooks("Checklist_form.xlsx").Sheets(1).Range("J3") = Workbooks("data_khongtai.xlsx").Sheets(1).Cells(i, 6) 'ten he thong
        Workbooks("Checklist_form.xlsx").Sheets(1).Range("B5") = Workbooks("data_khongtai.xlsx").Sheets(1).Cells(i, 10) ' ten thiet bi
              
    ' Save thanh file moi
        k = Workbooks("data_khongtai.xlsx").Sheets(1).Cells(i, 2).Value
        duong_dan = "E:\3.SDWP 1B\7.QC-Construction\Aone\Report\khong_tai\"
        Windows("Checklist_form.xlsx").Activate
        ActiveWorkbook.SaveAs duong_dan & "check list for " & k & ".xlsx"
        ActiveWorkbook.Close
    Next i
Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub

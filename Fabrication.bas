Attribute VB_Name = "Fabrication"
Sub Fitup()
'unhinde sheet
Application.ScreenUpdating = False
Workbooks("Report.xlsx").Sheets("Fit-up").Activate
Columns("A:A").Select
Selection.EntireRow.Hidden = False
' clear content
Range("B14:L400").Select
Selection.ClearContents
' Make report
Dim i As Integer
Dim j As Integer
j = 14
Workbooks("CMS.xlsx").Sheets("CMS").Activate
For i = 1 To 1000
If Workbooks("CMS.xlsx").Sheets("CMS").Cells(i, 10) = Workbooks("Report.xlsx").Sheets("Fit-up").Range("M1") Then
Workbooks("Report.xlsx").Sheets("Fit-up").Cells(j, 2).Value = Workbooks("CMS.xlsx").Sheets("CMS").Cells(i, 2)
Workbooks("Report.xlsx").Sheets("Fit-up").Cells(j, 4).Value = Workbooks("CMS.xlsx").Sheets("CMS").Cells(i, 4)
Workbooks("Report.xlsx").Sheets("Fit-up").Cells(j, 6).Value = Workbooks("CMS.xlsx").Sheets("CMS").Cells(i, 5)
Workbooks("Report.xlsx").Sheets("Fit-up").Cells(j, 7).Value = Workbooks("CMS.xlsx").Sheets("CMS").Cells(i, 19)
Workbooks("Report.xlsx").Sheets("Fit-up").Cells(j, 8).Value = Workbooks("CMS.xlsx").Sheets("CMS").Cells(i, 7)
Workbooks("Report.xlsx").Sheets("Fit-up").Cells(j, 9).Value = Workbooks("CMS.xlsx").Sheets("CMS").Cells(i, 8)
Workbooks("Report.xlsx").Sheets("Fit-up").Cells(j, 10).Value = "3~5"
Workbooks("Report.xlsx").Sheets("Fit-up").Cells(j, 11).Value = "ACC"
Workbooks("Report.xlsx").Sheets("Fit-up").Range("M14").Value = Workbooks("CMS.xlsx").Sheets("CMS").Cells(i, 9)
Workbooks("Report.xlsx").Sheets("Fit-up").Range("Q8").Value = Workbooks("CMS.xlsx").Sheets("CMS").Cells(i, 18)
Workbooks("Report.xlsx").Sheets("Fit-up").Range("M8").Value = Workbooks("CMS.xlsx").Sheets("CMS").Cells(i, 10)
j = j + 1
End If
Next i
' hidden row
Workbooks("Report.xlsx").Sheets("Fit-up").Activate
Dim h As Integer
For h = 14 To 399
If Cells(h, 2).Value = "" Then
Cells(h, 2).EntireRow.Hidden = True
End If
Next h
Application.ScreenUpdating = True
End Sub

Sub VISUAL()
' unhinde Sheet
Application.ScreenUpdating = False
Workbooks("Report.xlsx").Sheets("Visual").Activate
Columns("A:A").Select
Selection.EntireRow.Hidden = False
' clear content
Range("B13:P400").Select
Selection.ClearContents
' Make report
Dim i As Integer
Dim j As Integer
j = 13
Workbooks("CMS.xlsx").Sheets("CMS").Activate
For i = 1 To 1000
If Workbooks("CMS.xlsx").Sheets("CMS").Cells(i, 14) = Workbooks("Report.xlsx").Sheets("Visual").Range("Q1") Then
Workbooks("Report.xlsx").Sheets("Visual").Cells(j, 2).Value = Workbooks("CMS.xlsx").Sheets("CMS").Cells(i, 2)
Workbooks("Report.xlsx").Sheets("Visual").Cells(j, 3).Value = Workbooks("CMS.xlsx").Sheets("CMS").Cells(i, 3)
Workbooks("Report.xlsx").Sheets("Visual").Cells(j, 4).Value = Workbooks("CMS.xlsx").Sheets("CMS").Cells(i, 4)
Workbooks("Report.xlsx").Sheets("Visual").Cells(j, 5).Value = Workbooks("CMS.xlsx").Sheets("CMS").Cells(i, 5)
Workbooks("Report.xlsx").Sheets("Visual").Cells(j, 6).Value = Workbooks("CMS.xlsx").Sheets("CMS").Cells(i, 19)
Workbooks("Report.xlsx").Sheets("Visual").Cells(j, 7).Value = Workbooks("CMS.xlsx").Sheets("CMS").Cells(i, 7)
Workbooks("Report.xlsx").Sheets("Visual").Cells(j, 8).Value = Workbooks("CMS.xlsx").Sheets("CMS").Cells(i, 8)
Workbooks("Report.xlsx").Sheets("Visual").Cells(j, 9).Value = Workbooks("CMS.xlsx").Sheets("CMS").Cells(i, 20)
Workbooks("Report.xlsx").Sheets("Visual").Cells(j, 10).Value = Workbooks("CMS.xlsx").Sheets("CMS").Cells(i, 6)
Workbooks("Report.xlsx").Sheets("Visual").Cells(j, 11).Value = "S"
Workbooks("Report.xlsx").Sheets("Visual").Cells(j, 12).Value = Workbooks("CMS.xlsx").Sheets("CMS").Cells(i, 21)
Workbooks("Report.xlsx").Sheets("Visual").Cells(j, 13).Value = Workbooks("CMS.xlsx").Sheets("CMS").Cells(i, 22)
Workbooks("Report.xlsx").Sheets("Visual").Cells(j, 14).Value = Workbooks("CMS.xlsx").Sheets("CMS").Cells(i, 11)
Workbooks("Report.xlsx").Sheets("Visual").Cells(j, 15).Value = "ACC"
Workbooks("Report.xlsx").Sheets("Visual").Range("Q13").Value = Workbooks("CMS.xlsx").Sheets("CMS").Cells(i, 13)
Workbooks("Report.xlsx").Sheets("Visual").Range("Q7").Value = Workbooks("CMS.xlsx").Sheets("CMS").Cells(i, 18)
Workbooks("Report.xlsx").Sheets("Visual").Range("Q8").Value = Workbooks("CMS.xlsx").Sheets("CMS").Cells(i, 14)
j = j + 1
End If
Next i
' hidden row
Workbooks("Report.xlsx").Sheets("Visual").Activate
Dim h As Integer
For h = 13 To 400
If Cells(h, 2).Value = "" Then
Cells(h, 2).EntireRow.Hidden = True
End If
Next h
Application.ScreenUpdating = True
End Sub

Sub Paint()
Application.ScreenUpdating = False
Sheets("Paint").Activate
' inside
Range("E31").Value = Range("N31")
Range("E31").Characters(Start:=5, Length:=1).Font.Superscript = True
Range("I31").Value = Range("P31")
Range("I31").Characters(Start:=5, Length:=1).Font.Superscript = True
Range("E32").Value = Range("N32")
Range("E32").Characters(Start:=5, Length:=1).Font.Superscript = True
Range("I32").Value = Range("P32")
Range("I32").Characters(Start:=5, Length:=1).Font.Superscript = True
' outside
Range("E61").Value = Range("N61")
Range("E61").Characters(Start:=5, Length:=1).Font.Superscript = True
Range("E62").Value = Range("N62")
Range("E62").Characters(Start:=5, Length:=1).Font.Superscript = True
Range("G61").Value = Range("O61")
Range("G61").Characters(Start:=5, Length:=1).Font.Superscript = True
Range("G62").Value = Range("O62")
Range("G62").Characters(Start:=5, Length:=1).Font.Superscript = True
Range("J61").Value = Range("P61")
Range("J61").Characters(Start:=5, Length:=1).Font.Superscript = True
Range("J62").Value = Range("P62")
Range("J62").Characters(Start:=5, Length:=1).Font.Superscript = True
Application.ScreenUpdating = True
End Sub

Sub RTrequest()
' unhinde Sheet
Application.ScreenUpdating = False
Sheets("RT-").Activate
Columns("A:A").Select
Selection.EntireRow.Hidden = False
' clear content
Range("B8:P54").Select
Selection.ClearContents
' Make report
Dim i As Integer
Dim j As Integer
j = 8
Sheets("NDT").Activate
For i = 1 To 1000
If Sheets("NDT").Cells(i, 18) = Sheets("RT-").Range("BG1") Then
Sheets("RT-").Cells(j, 2).Value = Sheets("NDT").Cells(i, 2)
Sheets("RT-").Cells(j, 3).Value = Sheets("NDT").Cells(i, 4)
Sheets("RT-").Cells(j, 4).Value = Sheets("NDT").Cells(i, 3)
Sheets("RT-").Cells(j, 5).Value = Sheets("NDT").Cells(i, 5)
Sheets("RT-").Cells(j, 6).Value = Sheets("NDT").Cells(i, 6)
Sheets("RT-").Cells(j, 7).Value = Sheets("NDT").Cells(i, 7)
Sheets("RT-").Cells(j, 8).Value = Sheets("NDT").Cells(i, 8)
Sheets("RT-").Cells(j, 9).Value = Sheets("NDT").Cells(i, 9)
Sheets("RT-").Cells(j, 10).Value = Sheets("NDT").Cells(i, 10)
Sheets("RT-").Cells(j, 11).Value = Sheets("NDT").Cells(i, 11)
Sheets("RT-").Cells(j, 12).Value = Sheets("NDT").Cells(i, 12)
Sheets("RT-").Cells(j, 13).Value = Sheets("NDT").Cells(i, 13)
Sheets("RT-").Cells(j, 14).Value = Sheets("NDT").Cells(i, 14)
Sheets("RT-").Cells(j, 15).Value = Sheets("NDT").Cells(i, 15)
Sheets("RT-").Range("BD1").Value = Sheets("NDT").Cells(i, 18)
j = j + 1
End If
Next i
' hidden row
Sheets("RT-").Activate
Dim h As Integer
For h = 8 To 66
If Cells(h, 2).Value = "" Then
Cells(h, 2).EntireRow.Hidden = True
End If
Next h
Range("BG1").Select
Application.ScreenUpdating = True
End Sub

Sub CQ()
Application.ScreenUpdating = False
' Make report
Dim i As Integer
Sheets("CQ").Activate
For i = 1 To 1000
If Workbooks("CMS.xlsx").Sheets("Spool list").Cells(i, 17) = Workbooks("Report.xlsx").Sheets("CQ").Range("I1") Then
Range("C10").Value = Workbooks("CMS.xlsx").Sheets("Spool list").Cells(i, 2)
Range("C11").Value = Workbooks("CMS.xlsx").Sheets("Spool list").Cells(i, 3)
Range("D14").Value = Workbooks("CMS.xlsx").Sheets("Spool list").Cells(i, 20)
Range("I15").Value = Workbooks("CMS.xlsx").Sheets("Spool list").Cells(i, 19)
Range("I16").Value = Workbooks("CMS.xlsx").Sheets("Spool list").Cells(i, 5)
Range("I17").Value = Workbooks("CMS.xlsx").Sheets("Spool list").Cells(i, 18)
Range("F10").Value = Workbooks("CMS.xlsx").Sheets("Spool list").Cells(i, 21)
Range("I9").Value = Workbooks("CMS.xlsx").Sheets("Spool list").Cells(i, 17)
End If
Next i
Application.ScreenUpdating = True
End Sub

Sub CQHDPE()
Application.ScreenUpdating = False
' Make report
Dim i As Integer
Sheets("CQ-PE").Activate
For i = 1 To 1000
If Workbooks("CMS.xlsx").Sheets("Spool list").Cells(i, 17) = Workbooks("Report.xlsx").Sheets("CQ-PE").Range("I1") Then
Range("C10").Value = Workbooks("CMS.xlsx").Sheets("Spool list").Cells(i, 2)
Range("C11").Value = Workbooks("CMS.xlsx").Sheets("Spool list").Cells(i, 3)
Range("F10").Value = Workbooks("CMS.xlsx").Sheets("Spool list").Cells(i, 21)
Range("D18").Value = Workbooks("CMS.xlsx").Sheets("Spool list").Cells(i, 19)
Range("I19").Value = Workbooks("CMS.xlsx").Sheets("Spool list").Cells(i, 5)
Range("D20").Value = Workbooks("CMS.xlsx").Sheets("Spool list").Cells(i, 18)
Range("D21").Value = Workbooks("CMS.xlsx").Sheets("Spool list").Cells(i, 15)
Range("I9").Value = Workbooks("CMS.xlsx").Sheets("Spool list").Cells(i, 17)
End If
Next i
Application.ScreenUpdating = True
End Sub

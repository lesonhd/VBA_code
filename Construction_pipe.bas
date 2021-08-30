Attribute VB_Name = "Construction_pipe"
Sub RFI()
Application.ScreenUpdating = False
' Make report
Dim i As Integer
Dim j As Integer
j = 2
Workbooks("DATA.xlsx").Sheets("DATA").Activate
For i = 1 To 1000
If Workbooks("DATA.xlsx").Sheets("DATA").Cells(i, 33) = Workbooks("RFI.xlsx").Sheets("RFI").Range("T2") Then
Workbooks("RFI.xlsx").Sheets("Base").Range("D2").Value = Workbooks("DATA.xlsx").Sheets("DATA").Cells(i, 5)
Workbooks("RFI.xlsx").Sheets("Base").Range("E2").Value = Workbooks("DATA.xlsx").Sheets("DATA").Cells(i, 2)
End If
Next i
For i = 1 To 33
If Workbooks("DATA.xlsx").Sheets("DATA").Cells(2, i) = Workbooks("RFI.xlsx").Sheets("RFI").Range("M6") Then
For j = 1 To 1000
If Workbooks("DATA.xlsx").Sheets("DATA").Cells(j, i) = Workbooks("RFI.xlsx").Sheets("RFI").Range("Q1") Then
Workbooks("RFI.xlsx").Sheets("RFI").Range("H8").Value = Workbooks("DATA.xlsx").Sheets("DATA").Cells(j, i - 2)
Workbooks("RFI.xlsx").Sheets("RFI").Range("K8").Value = Workbooks("DATA.xlsx").Sheets("DATA").Cells(j, i - 1)
Workbooks("RFI.xlsx").Sheets("4AB").Range("G35").Value = Workbooks("DATA.xlsx").Sheets("DATA").Cells(j, 3)
End If
Next j
End If
Next i
Workbooks("RFI.xlsx").Sheets("RFI").Activate
Workbooks("RFI.xlsx").Sheets("RFI").Range("C14").Value = Workbooks("RFI.xlsx").Sheets("RFI").Range("Q14")
Application.ScreenUpdating = True
End Sub

Sub EXC()
Application.ScreenUpdating = False
'Clear picture
Workbooks("RFI.xlsx").Sheets("exc").Activate
ActiveSheet.Shapes.Range(Array("Picture 12", "Picture 13", "Picture 14", "Picture 16")).Select
    Selection.ShapeRange.Left = Range("O12").Left
    Selection.ShapeRange.Top = Range("O12").Top
' Make report
Dim i As Integer
Workbooks("DATA.xlsx").Sheets("DATA").Activate
For i = 1 To 1000
If Workbooks("DATA.xlsx").Sheets("DATA").Cells(i, 11) = Workbooks("RFI.xlsx").Sheets("exc").Range("N1") Then
Workbooks("RFI.xlsx").Sheets("exc").Range("N8").Value = Workbooks("DATA.xlsx").Sheets("DATA").Cells(i, 11)
Workbooks("RFI.xlsx").Sheets("exc").Range("O8").Value = Workbooks("DATA.xlsx").Sheets("DATA").Cells(i, 2)
Workbooks("RFI.xlsx").Sheets("exc").Range("O9").Value = Workbooks("DATA.xlsx").Sheets("DATA").Cells(i, 7)
Workbooks("RFI.xlsx").Sheets("exc").Range("I9").Value = Workbooks("DATA.xlsx").Sheets("DATA").Cells(i, 9)
End If
Next i
Workbooks("RFI.xlsx").Sheets("exc").Activate
Dim a As Integer
Dim b As Integer
Dim c As Integer
Dim d As Integer
Dim e As Integer
a = 200
b = 400
c = 900
d = 1600
e = 1800
If Workbooks("RFI.xlsx").Sheets("exc").Range("O9").Value >= a And Workbooks("RFI.xlsx").Sheets("exc").Range("O9").Value <= b Then
    ActiveSheet.Shapes.Range(Array("Picture 14")).Select
    Selection.ShapeRange.Left = Range("F12").Left
    Selection.ShapeRange.Top = Range("F12").Top
End If
If Workbooks("RFI.xlsx").Sheets("exc").Range("O9").Value < a Then
    ActiveSheet.Shapes.Range(Array("Picture 16")).Select
    Selection.ShapeRange.Left = Range("F12").Left
    Selection.ShapeRange.Top = Range("F12").Top
End If
If Workbooks("RFI.xlsx").Sheets("exc").Range("O9").Value = c Then
ActiveSheet.Shapes.Range(Array("Picture 13")).Select
    Selection.ShapeRange.Left = Range("F12").Left
    Selection.ShapeRange.Top = Range("F12").Top
End If
If Workbooks("RFI.xlsx").Sheets("exc").Range("O9").Value = d Or Workbooks("RFI.xlsx").Sheets("exc").Range("O9").Value = e Then
ActiveSheet.Shapes.Range(Array("Picture 12")).Select
    Selection.ShapeRange.Left = Range("F12").Left
    Selection.ShapeRange.Top = Range("F12").Top
End If
Application.ScreenUpdating = True
End Sub
Sub BED()
Application.ScreenUpdating = False
'Clear picture
Workbooks("RFI.xlsx").Sheets("bed").Activate
ActiveSheet.Shapes.Range(Array("Picture 12", "Picture 8", "Picture 9", "Picture 11")).Select
    Selection.ShapeRange.Left = Range("O12").Left
    Selection.ShapeRange.Top = Range("O12").Top
' Make report
Dim i As Integer
Workbooks("DATA.xlsx").Sheets("DATA").Activate
For i = 1 To 1000
If Workbooks("DATA.xlsx").Sheets("DATA").Cells(i, 14) = Workbooks("RFI.xlsx").Sheets("bed").Range("N1") Then
Workbooks("RFI.xlsx").Sheets("bed").Range("N8").Value = Workbooks("DATA.xlsx").Sheets("DATA").Cells(i, 14)
Workbooks("RFI.xlsx").Sheets("bed").Range("O8").Value = Workbooks("DATA.xlsx").Sheets("DATA").Cells(i, 2)
Workbooks("RFI.xlsx").Sheets("bed").Range("O9").Value = Workbooks("DATA.xlsx").Sheets("DATA").Cells(i, 7)
Workbooks("RFI.xlsx").Sheets("bed").Range("I9").Value = Workbooks("DATA.xlsx").Sheets("DATA").Cells(i, 12)
End If
Next i
Workbooks("RFI.xlsx").Sheets("bed").Activate
Dim a As Integer
Dim b As Integer
Dim c As Integer
Dim d As Integer
Dim e As Integer
a = 200
b = 400
c = 900
d = 1600
e = 1800
If Workbooks("RFI.xlsx").Sheets("bed").Range("O9").Value >= a And Workbooks("RFI.xlsx").Sheets("bed").Range("O9").Value <= b Then
    ActiveSheet.Shapes.Range(Array("Picture 11")).Select
    Selection.ShapeRange.Left = Range("E12").Left
    Selection.ShapeRange.Top = Range("E12").Top
End If
If Workbooks("RFI.xlsx").Sheets("bed").Range("O9").Value < a Then
    ActiveSheet.Shapes.Range(Array("Picture 12")).Select
    Selection.ShapeRange.Left = Range("E12").Left
    Selection.ShapeRange.Top = Range("E12").Top
End If
If Workbooks("RFI.xlsx").Sheets("bed").Range("O9").Value = c Then
ActiveSheet.Shapes.Range(Array("Picture 9")).Select
    Selection.ShapeRange.Left = Range("E12").Left
    Selection.ShapeRange.Top = Range("E12").Top
End If
If Workbooks("RFI.xlsx").Sheets("bed").Range("O9").Value = d Or Workbooks("RFI.xlsx").Sheets("bed").Range("O9").Value = e Then
ActiveSheet.Shapes.Range(Array("Picture 8")).Select
    Selection.ShapeRange.Left = Range("E12").Left
    Selection.ShapeRange.Top = Range("E12").Top
End If
Application.ScreenUpdating = True
End Sub
Sub FINAL()
Application.ScreenUpdating = False
' Make report
Dim i As Integer
Workbooks("DATA.xlsx").Sheets("DATA").Activate
For i = 1 To 1000
If Workbooks("DATA.xlsx").Sheets("DATA").Cells(i, 29) = Workbooks("RFI.xlsx").Sheets("final").Range("N1") Then
Workbooks("RFI.xlsx").Sheets("final").Range("N8").Value = Workbooks("DATA.xlsx").Sheets("DATA").Cells(i, 29)
Workbooks("RFI.xlsx").Sheets("final").Range("O8").Value = Workbooks("DATA.xlsx").Sheets("DATA").Cells(i, 2)
Workbooks("RFI.xlsx").Sheets("final").Range("O9").Value = Workbooks("DATA.xlsx").Sheets("DATA").Cells(i, 7)
Workbooks("RFI.xlsx").Sheets("final").Range("I9").Value = Workbooks("DATA.xlsx").Sheets("DATA").Cells(i, 27)
Workbooks("RFI.xlsx").Sheets("final").Range("O10").Value = Workbooks("DATA.xlsx").Sheets("DATA").Cells(i, 5)
Workbooks("RFI.xlsx").Sheets("final").Range("O11").Value = Workbooks("DATA.xlsx").Sheets("DATA").Cells(i, 3)
End If
Next i
Workbooks("RFI.xlsx").Sheets("final").Activate
Application.ScreenUpdating = True
End Sub
Sub BACKFILL()
Application.ScreenUpdating = False
'Clear picture
Workbooks("RFI.xlsx").Sheets("backfill").Activate
ActiveSheet.Shapes.Range(Array("Picture 12", "Picture 10", "Picture 9", "Picture 11")).Select
    Selection.ShapeRange.Left = Range("O12").Left
    Selection.ShapeRange.Top = Range("O12").Top
' Make report
Dim i As Integer
Workbooks("DATA.xlsx").Sheets("DATA").Activate
For i = 1 To 1000
If Workbooks("DATA.xlsx").Sheets("DATA").Cells(i, 32) = Workbooks("RFI.xlsx").Sheets("backfill").Range("N1") Then
Workbooks("RFI.xlsx").Sheets("backfill").Range("N8").Value = Workbooks("DATA.xlsx").Sheets("DATA").Cells(i, 32)
Workbooks("RFI.xlsx").Sheets("backfill").Range("O8").Value = Workbooks("DATA.xlsx").Sheets("DATA").Cells(i, 2)
Workbooks("RFI.xlsx").Sheets("backfill").Range("O9").Value = Workbooks("DATA.xlsx").Sheets("DATA").Cells(i, 7)
Workbooks("RFI.xlsx").Sheets("backfill").Range("I9").Value = Workbooks("DATA.xlsx").Sheets("DATA").Cells(i, 30)
Workbooks("RFI.xlsx").Sheets("backfill").Range("O10").Value = Workbooks("DATA.xlsx").Sheets("DATA").Cells(i, 5)
Workbooks("RFI.xlsx").Sheets("backfill").Range("O11").Value = Workbooks("DATA.xlsx").Sheets("DATA").Cells(i, 3)
End If
Next i
Workbooks("RFI.xlsx").Sheets("backfill").Activate
Dim a As Integer
Dim b As Integer
Dim c As Integer
Dim d As Integer
Dim e As Integer
a = 200
b = 400
c = 900
d = 1600
e = 1800
If Workbooks("RFI.xlsx").Sheets("backfill").Range("O9").Value >= a And Workbooks("RFI.xlsx").Sheets("backfill").Range("O9").Value <= b Then
    ActiveSheet.Shapes.Range(Array("Picture 11")).Select
    Selection.ShapeRange.Left = Range("C12").Left
    Selection.ShapeRange.Top = Range("C12").Top
End If
If Workbooks("RFI.xlsx").Sheets("backfill").Range("O9").Value < a Then
    ActiveSheet.Shapes.Range(Array("Picture 12")).Select
    Selection.ShapeRange.Left = Range("C12").Left
    Selection.ShapeRange.Top = Range("C12").Top
End If
If Workbooks("RFI.xlsx").Sheets("backfill").Range("O9").Value = c Then
ActiveSheet.Shapes.Range(Array("Picture 10")).Select
    Selection.ShapeRange.Left = Range("C12").Left
    Selection.ShapeRange.Top = Range("C12").Top
End If
If Workbooks("RFI.xlsx").Sheets("backfill").Range("O9").Value = d Or Workbooks("RFI.xlsx").Sheets("backfill").Range("O9").Value = e Then
ActiveSheet.Shapes.Range(Array("Picture 9")).Select
    Selection.ShapeRange.Left = Range("C12").Left
    Selection.ShapeRange.Top = Range("C12").Top
End If
Application.ScreenUpdating = True
End Sub
Sub FITUPC()
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

'unhide
Workbooks("RFI.xlsx").Sheets("Fit-up").Activate
Columns("A:A").Select
Selection.EntireRow.Hidden = False
' clear content
Range("B14:L500").Select
Selection.ClearContents
' Make report
Dim i As Integer
Dim j As Integer
j = 14
Workbooks("DATA.xlsx").Sheets("DATA").Activate
For i = 1 To 4000
If Workbooks("DATA.xlsx").Sheets("DATA").Cells(i, 17) = Workbooks("RFI.xlsx").Sheets("Fit-up").Range("N1") Then
Workbooks("RFI.xlsx").Sheets("Fit-up").Range("N8").Value = Workbooks("DATA.xlsx").Sheets("DATA").Cells(i, 17)
Workbooks("RFI.xlsx").Sheets("Fit-up").Range("O8").Value = Workbooks("DATA.xlsx").Sheets("DATA").Cells(i, 2)
Workbooks("RFI.xlsx").Sheets("Fit-up").Range("O9").Value = Workbooks("DATA.xlsx").Sheets("DATA").Cells(i, 7)
Workbooks("RFI.xlsx").Sheets("Fit-up").Range("H9").Value = Workbooks("DATA.xlsx").Sheets("DATA").Cells(i, 15)
Workbooks("RFI.xlsx").Sheets("Fit-up").Range("O10").Value = Workbooks("DATA.xlsx").Sheets("DATA").Cells(i, 5)
Workbooks("RFI.xlsx").Sheets("Fit-up").Range("O11").Value = Workbooks("DATA.xlsx").Sheets("DATA").Cells(i, 3)
End If
Next i
For i = 1 To 4000
If Workbooks("DATA.xlsx").Sheets("CMS").Cells(i, 13) = Workbooks("RFI.xlsx").Sheets("Fit-up").Range("N1") Then
Workbooks("RFI.xlsx").Sheets("Fit-up").Cells(j, 2).Value = Workbooks("DATA.xlsx").Sheets("CMS").Cells(i, 3)
Workbooks("RFI.xlsx").Sheets("Fit-up").Cells(j, 3).Value = Workbooks("DATA.xlsx").Sheets("CMS").Cells(i, 4)
Workbooks("RFI.xlsx").Sheets("Fit-up").Cells(j, 4).Value = Workbooks("DATA.xlsx").Sheets("CMS").Cells(i, 5)
Workbooks("RFI.xlsx").Sheets("Fit-up").Cells(j, 6).Value = Workbooks("DATA.xlsx").Sheets("CMS").Cells(i, 7)
Workbooks("RFI.xlsx").Sheets("Fit-up").Cells(j, 7).Value = Workbooks("DATA.xlsx").Sheets("CMS").Cells(i, 6)
Workbooks("RFI.xlsx").Sheets("Fit-up").Cells(j, 8).Value = Workbooks("DATA.xlsx").Sheets("CMS").Cells(i, 10)
Workbooks("RFI.xlsx").Sheets("Fit-up").Cells(j, 9).Value = Workbooks("DATA.xlsx").Sheets("CMS").Cells(i, 11)
Workbooks("RFI.xlsx").Sheets("Fit-up").Cells(j, 10).Value = "3~4"
Workbooks("RFI.xlsx").Sheets("Fit-up").Cells(j, 11).Value = "ACC"
j = j + 1
End If
Next i
' hidden row
Workbooks("RFI.xlsx").Sheets("Fit-up").Activate
Dim h As Integer
For h = 14 To 499
If Cells(h, 2).Value = "" Then
Cells(h, 2).EntireRow.Hidden = True
End If
Next h
Workbooks("RFI.xlsx").Sheets("Fit-up").Activate
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
End Sub

Sub VISUALC()
Application.ScreenUpdating = False
'unhide
Workbooks("RFI.xlsx").Sheets("Visual").Activate
Columns("A:A").Select
Selection.EntireRow.Hidden = False
' clear content
Range("B13:P515").Select
Selection.ClearContents
' Make report
Dim i As Integer
Dim j As Integer
j = 13
Workbooks("DATA.xlsx").Sheets("DATA").Activate
For i = 1 To 4000
If Workbooks("DATA.xlsx").Sheets("DATA").Cells(i, 20) = Workbooks("RFI.xlsx").Sheets("Visual").Range("R1") Then
Workbooks("RFI.xlsx").Sheets("Visual").Range("Q8").Value = Workbooks("DATA.xlsx").Sheets("DATA").Cells(i, 20)
Workbooks("RFI.xlsx").Sheets("Visual").Range("S8").Value = Workbooks("DATA.xlsx").Sheets("DATA").Cells(i, 2)
Workbooks("RFI.xlsx").Sheets("Visual").Range("K9").Value = Workbooks("DATA.xlsx").Sheets("DATA").Cells(i, 18)
End If
Next i
For i = 1 To 4000
If Workbooks("DATA.xlsx").Sheets("CMS").Cells(i, 17) = Workbooks("RFI.xlsx").Sheets("Visual").Range("R1") Then
Workbooks("RFI.xlsx").Sheets("Visual").Cells(j, 2).Value = Workbooks("DATA.xlsx").Sheets("CMS").Cells(i, 3)
Workbooks("RFI.xlsx").Sheets("Visual").Cells(j, 3).Value = Workbooks("DATA.xlsx").Sheets("CMS").Cells(i, 4)
Workbooks("RFI.xlsx").Sheets("Visual").Cells(j, 4).Value = Workbooks("DATA.xlsx").Sheets("CMS").Cells(i, 5)
Workbooks("RFI.xlsx").Sheets("Visual").Cells(j, 5).Value = Workbooks("DATA.xlsx").Sheets("CMS").Cells(i, 7)
Workbooks("RFI.xlsx").Sheets("Visual").Cells(j, 6).Value = Workbooks("DATA.xlsx").Sheets("CMS").Cells(i, 6)
Workbooks("RFI.xlsx").Sheets("Visual").Cells(j, 7).Value = Workbooks("DATA.xlsx").Sheets("CMS").Cells(i, 10)
Workbooks("RFI.xlsx").Sheets("Visual").Cells(j, 8).Value = Workbooks("DATA.xlsx").Sheets("CMS").Cells(i, 11)
Workbooks("RFI.xlsx").Sheets("Visual").Cells(j, 9).Value = Workbooks("DATA.xlsx").Sheets("CMS").Cells(i, 20)
Workbooks("RFI.xlsx").Sheets("Visual").Cells(j, 10).Value = Workbooks("DATA.xlsx").Sheets("CMS").Cells(i, 8)
Workbooks("RFI.xlsx").Sheets("Visual").Cells(j, 11).Value = "F"
Workbooks("RFI.xlsx").Sheets("Visual").Cells(j, 12).Value = Workbooks("DATA.xlsx").Sheets("CMS").Cells(i, 15)
Workbooks("RFI.xlsx").Sheets("Visual").Cells(j, 13).Value = "N/A"
Workbooks("RFI.xlsx").Sheets("Visual").Cells(j, 14).Value = Workbooks("DATA.xlsx").Sheets("CMS").Cells(i, 14)
Workbooks("RFI.xlsx").Sheets("Visual").Cells(j, 15).Value = "ACC"
j = j + 1
End If
Next i
' hidden row
Workbooks("RFI.xlsx").Sheets("Visual").Activate
Dim h As Integer
For h = 13 To 514
If Cells(h, 2).Value = "" Then
Cells(h, 2).EntireRow.Hidden = True
End If
Next h
Workbooks("RFI.xlsx").Sheets("Visual").Activate

Application.ScreenUpdating = True
End Sub

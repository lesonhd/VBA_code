Attribute VB_Name = "Construction_equipment"
Sub RFIequipment()
Application.ScreenUpdating = False
' Make report
Dim i As Integer
Dim j As Integer
j = 2
Workbooks("DATA.xlsx").Sheets("E").Activate
For i = 1 To 1000
If Workbooks("DATA.xlsx").Sheets("E").Cells(i, 17) = Workbooks("RFI-E.xlsx").Sheets("RFI").Range("T2") Then
Workbooks("RFI-E.xlsx").Sheets("Base").Range("C1").Value = Workbooks("DATA.xlsx").Sheets("E").Cells(i, 3)
Workbooks("RFI-E.xlsx").Sheets("Base").Range("F1").Value = Workbooks("DATA.xlsx").Sheets("E").Cells(i, 2)
Workbooks("RFI-E.xlsx").Sheets("Base").Range("G1").Value = Workbooks("DATA.xlsx").Sheets("E").Cells(i, 1)
End If
Next i
For i = 1 To 17
If Workbooks("DATA.xlsx").Sheets("E").Cells(2, i) = Workbooks("RFI-E.xlsx").Sheets("RFI").Range("M6") Then
For j = 1 To 1000
If Workbooks("DATA.xlsx").Sheets("E").Cells(j, i) = Workbooks("RFI-E.xlsx").Sheets("RFI").Range("Q1") Then
Workbooks("RFI-E.xlsx").Sheets("RFI").Range("H8").Value = Workbooks("DATA.xlsx").Sheets("E").Cells(j, i - 2)
Workbooks("RFI-E.xlsx").Sheets("RFI").Range("K8").Value = Workbooks("DATA.xlsx").Sheets("E").Cells(j, i - 1)
Workbooks("RFI-E.xlsx").Sheets("4AB").Range("G35").Value = Workbooks("DATA.xlsx").Sheets("E").Cells(j, 19)
End If
Next j
End If
Next i
Application.ScreenUpdating = True
Workbooks("RFI-E.xlsx").Sheets("RFI").Activate
End Sub

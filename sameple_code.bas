Attribute VB_Name = "sameple_code"
Option Explicit

Sub copy_paste()
' copy mot day cung lam tuong tu
Range("A1").Copy Range("B1")
End Sub

Sub move_range()
Range("A1").Cut Range("B1")
End Sub
' Tim dong cuoi tu tren xuong
Sub tim_dong_cuoi()
Dim t As Long
' giong nhu chon o A1 sau do an Ctrl+ di xuong
t = Range("A1").End(xlDown).Row
MsgBox t
End Sub
'tim dong cuoi tu duoi len
Sub tim_dong_cuoi2()
    Dim t As Long
' Column("A").rows.count se tra ve Tong so dong trong cot A
    t = Range("A" & Columns("A").Rows.Count).End(xlUp).Row
    MsgBox t
End Sub

Sub xoa_item_trong_collection()
Dim kts As New Collection
Dim j As Variant

kts.Add Item:=800
kts.Add Item:=900
kts.Add Item:=700
 For j = 1 To kts.Count
    kts.Remove (1)
 Next
MsgBox kts.Count
End Sub

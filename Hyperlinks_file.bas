Attribute VB_Name = "Hyperlinks_file"
Option Explicit

Sub link_file()
'Anchor dien vi tri muon dinh kem file
' Address : duong dan den file dinh k�m
' ScreenTip : hien dong nhac khi dua chuot v�o
' TexttoDisplay : ten cua chu khi hien thi
            ThisWorkbook.Sheets(1).Hyperlinks.Add anchor:=ThisWorkbook.Sheets(1).Range("A1"), _
            Address:="dien duong d�n d�n file dinh k�m v�o d�y", _
            ScreenTip:=ten_file, _
            TextToDisplay:=ten_file
End Sub


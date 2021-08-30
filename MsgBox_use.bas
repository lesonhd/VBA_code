Attribute VB_Name = "MsgBox_use"
Option Explicit

Sub use_msg()
Dim t As Variant
t = MsgBox(Prompt:=" Chon Yes or No ", Buttons:=vbYesNo)
' khi chon yes t se bang 6, chon No thi t =7. Co the su dung gia tri nay de lam cac viec tiep theo
MsgBox t
End Sub

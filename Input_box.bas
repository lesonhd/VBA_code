Attribute VB_Name = "Input_box"
Option Explicit

Sub input_box()
  Dim myrange As Range
  Dim rg As Range
  Set myrange = Application.InputBox(Prompt:="chon vung du lieu", Type:=8)
    For Each rg In myrange
        rg.Value = 123
    Next
End Sub


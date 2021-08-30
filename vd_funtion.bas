Attribute VB_Name = "vd_funtion"
Option Explicit
' phai copy function vao 1 module cua sheet thi moi chay duoc
Function findArea(Height As Double, Optional Width As Variant) As Double
   If IsMissing(Width) Then
      findArea = Height * Height
   Else
      findArea = Height * Width
   End If
End Function

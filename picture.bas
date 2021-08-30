Attribute VB_Name = "picture"
Option Explicit
' Chen anh vao o A1 cua sheet 1 va di chuyen
Sub vd_2()

    Sheets(1).Range("A1").Select
    ActiveSheet.Pictures.Insert("E:\Company's document\Welder Data\2-welder photo\Binh-1.jpg").Select
    ' di chuyen anh vua chen vao
    Selection.ShapeRange.IncrementLeft 20
    Selection.ShapeRange.IncrementTop 50
End Sub

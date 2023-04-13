Attribute VB_Name = "Module1"

Sub Zadanie_1()
Const C1 As Single = 0.5, C2 As Single = 1.3, C3 As Single = 10
Dim a As Single, b As Single, c As Single
Dim m As Single, d As Single
a = InputBox("")
b = InputBox("")
c = InputBox("")

Cells(1, 1).Value = "Исходные данные:"

Cells(2, 1).Value = "a="
Cells(2, 2).Value = a

Cells(3, 1).Value = "b="
Cells(3, 2).Value = b

Cells(4, 1).Value = "c="
Cells(4, 2).Value = c

If (a = 0 Or b = 0 Or c = 0) Then
Cells(6, 2) = "Ошибка : a = 0 или b = 0 или c = 0! Программа будет завершена"


Else
m = ((a) / (b * c)) ^ 2 + Sqr(Abs((a - b) / (c ^ 2 + 2 * a - 4 * b)))
Cells(6, 1).Value = "Результаты:"
Cells(7, 1).Value = "M="
Cells(7, 2).Value = m

d = Sin(m) + Cos(m ^ 2)
Cells(8, 1).Value = "D="
Cells(8, 2).Value = d
End If
End Sub

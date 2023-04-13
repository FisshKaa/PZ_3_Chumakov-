Attribute VB_Name = "Module2"

Sub Zadanie_2()
Const C4 As Single = 0.5, C5 As Single = 1.3, C6 As Single = 10
Dim k As Single, m As Single, c As Single
Dim x_1 As Single, x_2 As Single
k = InputBox("")
m = InputBox("")
c = InputBox("")

Cells(1, 4).Value = "Исходные данные:"

Cells(2, 4).Value = "k="
Cells(2, 5).Value = k

Cells(3, 4).Value = "m="
Cells(3, 5).Value = m

Cells(4, 4).Value = "c="
Cells(4, 5).Value = c

If (k = 0 Or m = 0 Or c = 0) Then
Cells(6, 5) = "Ошибка : k = 0 или m = 0 или c = 0! Программа будет завершена"
Else
x_1 = m ^ 2 - k ^ 2 - 4 * m * c
Cells(6, 4).Value = "Результаты:"
Cells(7, 4).Value = "x1="
Cells(7, 5).Value = x_1
End If


If x_1 >= 0 Then
x_1 = Sqr(x_1)
x_2 = x_1 ^ 2
Cells(8, 4).Value = "x2="
Cells(8, 5).Value = x_2
Else:
End If


If x_1 < 0 Then
x_1 = Abs(x_1)
x_2 = x_1 ^ 2
Cells(8, 4).Value = "x2="
Cells(8, 5).Value = x_2
Else:
End If


End Sub

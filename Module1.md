Private Sub CommandButton1_Click()
'заполняет элементы массива рандомными положительными значениями'
For i = 1 To 30
Cells(1, i) = Int((100 * Rnd) + 1)
Next i
End Sub

Private Sub CommandButton2_Click()
'находит наименьший номер элемента массива, число в котором эквивалентно заданному пользователем с клавиатуры числу,
'или сообщает, что эквивалентных чисел в элементах массива нет'
x = TextBox1.Text
For i = 30 To 1 Step -1
If Int(x) = Cells(1, i) Then
b = Str(i)
End If
Next i
If b = Empty Then
b = "There is no such element"
End If
MsgBox (b)
End Sub

Private Sub CommandButton3_Click()
'закрывает форму'
UserForm1.Hide
End Sub
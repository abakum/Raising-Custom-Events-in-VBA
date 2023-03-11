Attribute VB_Name = "сердце"
Private классы As New Collection
Public Sub жди(Optional класс = Nothing, Optional таймер = 0)
 Dim i As Integer
 If класс Is Nothing Then
  If таймер < 0 Then Set классы = New Collection 'жди , -1 - деструкция всех классов в коллекции
 Else
  If таймер = 0 Then
   классы.Add класс 'жди Me - конструктор для асинхронных событий
   пора классы.Count
  ElseIf таймер > 0 Then
   пора таймер 'когда несколько классов вызывают асинхронных событий с одной частотой
  Else
   For i = классы.Count To 1 Step -1
    If VarType(классы(i)) = vbObject Then
     If класс Is классы(i) Then
      классы.Add Nothing, after:=i
      классы.Remove i 'жди Me, -1 - останов асинхронных событий одного класса
      Exit Sub
     End If
    End If
   Next
  End If
 End If
End Sub
Public Sub пора(i)
 On Error Resume Next
 CallByName классы(i), "пора", VbMethod, i
End Sub
Public Sub пора1(): пора 1: End Sub
Public Sub пора2(): пора 2: End Sub
Public Sub пора3(): пора 3: End Sub
Public Sub пора4(): пора 4: End Sub
Public Sub пора5(): пора 5: End Sub
Public Sub пора6(): пора 6: End Sub
Public Sub пора7(): пора 7: End Sub
Public Sub пора8(): пора 8: End Sub
Public Sub пора9(): пора 9: End Sub
Public Sub пора10(): пора 10: End Sub
Public Sub пора11(): пора 11: End Sub
Public Sub пора12(): пора 12: End Sub
Public Sub пора13(): пора 13: End Sub
Public Sub пора14(): пора 14: End Sub
Public Sub пора15(): пора 15: End Sub
Public Sub пора16(): пора 16: End Sub
Public Sub пора17(): пора 17: End Sub
Public Sub пора18(): пора 18: End Sub
Public Sub пора19(): пора 19: End Sub
Public Sub пора20(): пора 20: End Sub



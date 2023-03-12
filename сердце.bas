Attribute VB_Name = "сердце"
Private классы As New Collection
Public Sub жди(Optional класс, Optional конструктор = True)
 Dim и As Integer
 If IsMissing(класс) Then
  Set классы = New Collection 'жди - деструкция всех классов в коллекции 'классы'
 Else
  If конструктор Then
   классы.Add класс 'жди Me - вызывать из конструктора для асинхронных событий
   пора классы.Count
  Else
   For и = классы.Count To 1 Step -1
    If VarType(классы(и)) = vbObject Then
     If класс Is классы(и) Then
      классы.Add Nothing, after:=i
      классы.Remove и 'жди Me, False - вызывать для останова асинхронных событий
      Exit Sub
     End If
    End If
   Next и
  End If
 End If
End Sub
Public Sub пора(целое)
 On Error Resume Next
 CallByName классы(целое), "пора", VbMethod, целое
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



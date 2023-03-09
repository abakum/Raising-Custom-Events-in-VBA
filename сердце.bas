Attribute VB_Name = "сердце"
Private живые As New Collection
Public Sub жди(Optional объект_ = Nothing, Optional повтор = True)
 Dim i As Integer
 If Not объект_ Is Nothing Then
  'жди Me
  If повтор Then
   живые.Add объект_
  Else 'жди Me, False
   For i = живые.Count To 1 Step -1
    If IsEmpty(живые(i)) Then
     живые.Remove i
    Else
     If объект_ Is живые(i) Then
      CallByName объект_, "стоп", VbMethod
      живые.Remove i
      Exit Sub
     End If
    End If
   Next
  End If
 End If
 For i = живые.Count To 1 Step -1
  If IsEmpty(живые(i)) Then
   живые.Remove i
  Else
   If живые(i) Is Nothing Then
    живые.Remove i
   Else
    If повтор Then
     CallByName живые(i), "пора", VbMethod
    Else
     живые.Remove i
    End If
   End If
  End If
 Next i
End Sub

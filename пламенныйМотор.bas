Attribute VB_Name = "пламенныйМотор"
Const ЧСС = 60
Public сердце As часы

'методы
'не блокирующий цикл повторных вызовов кардиоСтимулятор
Public Sub кардиоСтимулятор(Optional repeat = True)
 Static систола As Date
 On Error Resume Next
 Application.OnTime систола, "кардиоСтимулятор", , False
 If repeat Then
  If сердце Is Nothing Then Set сердце = New часы
 Else
  Set сердце = Nothing
  Exit Sub
 End If
 сердце.стук
 систола = Now + TimeSerial(0, 0, 60 / ЧСС)
 Application.OnTime систола, "кардиоСтимулятор"
End Sub

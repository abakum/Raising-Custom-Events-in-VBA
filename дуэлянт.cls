VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "дуэлянт"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'класс Слушатель обрабатывает События объектов через WithEvents
Public WithEvents револьвер As револьверСоднимПатроном

'класс Вещатель который описыват События Event и вызывает RaiseEvent в теле methods/functions/properties
Public Event пульс()

Private имя_
Private позже
Private какПозже
Private ждун
Private конец
Private дулоУвиска

'обработчики событий
Private Sub Class_Initialize()
 ЧСС = 60
 жди Me
End Sub
Private Sub Class_Terminate()
 'конец = True
End Sub
Private Sub револьвер_ЗвукВращающегосяБарабана()
 If конец Then Exit Sub
 Debug.Print Format(Now, "h:mm:ss"); " "; имя; " услышал звук вращающегося барабана"
End Sub
Private Sub револьвер_ЗвукВращающегосяБарабанаЗатих()
 If конец Then Exit Sub
 Debug.Print Format(Now, "h:mm:ss"); " "; имя; " услышал, что звук вращающегося барабана затих"
End Sub
Private Sub револьвер_ЩелкнулКурком()
 If конец Then Exit Sub
 Debug.Print Format(Now, "h:mm:ss"); " "; имя; " услышал щелчок курка"
End Sub
Private Sub револьвер_Выстрелил()
 If конец Then Exit Sub
 If дулоУвиска Then
  конец = True
  стоп
  жди Me, -1
 Else
  Debug.Print Format(Now, "h:mm:ss"); " "; имя; " услышал выстрел"
 End If
End Sub

'свойства
Public Property Let имя(строка)
 имя_ = строка
End Property
Public Property Get имя()
 имя = имя_
End Property
Public Property Let ЧСС(целое)
 какПозже = TimeSerial(0, 0, 60 / IIf(целое > 119, 119, целое)) '120 вешает эксел
End Property

'методы
Public Sub ПриставилДулоКвиску()
 If конец Then Exit Sub
 Debug.Print
 Debug.Print Format(Now, "h:mm:ss"); " "; имя; " приставил дуло к своему виску"
 дулоУвиска = True
 WaitSec RandomInt(1, 6)
End Sub
Public Function ПередалРевольвер()
 If конец Then Exit Function
 дулоУвиска = False
 WaitSec RandomInt(1, 6)
 Debug.Print Format(Now, "h:mm:ss"); " "; имя; " передал револьвер"
 ПередалРевольвер = True
End Function
Public Sub НажалНаСпусковойКрючок()
 If конец Then Exit Sub
 Debug.Print Format(Now, "h:mm:ss"); " "; имя; " нажал на спусковой крючок"
 револьвер.НажатСпусковойКрючок
End Sub
Public Sub КрутанулБарабан()
 If конец Then Exit Sub
 Debug.Print Format(Now, "h:mm:ss"); " "; имя; " крутанул барабан"
 револьвер.КрутанутБарабан
End Sub
Public Sub пора(целое)
 ждун = целое
 стоп
 If конец Then Exit Sub
 RaiseEvent пульс
 позже = Now + какПозже
 Application.OnTime позже, "пора" & ждун
 'Debug.Print Format(позже, "h:mm:ss"); " "; имя; " OnTime", ждун
End Sub
Private Sub стоп()
 On Error Resume Next
 Application.OnTime позже, "пора" & ждун, , False
 'Debug.Print Format(Now, "h:mm:ss"); " "; имя; " стоп", ждун
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "дуэлянт"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
#Const отладка = 0
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
 'Class_Terminate не сработает пока есть ссылки на дуэлянта а они есть в коллекции классы
 'поэтому  надо запускать 'жди' в main перед деструкцией классов из которых вызывается жди Me
 стоп
End Sub
Private Sub револьвер_ЗвукВращающегосяБарабана()
 Debug.Print Format(Now, "h:mm:ss"); " "; имя; " услышал звук вращающегося барабана"
End Sub
Private Sub револьвер_ЗвукВращающегосяБарабанаЗатих()
 Debug.Print Format(Now, "h:mm:ss"); " "; имя; " услышал, что звук вращающегося барабана затих"
End Sub
Private Sub револьвер_ЩелкнулКурком()
 Debug.Print Format(Now, "h:mm:ss"); " "; имя; " услышал щелчок курка"
End Sub
Private Sub револьвер_Выстрелил()
 If дулоУвиска Then
  стоп
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
 '120 вешает эксел так как параметр EarliestTime округляется до ближайшей целой секунды. То есть 60/120~0
 какПозже = TimeSerial(0, 0, 60 / IIf(целое > 119, 119, целое))
End Property

'методы
Public Sub ПриставилДулоКвиску()
 If конец Then Exit Sub
 Debug.Print
 Debug.Print Format(Now, "h:mm:ss"); " "; имя; " приставил дуло к своему виску"
 дулоУвиска = True
 ждатьСекунд Случайное(1, 6)
End Sub
Public Function ПередалРевольвер()
 If конец Then Exit Function
 дулоУвиска = False
 ждатьСекунд Случайное(1, 6)
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
 If конец Then Exit Sub
 ждун = целое
 RaiseEvent пульс
 позже = Now + какПозже
 Application.OnTime позже, "пора" & ждун
 #If отладка Then
  Debug.Print Format(Now, "h:mm:ss"); " "; имя; " OnTime", ждун, Format(позже, "h:mm:ss")
 #End If
End Sub
Private Sub стоп()
 #If отладка Then
  Debug.Print Format(Now, "h:mm:ss"); " "; имя; " стоп", ждун
 #End If
 конец = True
 Set револьвер = Nothing
 On Error Resume Next
 Application.OnTime позже, "пора" & ждун, , False
 жди Me, False
End Sub


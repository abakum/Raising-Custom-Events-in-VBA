VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "дуэлянт"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
#Const отладка = 1
'класс Слушатель обрабатывает События объектов через WithEvents
Public WithEvents револьвер As револьверСоднимПатроном

'класс Вещатель который описыват События Event и вызывает RaiseEvent в теле methods/functions/properties
Public Event пульс()

Private имя_
Private позже
Private какПозже
Private конец
Private дулоУвиска

'обработчики событий
Private Sub Class_Initialize()
 ЧСС = 60
 жди Me
End Sub
Private Sub Class_Terminate()
 #If отладка Then
  Debug.Print Format(Now, "h:mm:ss"); " "; имя; " Class_Terminate"
 #End If
 конец = True
 Set револьвер = Nothing
 On Error Resume Next
 жди Me, False 'отменяет ожидаемое в 'позже' событие Application.onTime
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
  Class_Terminate
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
Public Property Let пора(целое)
 If конец Then Exit Property
 RaiseEvent пульс
 позже = Now + какПозже
 Application.onTime позже, "пора" & целое 'пораX это заменители 'дуэлянт(X).onTime=X' потому, что Application.onTime требует процедуру в обычном модуле
 #If отладка Then
  Debug.Print Format(Now, "h:mm:ss"); " "; имя; " пора", целое, Format(позже, "h:mm:ss")
 #End If
End Property
Public Property Get пора()
 пора = позже
End Property

'методы
Public Sub ПриставилДулоКвиску()
 If конец Then Exit Sub
 Debug.Print
 Debug.Print Format(Now, "h:mm:ss"); " "; имя; " приставил дуло к своему виску"
 дулоУвиска = True
 ждатьСек cлучайное(1, 6)
End Sub
Public Function ПередалРевольвер()
 If конец Then Exit Function
 дулоУвиска = False
 ждатьСек cлучайное(1, 6)
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

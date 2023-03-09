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
Public Event Пульс()

Private имя_
Private какПозже_
Private позже
Private нетПульса
Private дулоУвиска

'обработчики событий
Private Sub Class_Initialize()
 нетПульса = False
 ЧСС = 60
 жди Me
End Sub
Private Sub Class_Terminate()
 нетПульса = True
 'жди Me, False
End Sub
Private Sub револьвер_ЗвукВращающегосяБарабана()
 If нетПульса Then Exit Sub
 Debug.Print Format(Now, "h:mm:ss"); " "; имя; " услышал звук вращающегося барабана"
End Sub
Private Sub револьвер_ЗвукВращающегосяБарабанаЗатих()
 If нетПульса Then Exit Sub
 Debug.Print Format(Now, "h:mm:ss"); " "; имя; " услышал, что звук вращающегося барабана затих"
End Sub
Private Sub револьвер_ЩелкнулКурком()
 If нетПульса Then Exit Sub
 Debug.Print Format(Now, "h:mm:ss"); " "; имя; " услышал щелчок курка"
End Sub
Private Sub револьвер_Выстрелил()
 If дулоУвиска Then
  планировщик Me, False
  нетПульса = True
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
 какПозже_ = TimeSerial(0, 0, 60 / целое)
End Property

'методы
Public Sub ПриставилДулоКвиску()
 If нетПульса Then Exit Sub
 Debug.Print
 Debug.Print Format(Now, "h:mm:ss"); " "; имя; " приставил дуло к своему виску"
 дулоУвиска = True
 WaitSec RandomInt(1, 6)
End Sub
Public Function ПередалРевольвер()
 If нетПульса Then Exit Function
 дулоУвиска = False
 WaitSec RandomInt(1, 6)
 Debug.Print Format(Now, "h:mm:ss"); " "; имя; " передал револьвер"
 ПередалРевольвер = True
End Function
Public Sub НажалНаСпусковойКрючок()
 If нетПульса Then Exit Sub
 Debug.Print Format(Now, "h:mm:ss"); " "; имя; " нажал на спусковой крючок"
 револьвер.НажатСпусковойКрючок
End Sub
Public Sub КрутанулБарабан()
 If нетПульса Then Exit Sub
 Debug.Print Format(Now, "h:mm:ss"); " "; имя; " крутанул барабан"
 револьвер.КрутанутБарабан
End Sub
Public Sub пора()
 стоп
 If нетПульса Then Exit Sub
 RaiseEvent Пульс
 позже = Now + какПозже_
 Application.OnTime позже, "жди"
End Sub
Public Sub стоп()
 On Error Resume Next
 Application.OnTime позже, "жди", , False
End Sub

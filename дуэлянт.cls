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
Public WithEvents мотор As часы

'класс Вещатель который описыват События Event и вызывает RaiseEvent в теле methods/functions/properties
Public Event Пульс()

Private нетПульса
Private дулоУвиска
Private имя_

'обработчики событий
Private Sub Class_Initialize()
 нетПульса = False
 Set мотор = сердце
End Sub
Private Sub Class_Terminate()
 нетПульса = True
 Set мотор = Nothing
End Sub
Private Sub мотор_такт()
 If нетПульса Then Exit Sub
 RaiseEvent Пульс
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
  нетПульса = True
 Else
  Debug.Print Format(Now, "h:mm:ss"); " "; имя; " услышал выстрел"
 End If
End Sub

'свойства
Public Property Let имя(s)
 имя_ = s
End Property
Public Property Get имя()
 имя = имя_
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

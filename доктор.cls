VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "доктор"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'класс Слушатель обрабатывает События объектов через WithEvents
Public WithEvents револьвер As револьверСоднимПатроном
Attribute револьвер.VB_VarHelpID = -1
Public WithEvents пациент As дуэлянт

Private пациентСкорееЖивЧемМёртв
Private человек_ As дуэлянт 'пока нет выстрела нет и пациента

'обработчики событий
Private Sub пациент_пульс()
 пациентСкорееЖивЧемМёртв = True
End Sub
Private Sub револьвер_Выстрелил()
 СлушаетПульс
End Sub

'методы
Public Sub СлушаетПульс()
 пациентСкорееЖивЧемМёртв = False
 Set пациент = человек_ 'слушаем событие пациент_пульс
 Debug.Print Format(Now, "h:mm:ss"); " Доктор слушает пульс у "; пациент.имя
 WaitSec 3
 If пациентСкорееЖивЧемМёртв Then
  Debug.Print Format(Now, "h:mm:ss"); " Доктор сказал, что "; пациент.имя; " ещё жив"
 Else
  Debug.Print Format(Now, "h:mm:ss"); " Доктор сказал, что "; пациент.имя; " убит"
 End If
 Set пациент = Nothing ' перестаём слушать событие пациент_пульс
End Sub
Public Sub ВидитЧтоРевольверВзял(человек As дуэлянт)
 Set человек_ = человек
 Set револьвер = человек.револьвер 'слушаем событие револьвер_Выстрелил
End Sub

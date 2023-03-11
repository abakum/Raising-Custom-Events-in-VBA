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
Public WithEvents пациент As дуэлянт

Private человек_ As дуэлянт 'пока нет выстрела нет и пациента
Private пульс

'обработчики событий
Private Sub пациент_пульс()
 пульс = пульс + 1
End Sub
Private Sub револьвер_Выстрелил()
 СчитаетПульс
End Sub

'методы
Public Sub СчитаетПульс()
 пульс = 0
 Set пациент = человек_ 'слушаем событие пациент_пульс
 Debug.Print Format(Now, "h:mm:ss"); " Доктор слушает пульс у "; пациент.имя
 WaitSec 6
 If пульс > 0 Then
  Debug.Print Format(Now, "h:mm:ss"); " Доктор сказал, что у "; пациент.имя; " пульс "; пульс * 10
 Else
  Debug.Print Format(Now, "h:mm:ss"); " Доктор сказал, что у "; пациент.имя; " нет пульса"
 End If
 Set пациент = Nothing ' перестаём слушать событие пациент_пульс
End Sub
Public Sub ВидитЧтоРевольверВзял(человек As дуэлянт)
 Set человек_ = человек
 Set револьвер = человек.револьвер 'слушаем событие револьвер_Выстрелил
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "револьверСоднимПатроном"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'класс Вещатель который описыват События Event и вызывает RaiseEvent в теле methods/functions/properties
Public Event Выстрелил()
Public Event ЩелкнулКурком()
Public Event ЗвукВращающегосяБарабана()
Public Event ЗвукВращающегосяБарабанаЗатих()

Private сколькоЗарядный_
Private камораСпатроном
Private камораПодКурком

'свойства
Public Property Let сколькоЗарядный(байт)
 сколькоЗарядный_ = байт
End Property
Public Property Get сколькоЗарядный()
 сколькоЗарядный = сколькоЗарядный_
End Property

'методы
Public Sub КрутанутБарабан()
 RaiseEvent ЗвукВращающегосяБарабана
 Randomize
 камораПодКурком = 1
 WaitSec RandomInt(1, 6)
 камораСпатроном = RandomInt(0, сколькоЗарядный_)
 'камораСпатроном = 0 'осечка
 RaiseEvent ЗвукВращающегосяБарабанаЗатих
End Sub
Public Sub НажатСпусковойКрючок()
 If камораПодКурком = сколькоЗарядный_ Then
  камораПодКурком = 1
 Else
  камораПодКурком = камораПодКурком + 1
 End If
 If камораПодКурком = камораСпатроном Then
  камораСпатроном = 0
  RaiseEvent Выстрелил
 Else
  RaiseEvent ЩелкнулКурком
 End If
End Sub

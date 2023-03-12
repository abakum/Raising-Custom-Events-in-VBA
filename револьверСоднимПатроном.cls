VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "револьверСоднимПатроном"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
#Const отладка = 0
#Const осечка = 0
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
 ждатьСекунд Случайное(1, 6)
 камораСпатроном = Случайное(0, сколькоЗарядный_)
 #If отладка Then
  #If осечка Then
   камораСпатроном = 0 'осечка
  #Else
   камораСпатроном = 2 'выстрел с первого нажатия крючка
  #End If
 #End If
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

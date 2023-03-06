VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "дуэлянт"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'a Listener class that handles events via WithEvents

Option Explicit

Public WithEvents револьвер As револьверСоднимПатроном
Attribute револьвер.VB_VarHelpID = -1
Private убит_
Private имя_
Private дулоУвиска

Private Sub Class_Initialize()
 убит_ = False
End Sub

Private Sub револьвер_ЗвукВращающегосяБарабана()
 If убит_ Then Exit Sub
 Debug.Print имя; " услышал звук вращающегося барабана в "; Format(Now, "h:mm:ss")
End Sub
Private Sub револьвер_ЗвукВращающегосяБарабанаЗатих()
 If убит_ Then Exit Sub
 Debug.Print имя; " услышал, что звук вращающегося барабана затих в "; Format(Now, "h:mm:ss")
End Sub
Private Sub револьвер_ЩелкнулКурком()
 If убит_ Then Exit Sub
 Debug.Print имя; " услышал щелчок курка в "; Format(Now, "h:mm:ss")
End Sub
Private Sub револьвер_Выстрелил()
 If дулоУвиска Then
  убит_ = True
 Else
  Debug.Print имя; " услышал выстрел в "; Format(Now, "h:mm:ss")
 End If
End Sub

Public Property Let имя(s)
 имя_ = s
End Property
Public Property Get имя()
 имя = имя_
End Property

Public Property Get убит()
 убит = убит_
End Property

Public Sub ПриставилДулоКвиску()
 If убит_ Then Exit Sub
 Debug.Print
 Debug.Print имя; " приставил дуло к своему виску в "; Format(Now, "h:mm:ss")
 дулоУвиска = True
 WaitSec RandomInt(1, 6)
End Sub

Public Sub ПередалРевольвер()
 If убит_ Then Exit Sub
 дулоУвиска = False
 WaitSec RandomInt(1, 6)
 Debug.Print имя; " передал револьвер "; Format(Now, "h:mm:ss")
End Sub

Public Sub НажалНаСпусковойКрючок()
 If убит_ Then Exit Sub
 Debug.Print имя; " нажал на спусковой крючок "; Format(Now, "h:mm:ss")
 револьвер.НажатСпусковойКрючок
End Sub

Public Sub КрутанулБарабан()
 If убит_ Then Exit Sub
 Debug.Print имя; " крутанул барабан "; Format(Now, "h:mm:ss")
 револьвер.КрутанутБарабан
End Sub


Attribute VB_Name = "main"
'https://nolongerset.com/raising-custom-events-in-vba/
Option Base 1
Sub КакЭтоБыло()
 Set наган = New револьверСоднимПатроном
 наган.сколькоЗарядный = 6
 Dim дуэлянты(2) As дуэлянт
 Dim j As Integer
 For j = LBound(дуэлянты) To UBound(дуэлянты)
  Set дуэлянты(j) = New дуэлянт
  With дуэлянты(j)
   .имя = "Дуэлянт №" & j
   Set .револьвер = наган
  End With
 Next j
 дуэлянты(1).КрутанулБарабан
 Dim i As Integer
 For i = 1 To наган.сколькоЗарядный
  With дуэлянты(((i - 1) Mod 2) + 1)
   .ПриставилДулоКвиску
   .НажалНаСпусковойКрючок
   If .убит Then GoTo finita
   .ПередалРевольвер
  End With
 Next i
 Debug.Print "— и у меня, граф, бывают осечки, слава Богу."
 Exit Sub
finita:
 Debug.Print "– Finita la comedia! – сказал я доктору."
End Sub

'Returns a random long integer between the Min and Max values, inclusive
Public Function RandomInt(ByVal Min As Long, ByVal Max As Long) As Long
 RandomInt = (Rnd() * (Max - Min)) + Min
End Function

Public Sub WaitSec(Optional sec As Single = 1)
 Dim T0 As Single
 T0 = Timer
 Do
  DoEvents
 Loop Until Timer - T0 >= sec
End Sub

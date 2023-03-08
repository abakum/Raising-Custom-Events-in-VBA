Attribute VB_Name = "main"
'https://nolongerset.com/raising-custom-events-in-vba/
Option Base 1
Sub какЭтоБыло()
 кардиоСтимулятор 'Application.OnTime не может запускать процедуры из модулей класса
 Set наган = New револьверСоднимПатроном
 наган.сколькоЗарядный = 6
 Dim дуэлянты(2) As дуэлянт
 Dim i As Integer
 For i = LBound(дуэлянты) To UBound(дуэлянты)
  Set дуэлянты(i) = New дуэлянт
  With дуэлянты(i)
   .имя = "Дуэлянт №" & i
   Set .револьвер = наган 'дуэлянты используют один и тот же револьвер
  End With
 Next i
 Dim Вернер As New доктор
 дуэлянты(1).КрутанулБарабан
 For i = 1 To наган.сколькоЗарядный
  чьяОчередь = ((i - 1) Mod 2) + 1
  Вернер.ВидитЧтоРевольверВзял дуэлянты(чьяОчередь)
  With дуэлянты(чьяОчередь)
   .ПриставилДулоКвиску
   .НажалНаСпусковойКрючок
    If Not .ПередалРевольвер Then GoTo La_commedia_e_finita
  End With
 Next i
 Debug.Print "— и у меня, граф, бывают осечки, слава Богу."
 GoTo finally
La_commedia_e_finita:
 Debug.Print "– Finita la comedia! – сказал я доктору."
finally:
 Set Вернер = Nothing
 For i = LBound(дуэлянты) To UBound(дуэлянты)
  Set дуэлянты(i) = Nothing
 Next i
 Set наган = Nothing
 кардиоСтимулятор False
End Sub

'возвращает случайное целое между значениями Min и Max включительно
Public Function RandomInt(ByVal Min As Long, ByVal Max As Long) As Long
 RandomInt = (Rnd() * (Max - Min)) + Min
End Function

'блокирующая задержка на sec секунд
Public Sub WaitSec(Optional sec As Single = 1)
 Dim T0 As Single
 T0 = Timer
 Do
  DoEvents
 Loop Until Timer - T0 >= sec
End Sub

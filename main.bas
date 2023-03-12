Attribute VB_Name = "main"
'https://nolongerset.com/raising-custom-events-in-vba/
Option Base 1
#Const отладка = 0
Sub какЭтоБыло()
 Set наган = New револьверСоднимПатроном
 наган.сколькоЗарядный = 6
 Dim дуэлянты(2) As дуэлянт
 Dim и As Integer
 For и = LBound(дуэлянты) To UBound(дуэлянты)
  Set дуэлянты(и) = New дуэлянт
  With дуэлянты(и)
   .имя = "Дуэлянт №" & и
   Set .револьвер = наган 'дуэлянты используют один и тот же револьвер
  End With
 Next и
 Dim Вернер As New доктор
 #If отладка Then
  дуэлянты(1).ЧСС = 10 'тестим события пульс с разной частотой
  Вернер.ВидитЧтоРевольверВзял дуэлянты(1)
  Вернер.СчитаетПульс 'тестим доктора на умение считать пульс на живом пациенте
 #End If
 дуэлянты(1).КрутанулБарабан
 For и = 1 To наган.сколькоЗарядный
  чьяОчередь = ((и - 1) Mod 2) + 1
  Вернер.ВидитЧтоРевольверВзял дуэлянты(чьяОчередь)
  With дуэлянты(чьяОчередь)
   .ПриставилДулоКвиску
   .НажалНаСпусковойКрючок
    If Not .ПередалРевольвер Then GoTo La_commedia_e_finita
  End With
 Next и
 Debug.Print "— и у меня, граф, бывают осечки, слава Богу."
 GoTo finally
La_commedia_e_finita:
 Debug.Print "– Finita la comedia! – сказал я доктору."
finally:
 жди 'надо запускать перед деструкцией классов из которых вызывается жди Me
 Set Вернер = Nothing
 For и = LBound(дуэлянты) To UBound(дуэлянты)
  Set дуэлянты(и) = Nothing
 Next и
 Set наган = Nothing
End Sub

'возвращает случайное целое между значениями мин и макс включительно
Public Function Случайное(ByVal мин As Long, ByVal макс As Long) As Long
 Случайное = (Rnd() * (макс - мин)) + мин
End Function
'не блокирующая событий задержка на 'сек' секунд
Public Sub ждатьСекунд(Optional сек As Single = 1)
 T0 = Timer
 Do
  DoEvents
 Loop Until Timer - T0 >= сек
End Sub
'блокирующая события задержка выполнения на 'сек' секунд
Public Sub висетьСекунд(Optional сек As Single = 1)
 Application.Wait (Now + TimeSerial(0, 0, сек))
End Sub


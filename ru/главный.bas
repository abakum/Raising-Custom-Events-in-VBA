Attribute VB_Name = "главный"
#Const отладка = 1
Sub какЭтоБыло()
 Set наган = New револьверСоднимПатроном
 наган.сколькоЗарядный = 7
 Dim дуэлянты As New Collection
 Dim и As Integer
 Dim второйСтреляется As Boolean
 For и = 1 To 2
  дуэлянты.Add New дуэлянт
  дуэлянты(и).имя = "Дуэлянт №" & и
  Set дуэлянты(и).револьвер = наган 'дуэлянты используют один и тот же револьвер
 Next и
 Dim Вернер As New доктор
 #If отладка Then
  дуэлянты(1).ЧСС = 10 'тестим события пульс с разной частотой
  Вернер.ВидитЧтоРевольверВзял дуэлянты(1)
  Вернер.СчитаетПульс 'тестим доктора на умение считать пульс на живом пациенте
 #End If
 дуэлянты(1).КрутанулБарабан
 For и = 1 To наган.сколькоЗарядный
  Вернер.ВидитЧтоРевольверВзял дуэлянты(1 + второйСтреляется)
  дуэлянты(1 + второйСтреляется).ПриставилДулоКвиску
  дуэлянты(1 + второйСтреляется).НажалНаСпусковойКрючок
  If Not дуэлянты(1 + второйСтреляется).ПередалРевольвер Then GoTo La_commedia_e_finita
  второйСтреляется = Not второйСтреляется
 Next и
 Debug.Print "— и у меня, граф, бывают осечки, слава Богу."
 GoTo finally
La_commedia_e_finita:
 Debug.Print "– Finita la comedia! – сказал я доктору."
finally:
 жди 'надо запускать для отмены всех ожидаемых 'пораX' и для деструкции классов из которых вызывался 'жди Me'
 Set Вернер = Nothing
 Set дуэлянты = Nothing
 Set наган = Nothing
End Sub

'возвращает cлучайное целое между значениями мин и макс включительно
Public Function cлучайное(ByVal мин As Long, ByVal макс As Long) As Long
 cлучайное = (Rnd() * (макс - мин)) + мин
End Function
'не блокирующая событий задержка на 'сек' секунд
Public Sub ждатьСек(Optional сек As Single = 1)
 T0 = Timer
 Do
  DoEvents
 Loop Until Timer - T0 >= сек
End Sub
'блокирующая события задержка выполнения на 'сек' секунд
Public Sub висетьСек(Optional сек As Single = 1)
 Application.wait (Now + TimeSerial(0, 0, сек))
End Sub

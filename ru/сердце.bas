Attribute VB_Name = "сердце"
#If False Then
'пример использование 'жди' в классе
Private позже
Private Sub Class_Initialize()
 жди Me
End Sub
Private Sub Class_Terminate()
 жди Me, False
End Sub
Public Property Let пора(целое)
 позже = Now + TimeSerial(0, 0, 1)
 Application.onTime позже, "пора" & целое
End Property
Public Property Get пора()
 пора = позже
End Property
#End If
Private классы As New Collection
Public Sub жди(Optional класс, Optional конструктор = True)
 Dim и As Integer
 If IsMissing(класс) Then 'жди' вызывать из 'главный'
  On Error Resume Next
  For и = классы.Count To 1 Step -1
   If Not классы(и) Is Nothing Then Application.onTime CallByName(классы(и), "пора", VbGet), "пора" & и, , False
   классы.Remove и
  Next и
 Else
  If конструктор Then 'жди Me' вызывать из конструктора
   For и = классы.Count To 1 Step -1
    If классы(и) Is Nothing Then
     классы.Add класс, after:=и
     классы.Remove и
     пора и
     Exit Sub
    End If
   Next и
   классы.Add класс
   пора классы.Count
  Else
   On Error Resume Next
   For и = классы.Count To 1 Step -1
    If класс Is классы(и) Then 'жди Me, False' вызывать из деструктора
     Application.onTime CallByName(классы(и), "пора", VbGet), "пора" & и, , False
     классы.Add Nothing, after:=и
     классы.Remove и
     Exit Sub
    End If
   Next и
  End If
 End If
End Sub
Public Sub пора(целое)
 On Error Resume Next
 CallByName классы(целое), "пора", VbLet, целое
End Sub
'пораX это заменители дуэлянт(X).пора=X потому, что Application.onTime требует процедуру в обычном модуле
Public Sub пора1(): пора 1: End Sub
Public Sub пора2(): пора 2: End Sub
Public Sub пора3(): пора 3: End Sub
Public Sub пора4(): пора 4: End Sub
Public Sub пора5(): пора 5: End Sub
Public Sub пора6(): пора 6: End Sub
Public Sub пора7(): пора 7: End Sub
Public Sub пора8(): пора 8: End Sub
Public Sub пора9(): пора 9: End Sub
Public Sub пора10(): пора 10: End Sub
Public Sub пора11(): пора 11: End Sub
Public Sub пора12(): пора 12: End Sub
Public Sub пора13(): пора 13: End Sub
Public Sub пора14(): пора 14: End Sub
Public Sub пора15(): пора 15: End Sub
Public Sub пора16(): пора 16: End Sub
Public Sub пора17(): пора 17: End Sub
Public Sub пора18(): пора 18: End Sub
Public Sub пора19(): пора 19: End Sub
Public Sub пора20(): пора 20: End Sub
Public Sub пора21(): пора 21: End Sub
Public Sub пора22(): пора 22: End Sub
Public Sub пора23(): пора 23: End Sub
Public Sub пора24(): пора 24: End Sub
Public Sub пора25(): пора 25: End Sub
Public Sub пора26(): пора 26: End Sub
Public Sub пора27(): пора 27: End Sub
Public Sub пора28(): пора 28: End Sub
Public Sub пора29(): пора 29: End Sub
Public Sub пора30(): пора 30: End Sub
Public Sub пора31(): пора 31: End Sub
Public Sub пора32(): пора 32: End Sub
Public Sub пора33(): пора 33: End Sub
Public Sub пора34(): пора 34: End Sub
Public Sub пора35(): пора 35: End Sub
Public Sub пора36(): пора 36: End Sub
Public Sub пора37(): пора 37: End Sub
Public Sub пора38(): пора 38: End Sub
Public Sub пора39(): пора 39: End Sub
Public Sub пора40(): пора 40: End Sub
Public Sub пора41(): пора 41: End Sub
Public Sub пора42(): пора 42: End Sub
Public Sub пора43(): пора 43: End Sub
Public Sub пора44(): пора 44: End Sub
Public Sub пора45(): пора 45: End Sub
Public Sub пора46(): пора 46: End Sub
Public Sub пора47(): пора 47: End Sub
Public Sub пора48(): пора 48: End Sub
Public Sub пора49(): пора 49: End Sub
Public Sub пора50(): пора 50: End Sub
Public Sub пора51(): пора 51: End Sub
Public Sub пора52(): пора 52: End Sub
Public Sub пора53(): пора 53: End Sub
Public Sub пора54(): пора 54: End Sub
Public Sub пора55(): пора 55: End Sub
Public Sub пора56(): пора 56: End Sub
Public Sub пора57(): пора 57: End Sub
Public Sub пора58(): пора 58: End Sub
Public Sub пора59(): пора 59: End Sub
Public Sub пора60(): пора 60: End Sub
Public Sub пора61(): пора 61: End Sub
Public Sub пора62(): пора 62: End Sub
Public Sub пора63(): пора 63: End Sub
Public Sub пора64(): пора 64: End Sub
Public Sub пора65(): пора 65: End Sub
Public Sub пора66(): пора 66: End Sub
Public Sub пора67(): пора 67: End Sub
Public Sub пора68(): пора 68: End Sub
Public Sub пора69(): пора 69: End Sub
Public Sub пора70(): пора 70: End Sub
Public Sub пора71(): пора 71: End Sub
Public Sub пора72(): пора 72: End Sub
Public Sub пора73(): пора 73: End Sub
Public Sub пора74(): пора 74: End Sub
Public Sub пора75(): пора 75: End Sub
Public Sub пора76(): пора 76: End Sub
Public Sub пора77(): пора 77: End Sub
Public Sub пора78(): пора 78: End Sub
Public Sub пора79(): пора 79: End Sub
Public Sub пора80(): пора 80: End Sub
Public Sub пора81(): пора 81: End Sub
Public Sub пора82(): пора 82: End Sub
Public Sub пора83(): пора 83: End Sub
Public Sub пора84(): пора 84: End Sub
Public Sub пора85(): пора 85: End Sub
Public Sub пора86(): пора 86: End Sub
Public Sub пора87(): пора 87: End Sub
Public Sub пора88(): пора 88: End Sub
Public Sub пора89(): пора 89: End Sub
Public Sub пора90(): пора 90: End Sub
Public Sub пора91(): пора 91: End Sub
Public Sub пора92(): пора 92: End Sub
Public Sub пора93(): пора 93: End Sub
Public Sub пора94(): пора 94: End Sub
Public Sub пора95(): пора 95: End Sub
Public Sub пора96(): пора 96: End Sub
Public Sub пора97(): пора 97: End Sub
Public Sub пора98(): пора 98: End Sub
Public Sub пора99(): пора 99: End Sub
Public Sub пора100(): пора 100: End Sub

Attribute VB_Name = "heart"
#If False Then
'example using 'expect' in a class
Private later
Private Sub Class_Initialize()
 expect Me
End Sub
Private Sub Class_Terminate()
 expect Me, False
End Sub
Public Property Let onTime(i)
 later = Now + TimeSerial(0, 0, 1)
 Application.onTime later, "onTime" & i
End Property
Public Property Get onTime()
 onTime = later
End Property
#End If

Private classes As New Collection
Public Sub expect(Optional Class, Optional constructor = True)
 Dim i As Integer
 If IsMissing(Class) Then 'expect' call from 'main'
  On Error Resume Next
  For i = classes.Count To 1 Step -1
   If Not classes(i) Is Nothing Then Application.onTime CallByName(classes(i), "onTime", VbGet), "onTime" & i, , False
   classes.Remove i
  Next i
 Else
  If constructor Then 'expect Me' call from constructor
   For i = classes.Count To 1 Step -1
    If classes(i) Is Nothing Then
     classes.Add Class, after:=i
     classes.Remove i
     onTime i
     Exit Sub
    End If
   Next i
   classes.Add Class
   onTime classes.Count
  Else
   On Error Resume Next
   For i = classes.Count To 1 Step -1
    If Class Is classes(i) Then 'expect Me, False' call from destructor
     Application.onTime CallByName(classes(i), "onTime", VbGet), "onTime" & i, , False
     classes.Add Nothing, after:=i
     classes.Remove i
     Exit Sub
    End If
   Next i
  End If
 End If
End Sub
Public Sub onTime(i)
 On Error Resume Next
 CallByName classes(i), "onTime", VbLet, i
End Sub
'onTimeX is a substitute for duelist(X).onTime=X because Application.onTime requires a procedure in a regular module
Public Sub onTime1(): onTime 1: End Sub
Public Sub onTime2(): onTime 2: End Sub
Public Sub onTime3(): onTime 3: End Sub
Public Sub onTime4(): onTime 4: End Sub
Public Sub onTime5(): onTime 5: End Sub
Public Sub onTime6(): onTime 6: End Sub
Public Sub onTime7(): onTime 7: End Sub
Public Sub onTime8(): onTime 8: End Sub
Public Sub onTime9(): onTime 9: End Sub
Public Sub onTime10(): onTime 10: End Sub
Public Sub onTime11(): onTime 11: End Sub
Public Sub onTime12(): onTime 12: End Sub
Public Sub onTime13(): onTime 13: End Sub
Public Sub onTime14(): onTime 14: End Sub
Public Sub onTime15(): onTime 15: End Sub
Public Sub onTime16(): onTime 16: End Sub
Public Sub onTime17(): onTime 17: End Sub
Public Sub onTime18(): onTime 18: End Sub
Public Sub onTime19(): onTime 19: End Sub
Public Sub onTime20(): onTime 20: End Sub
Public Sub onTime21(): onTime 21: End Sub
Public Sub onTime22(): onTime 22: End Sub
Public Sub onTime23(): onTime 23: End Sub
Public Sub onTime24(): onTime 24: End Sub
Public Sub onTime25(): onTime 25: End Sub
Public Sub onTime26(): onTime 26: End Sub
Public Sub onTime27(): onTime 27: End Sub
Public Sub onTime28(): onTime 28: End Sub
Public Sub onTime29(): onTime 29: End Sub
Public Sub onTime30(): onTime 30: End Sub
Public Sub onTime31(): onTime 31: End Sub
Public Sub onTime32(): onTime 32: End Sub
Public Sub onTime33(): onTime 33: End Sub
Public Sub onTime34(): onTime 34: End Sub
Public Sub onTime35(): onTime 35: End Sub
Public Sub onTime36(): onTime 36: End Sub
Public Sub onTime37(): onTime 37: End Sub
Public Sub onTime38(): onTime 38: End Sub
Public Sub onTime39(): onTime 39: End Sub
Public Sub onTime40(): onTime 40: End Sub
Public Sub onTime41(): onTime 41: End Sub
Public Sub onTime42(): onTime 42: End Sub
Public Sub onTime43(): onTime 43: End Sub
Public Sub onTime44(): onTime 44: End Sub
Public Sub onTime45(): onTime 45: End Sub
Public Sub onTime46(): onTime 46: End Sub
Public Sub onTime47(): onTime 47: End Sub
Public Sub onTime48(): onTime 48: End Sub
Public Sub onTime49(): onTime 49: End Sub
Public Sub onTime50(): onTime 50: End Sub
Public Sub onTime51(): onTime 51: End Sub
Public Sub onTime52(): onTime 52: End Sub
Public Sub onTime53(): onTime 53: End Sub
Public Sub onTime54(): onTime 54: End Sub
Public Sub onTime55(): onTime 55: End Sub
Public Sub onTime56(): onTime 56: End Sub
Public Sub onTime57(): onTime 57: End Sub
Public Sub onTime58(): onTime 58: End Sub
Public Sub onTime59(): onTime 59: End Sub
Public Sub onTime60(): onTime 60: End Sub
Public Sub onTime61(): onTime 61: End Sub
Public Sub onTime62(): onTime 62: End Sub
Public Sub onTime63(): onTime 63: End Sub
Public Sub onTime64(): onTime 64: End Sub
Public Sub onTime65(): onTime 65: End Sub
Public Sub onTime66(): onTime 66: End Sub
Public Sub onTime67(): onTime 67: End Sub
Public Sub onTime68(): onTime 68: End Sub
Public Sub onTime69(): onTime 69: End Sub
Public Sub onTime70(): onTime 70: End Sub
Public Sub onTime71(): onTime 71: End Sub
Public Sub onTime72(): onTime 72: End Sub
Public Sub onTime73(): onTime 73: End Sub
Public Sub onTime74(): onTime 74: End Sub
Public Sub onTime75(): onTime 75: End Sub
Public Sub onTime76(): onTime 76: End Sub
Public Sub onTime77(): onTime 77: End Sub
Public Sub onTime78(): onTime 78: End Sub
Public Sub onTime79(): onTime 79: End Sub
Public Sub onTime80(): onTime 80: End Sub
Public Sub onTime81(): onTime 81: End Sub
Public Sub onTime82(): onTime 82: End Sub
Public Sub onTime83(): onTime 83: End Sub
Public Sub onTime84(): onTime 84: End Sub
Public Sub onTime85(): onTime 85: End Sub
Public Sub onTime86(): onTime 86: End Sub
Public Sub onTime87(): onTime 87: End Sub
Public Sub onTime88(): onTime 88: End Sub
Public Sub onTime89(): onTime 89: End Sub
Public Sub onTime90(): onTime 90: End Sub
Public Sub onTime91(): onTime 91: End Sub
Public Sub onTime92(): onTime 92: End Sub
Public Sub onTime93(): onTime 93: End Sub
Public Sub onTime94(): onTime 94: End Sub
Public Sub onTime95(): onTime 95: End Sub
Public Sub onTime96(): onTime 96: End Sub
Public Sub onTime97(): onTime 97: End Sub
Public Sub onTime98(): onTime 98: End Sub
Public Sub onTime99(): onTime 99: End Sub
Public Sub onTime100(): onTime 100: End Sub

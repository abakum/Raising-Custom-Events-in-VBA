VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "duelist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
#Const deb = 0
'class Listener handles events of objects via WithEvents
Public WithEvents revolver As RevolverWithSingleCartridge

'class Broadcaster which describes events and calls RaiseEvent in methods/functions/properties body
Public Event pulse()

Private name_
Private later
Private asLater
Private last
Private gunToHead

'event handlers
Private Sub Class_Initialize()
 heartRate = 60
 expect Me
End Sub
Private Sub Class_Terminate()
 #If deb Then
  Debug.Print Format(Now, "h:mm:ss"); " "; name; " Class_Terminate"
 #End If
 last = True
 Set revolver = Nothing
 On Error Resume Next
 expect Me, False 'cancels the expected Application.onTime event in 'later'
End Sub
Private Sub revolver_soundSpinningDrum()
 Debug.Print Format(Now, "h:mm:ss"); " "; name; " heard the sound of a spinning drum"
End Sub
Private Sub revolver_soundSpinningDrumFade()
 Debug.Print Format(Now, "h:mm:ss"); " "; name; " heard the sound of the spinning drum fade away"
End Sub
Private Sub revolver_soundHammer()
 Debug.Print Format(Now, "h:mm:ss"); " "; name; " heard the click of a hammer"
End Sub
Private Sub revolver_fired()
 If gunToHead Then
  Class_Terminate
 Else
  Debug.Print Format(Now, "h:mm:ss"); " "; name; " heard a shot"
 End If
End Sub

'properties
Public Property Let name(s)
 name_ = s
End Property
Public Property Get name()
 name = name_
End Property
Public Property Let heartRate(i)
 '120 hangs Excel because the EarliestTime parameter is rounded to the nearest whole second. That is 60/120~0
 asLater = TimeSerial(0, 0, 60 / IIf(i > 119, 119, i))
End Property
Public Property Let onTime(i)
 If last Then Exit Property
 RaiseEvent pulse
 later = Now + asLater
 Application.onTime later, "onTime" & i
 #If deb Then
  Debug.Print Format(Now, "h:mm:ss"); " "; name; " onTime", i, Format(later, "h:mm:ss")
 #End If
End Property
Public Property Get onTime()
 onTime = later
End Property

'methods
Public Sub putGunToHead()
 If last Then Exit Sub
 Debug.Print
 Debug.Print Format(Now, "h:mm:ss"); " "; name; " put the gun to his head"
 gunToHead = True
 waitSec random(1, 6)
End Sub
Public Function handedRevolver()
 If last Then Exit Function
 gunToHead = False
 waitSec random(1, 6)
 Debug.Print Format(Now, "h:mm:ss"); " "; name; " handed revolver"
 handedRevolver = True
End Function
Public Sub pulledTrigger()
 If last Then Exit Sub
 Debug.Print Format(Now, "h:mm:ss"); " "; name; " pulled the trigger"
 revolver.pulledTrigger
End Sub
Public Sub spunDrum()
 If last Then Exit Sub
 Debug.Print Format(Now, "h:mm:ss"); " "; name; " spun the drum"
 revolver.spunDrum
End Sub

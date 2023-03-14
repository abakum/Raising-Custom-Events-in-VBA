VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "doctor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'class Listener handles events of objects via WithEvents
Public WithEvents revolver As RevolverWithSingleCartridge
Public WithEvents patient As duelist

Private man_ As duelist 'while there is no fired, there is no patient
Private pulse

'event handlers
Private Sub patient_pulse()
 pulse = pulse + 1
End Sub
Private Sub revolver_fired()
 countsPulse
End Sub

'methods
Public Sub countsPulse()
 pulse = 0
 Set patient = man_ 'listening to events patient_pulse
 Debug.Print Format(Now, "h:mm:ss"); " The doctor counts the pulse "; patient.name
 waitSec 6
 If pulse > 0 Then
  Debug.Print Format(Now, "h:mm:ss"); " The doctor said that "; patient.name; " pulse "; pulse * 10
 Else
  Debug.Print Format(Now, "h:mm:ss"); " The doctor said that "; patient.name; " has no pulse"
 End If
 Set patient = Nothing 'stop listening to the patient_pulse events
End Sub
Public Sub SeesThatRevolverTook(man As duelist)
 Set man_ = man
 Set revolver = man.revolver 'listen event revolver fired
End Sub


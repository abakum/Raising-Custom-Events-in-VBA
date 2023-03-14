VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "revolverWithSingleCartridge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
#Const deb = 0
#Const misfire = 0
'class Listener handles events of objects via WithEvents
Public Event fired()
Public Event soundHammer()
Public Event soundSpinningDrum()
Public Event soundSpinningDrumFade()

Private howMuchCharger_
Private chamberWithCartridge
Private chamberUnderHammer

'properties
Public Property Let howMuchCharger(i)
 howMuchCharger_ = i
End Property
Public Property Get howMuchCharger()
 howMuchCharger = howMuchCharger_
End Property

'methods
Public Sub spunDrum()
 RaiseEvent soundSpinningDrum
 Randomize
 chamberUnderHammer = 1
 waitSec random(1, 6)
 chamberWithCartridge = random(0, howMuchCharger_)
 #If deb Then
  #If misfire Then
   chamberWithCartridge = 0 'misfire
  #Else
   chamberWithCartridge = 2 'fired on first pull of the trigger
  #End If
 #End If
 RaiseEvent soundSpinningDrumFade
End Sub
Public Sub pulledTrigger()
 If chamberUnderHammer = howMuchCharger_ Then
  chamberUnderHammer = 1
 Else
  chamberUnderHammer = chamberUnderHammer + 1
 End If
 If chamberUnderHammer = chamberWithCartridge Then
  chamberWithCartridge = 0
  RaiseEvent fired
 Else
  RaiseEvent soundHammer
 End If
End Sub

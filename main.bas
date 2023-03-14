Attribute VB_Name = "main"
'https://nolongerset.com/raising-custom-events-in-vba/
#Const deb = 0
Sub howItWas()
 Dim nagant As New RevolverWithSingleCartridge
 Dim first As Boolean
 Dim i As Integer
 Dim duelists As New Collection
 nagant.howMuchCharger = 7
 first = True
 For i = 1 To 2
  duelists.Add New duelist
  duelists(i).name = "Duelist #" & i
  Set duelists(i).revolver = nagant 'duelists use the same revolver
 Next i
 Dim Verner As New doctor
 #If deb Then
  duelists(1).heartRate = 10 'test pulse events with different frequencies
  Verner.SeesThatRevolverTook duelists(1)
  Verner.countsPulse  'testing a doctor for the ability to count the pulse on a living patient
 #End If
 duelists(1).spunDrum
 For i = 1 To nagant.howMuchCharger
  Verner.SeesThatRevolverTook duelists(first + 2)
  duelists(first + 2).putGunToHead
  duelists(first + 2).pulledTrigger
  If Not duelists(first + 2).handedRevolver Then GoTo La_commedia_e_finita
  first = Not first
 Next i
 Debug.Print "— and I, graf, have misfires, thank God."
 GoTo finally
La_commedia_e_finita:
 Debug.Print "— Finita la comedia! I said to the doctor."
finally:
 expect 'must be run to cancel all expected 'onTimeX' and to destroy classes from which 'expect Me' was called
 Set Verner = Nothing
 Set duelists = Nothing
 Set nagant = Nothing
End Sub

'returns a random integer between min and max values inclusive
Public Function random(ByVal min As Long, ByVal max As Long) As Long
 random = (Rnd() * (max - min)) + min
End Function
'delay allowing events to happen by 'sec' seconds
Public Sub waitSec(Optional sec As Single = 1)
 T0 = Timer
 Do
  DoEvents
 Loop Until Timer - T0 >= sec
End Sub
'event-blocking execution delay of 'sec' seconds
Public Sub hangSec(Optional sec As Single = 1)
 Application.wait (Now + TimeSerial(0, 0, sec))
End Sub

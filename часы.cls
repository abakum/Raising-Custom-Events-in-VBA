VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "часы"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'класс Вещатель который описыват События Event и вызывает RaiseEvent в теле methods/functions/properties
Public Event такт()

'методы
Public Sub стук()
 RaiseEvent такт
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RGBColour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Public red As Integer
Public green As Integer
Public blue As Integer

Public Sub Init(ByVal r As Integer, ByVal g As Integer, ByVal b As Integer)
    red = r
    green = g
    blue = b
End Sub

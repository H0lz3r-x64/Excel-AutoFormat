VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NonVisualRow"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Public name As String
Public Parent As GeneralRow


Public Sub Init(rng As Range, instanceName)
    name = instanceName
    Debug.Print "========================" & vbNewLine & name & " - Init"
    Set Parent = Factory.NewGeneralRow(rng, name)
    
    Debug.Print "========================" & vbNewLine
End Sub


Public Function Form2Config():
    Debug.Print name & " - Form2Config"
    ReDim tmpCfg(Parent.ConfigRange.Columns.Count) As Double
    tmpCfg(0) = (CInt(AutoFormat.checkb_useTableCaptions.Value) * -1)
    tmpCfg(1) = (CInt(AutoFormat.checkb_useHeaders.Value) * -1)
    tmpCfg(2) = (CInt(AutoFormat.checkb_autoCol.Value) * -1)
    
    Parent.Config = tmpCfg
    yes = Parent.Config

End Function

Public Function Config2Form():
    Debug.Print name & " - Config2Form"
    Dim i As Integer
    i = 0
    
    For Each Item In Parent.Config
        Select Case (i)
            Case 0
            AutoFormat.checkb_useTableCaptions.Value = CBool(Item)
            Case 1
            AutoFormat.checkb_useHeaders.Value = CBool(Item)
            Case 2
            AutoFormat.checkb_autoCol.Value = CBool(Item)
        End Select
        i = i + 1
    Next Item
End Function

Public Function ClearConfig()
    ReDim tmpCfg(Parent.ConfigRange.Columns.Count) As Double
    For Each Item In tmpCfg
        Item = 0
    Next Item
    Parent.Config = tmpCfg
End Function

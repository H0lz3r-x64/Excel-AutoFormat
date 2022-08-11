VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NonVisualRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Name As String
Public Parent As GeneralRow


Public Sub Init(rng As Range, instanceName)
    Name = instanceName
    Debug.Print "========================" & vbNewLine & Name & " - Init"
    Set Parent = Factory.NewGeneralRow(rng, Name)
    
    Debug.Print "========================" & vbNewLine
End Sub


Public Function Form2Config():
    Debug.Print Name & " - Form2Config"
    ReDim tmpCfg(Parent.ConfigRange.Columns.Count) As Double
    tmpCfg(0) = (CInt(AutoFormat.checkb_useColumns.Value) * -1)
    tmpCfg(1) = (CInt(AutoFormat.checkb_useHeaders.Value) * -1)
    tmpCfg(2) = (CInt(AutoFormat.checkb_autoCol.Value) * -1)
    
    Parent.Config = tmpCfg
    yes = Parent.Config

End Function

Public Function Config2Form():
    Debug.Print Name & " - Config2Form"
    Dim i As Integer
    i = 0
    
    For Each Item In Parent.Config
        Select Case (i)
            Case 0
            AutoFormat.checkb_useColumns.Value = CBool(Item)
            Case 1
            AutoFormat.checkb_useHeaders.Value = CBool(Item)
            Case 2
            AutoFormat.checkb_autoCol.Value = CBool(Item)
        End Select
        i = i + 1
    Next Item
End Function


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VisualRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public PreviewRange As Range
Public Parent As GeneralRow
Public Name As String

Enum Border
 None = 0
 Thin = 1
 Normal = 2
 Thick = 3
End Enum


Sub Init(rng As Range, previewRng As Range, Optional instanceName)
    Name = instanceName
    Debug.Print "========================" & vbNewLine & Name & " - Init"
    Set Parent = Factory.NewGeneralRow(rng, Name)
    
    Set PreviewRange = previewRng
    
    Debug.Print "========================" & vbNewLine
End Sub


Public Function Form2Config():
    Debug.Print Name & " - Form2Config"
    ReDim tmpCfg(Parent.ConfigRange.Columns.Count) As Double
    Dim i As Integer
    i = 0
    
    tmpCfg(0) = (CInt(AutoFormat.checkb_bold.Value) * -1)
    tmpCfg(1) = (CInt(AutoFormat.checkb_underlined.Value) * -1)
    tmpCfg(2) = (CInt(AutoFormat.checkb_italic.Value) * -1)
    tmpCfg(3) = (CInt(AutoFormat.checkb_wordWrap.Value) * -1)
    tmpCfg(4) = AutoFormat.tb_interiour.Value
    tmpCfg(5) = AutoFormat.tb_altInteriour.Value
    tmpCfg(6) = AutoFormat.tb_textColour.Value
    tmpCfg(7) = 0 'TODO Border
    tmpCfg(8) = (CInt(AutoFormat.checkb_altRowColour.Value) * -1)
    
    Parent.Config = tmpCfg
End Function

Public Function Config2Form():
    Debug.Print Name & " - Config2Form"
    Dim i As Integer
    i = 0
    
    For Each Item In Parent.Config
        Select Case (i)
            Case 0 'Bold
            AutoFormat.checkb_bold.Value = CBool(Item)
            Case 1 'Underlined
            AutoFormat.checkb_underlined.Value = CBool(Item)
            Case 2 'Italic
            AutoFormat.checkb_italic.Value = CBool(Item)
            Case 3 'WordWrap
            AutoFormat.checkb_wordWrap.Value = CBool(Item)
            Case 4 'interiour colour
            AutoFormat.tb_interiour.Value = Item
            AutoFormat.tb_interiour.BorderColor = Item
            Case 5 'alternate interiour colour
            AutoFormat.tb_altInteriour.Value = Item
            AutoFormat.tb_altInteriour.BorderColor = Item
            Case 6 'text colour
            AutoFormat.tb_textColour.Value = Item
            AutoFormat.tb_textColour.BorderColor = Item
            Case 7 'Border
            ' TODO Border Style
            Case 8 'Alternate Row Colours
            AutoFormat.checkb_altRowColour.Value = CBool(Item)
        End Select
        i = i + 1
    Next Item
End Function



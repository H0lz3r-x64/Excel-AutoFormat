VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VisualRow"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Public PreviewRange As Range
Public Parent As GeneralRow
Public name As String

Enum Border
 None = 0
 Thin = 1
 Normal = 2
 Thick = 3
End Enum


Sub Init(rng As Range, previewRng As Range, Optional instanceName)
    name = instanceName
    Debug.Print "========================" & vbNewLine & name & " - Init"
    Set Parent = Factory.NewGeneralRow(rng, name)
    
    Set PreviewRange = previewRng
    
    Debug.Print "========================" & vbNewLine
End Sub


Public Function Form2Config():
    Debug.Print name & " - Form2Config"
    ReDim tmpCfg(Parent.ConfigRange.Columns.Count) As Double
    
    tmpCfg(0) = (CInt(AutoFormat.checkb_bold.Value) * -1)
    tmpCfg(1) = (CInt(AutoFormat.checkb_underlined.Value) * -1)
    tmpCfg(2) = (CInt(AutoFormat.checkb_italic.Value) * -1)
    tmpCfg(3) = (CInt(AutoFormat.checkb_wordWrap.Value) * -1)
    tmpCfg(4) = AutoFormat.tb_interiour.Value
    tmpCfg(5) = AutoFormat.tb_altInteriour.Value
    tmpCfg(6) = AutoFormat.tb_textColour.Value
    tmpCfg(7) = AutoFormat.cb_borderThickness.ListIndex
    tmpCfg(8) = (CInt(AutoFormat.checkb_altRowColour.Value) * -1)
    If AutoFormat.cb_allBorders.Value = True Then tmpCfg(9) = 9 Else: tmpCfg(9) = BorderStr2Double(AutoFormat.tb_borderLocation.Text)
    
    Parent.Config = tmpCfg
End Function

Public Function Config2Form():
    Debug.Print name & " - Config2Form"
    ReDim tmpCfg(Parent.ConfigRange.Columns.Count) As Double
    tmpCfg = Parent.Config
    
    'Bold
    AutoFormat.checkb_bold.Value = CBool(tmpCfg(0))
    'Underlined
    AutoFormat.checkb_underlined.Value = CBool(tmpCfg(1))
    'Italic
    AutoFormat.checkb_italic.Value = CBool(tmpCfg(2))
    'WordWrap
    AutoFormat.checkb_wordWrap.Value = CBool(tmpCfg(3))
    'interiour colour
    AutoFormat.tb_interiour.Value = tmpCfg(4)
    AutoFormat.tb_interiour.BorderColor = tmpCfg(4)
    'alternate interiour colour
    AutoFormat.tb_altInteriour.Value = tmpCfg(5)
    AutoFormat.tb_altInteriour.BorderColor = tmpCfg(5)
    'text colour
    AutoFormat.tb_textColour.Value = tmpCfg(6)
    AutoFormat.tb_textColour.BorderColor = tmpCfg(6)
    'Border thickness
    AutoFormat.cb_borderThickness.ListIndex = tmpCfg(7)
    'Alternate Row Colours
    AutoFormat.checkb_altRowColour.Value = CBool(tmpCfg(8))
    'Border Location
    If tmpCfg(9) = 9 Then
        AutoFormat.tb_borderLocation.Value = 0
        AutoFormat.cb_allBorders.Value = True
    Else
        AutoFormat.tb_borderLocation.Value = BorderDouble2Str(tmpCfg(9))
        AutoFormat.cb_allBorders.Value = False
    End If
    
End Function

Public Function ClearConfig()
    ReDim tmpCfg(Parent.ConfigRange.Columns.Count) As Double
    For Each Item In tmpCfg
        Item = 0
    Next Item
    Parent.Config = tmpCfg
End Function

'-------------------------------------------------------------
Private Function BorderStr2Double(ByVal inp As String)
    Dim outp As Double
    outp = 0
    '1000
    If InStr(inp, "bottom") Then
        outp = outp + 1000
    End If
    '0100
    If InStr(inp, "top") Then
        outp = outp + 100
    End If
    '0010
    If InStr(inp, "left") Then
        outp = outp + 10
    End If
    '0001
    If InStr(inp, "right") Then
        outp = outp + 1
    End If
    BorderStr2Double = outp
End Function

Private Function BorderDouble2Str(ByVal inp As Double)
    Dim outp As String
    Dim borderArr(0 To 3) As String
    borderArr(0) = "bottom"
    borderArr(1) = "top"
    borderArr(2) = "left"
    borderArr(3) = "right"
    
    op_val = 1000
    For i = 0 To 3
        If CInt(inp / op_val) <> 0 Then
            'places comma
            If outp <> "" Then outp = outp & ", "
            'reduce inp for next iteration
            inp = inp Mod op_val
            'append location to ouput
            outp = outp & borderArr(i)
        End If
        'Next operator value
        op_val = op_val / 10
    Next i
    
    BorderDouble2Str = outp
End Function

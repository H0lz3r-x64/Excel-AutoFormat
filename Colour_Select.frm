VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Colour_Select 
   Caption         =   "UserForm1"
   ClientHeight    =   2370
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5745
   OleObjectBlob   =   "Colour_Select.frx":0000
   StartUpPosition =   2  'Bildschirmmitte
End
Attribute VB_Name = "Colour_select"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ColourLastEdit As Integer
Dim CallerID As Integer


Private Sub UserForm_Initialize()
    tb_colourCode.Value = Factory.ColourSelect_Colour
    CallerID = Factory.ColourSelect_CallerID
    ColourLastEdit = 1
    btn_set_Click
End Sub


'Buttons ---------
Private Sub btn_set_Click()
    If ColourLastEdit = 1 Then
        Dim Value
        Value = Util.Colour2RGB(tb_colourCode.Value)
        tb_red.Text = Value(0)
        tb_green.Text = Value(1)
        tb_blue.Text = Value(2)
        ColourLastEdit = -1
    ElseIf ColourLastEdit = 0 Then
        Dim r As Integer, g As Integer, b As Integer
        r = tb_red.Text
        g = tb_green.Text
        b = tb_blue.Text
        tb_colourCode.Text = rgb(r, g, b)
        ColourLastEdit = -1
    End If
    ColourPreview.BackColor = tb_colourCode.Value
    btn_set.Enabled = False
    btn_Ok.Enabled = True
End Sub

Private Sub btn_EnterPresets_Click()
    Call Colour_Presets.Show
End Sub

Private Sub btn_Ok_Click()
    Select Case CallerID
        Case 0
        AutoFormat.tb_interiour.Value = tb_colourCode.Value
        Case 1
        AutoFormat.tb_altInteriour.Value = tb_colourCode.Value
        Case 2
        AutoFormat.tb_textColour.Value = tb_colourCode.Value
    End Select
    Unload Me
End Sub

Private Sub btn_cancel_Click()
    Unload Me
End Sub


'Text changes --------
Private Sub tb_blue_Change()
    ColourLastEdit = 0
    btn_set.Enabled = True
    btn_Ok.Enabled = False
End Sub

Private Sub tb_green_Change()
    ColourLastEdit = 0
    btn_set.Enabled = True
    btn_Ok.Enabled = False
End Sub

Private Sub tb_red_Change()
    ColourLastEdit = 0
    btn_set.Enabled = True
    btn_Ok.Enabled = False
End Sub

Private Sub tb_colourCode_Change()
    ColourLastEdit = 1
    btn_set.Enabled = True
    btn_Ok.Enabled = False
End Sub


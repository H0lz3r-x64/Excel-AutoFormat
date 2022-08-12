VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AutoFormat 
   Caption         =   "UserForm1"
   ClientHeight    =   6165
   ClientLeft      =   1260
   ClientTop       =   5025
   ClientWidth     =   9615
   OleObjectBlob   =   "AutoFormat.frx":0000
   StartUpPosition =   2  'Bildschirmmitte
End
Attribute VB_Name = "AutoFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private selectedSetting As Integer
Private Column As Object
Private Header As Object
Private Body As Object
Private General As Object


Private Sub UserForm_Initialize()
    With wsConsole
        Set Column = Factory.NewVisualRow(.Range("AJ5", "AR5"), .Range("AA3", "AD3"), "Column")
        Set Header = Factory.NewVisualRow(.Range("AJ6", "AR6"), .Range("AA4", "AD4"), "Header")
        Set Body = Factory.NewVisualRow(.Range("AJ7", "AR7"), .Range("AA5", "AD5"), "Body")
        
        Set General = Factory.NewNonVisualRow(.Range("AJ10", "AL10"), "General")
    End With
    selectedSetting = -1
    Fill_cb_setting_selector
    
End Sub

Private Sub Fill_cb_setting_selector()
    With cb_setting_selector
        .AddItem ("General")
        .AddItem ("Column")
        .AddItem ("Header")
        .AddItem ("Body")
        .Text = "General"
    End With
End Sub

Private Sub SafeForm()
'Save Form to Config
    Select Case selectedSetting
        Case 0
            General.Form2Config
        Case 1
            Column.Form2Config
        Case 2
            Header.Form2Config
        Case 3
            Body.Form2Config
    End Select
End Sub

'Events ---------

Private Sub cb_setting_selector_Change()
    'Save Form to Config
    SafeForm
    'Change Form
    Select Case (cb_setting_selector.ListIndex)
        Case 0
            If selectedSetting <> 0 And selectedSetting <> -1 Then
                tmpL = Row_Frame.Left
                tmpT = Row_Frame.Top
                Row_Frame.Left = General_Frame.Left
                Row_Frame.Top = General_Frame.Top
                General_Frame.Left = tmpL
                General_Frame.Top = tmpT
            End If
        Case 1 To 3
            If selectedSetting = 0 Then
                tmpL = General_Frame.Left
                tmpT = General_Frame.Top
                General_Frame.Left = Row_Frame.Left
                General_Frame.Top = Row_Frame.Top
                Row_Frame.Left = tmpL
                Row_Frame.Top = tmpT
            End If
    End Select
    'Set new selected setting
    selectedSetting = cb_setting_selector.ListIndex
    'Load Config in Form
    Select Case selectedSetting
        Case 0
            'tmp = Util.cvt2StrArr(General.Parent.Config)
'            Util.msgBoxStrArray (tmp)
            General.Config2Form
        Case 1
            'tmp = Util.cvt2StrArr(Column.Parent.Config)
            'Util.msgBoxStrArray (tmp)
            Column.Config2Form
        Case 2
            'tmp = Util.cvt2StrArr(Header.Parent.Config)
            'Util.msgBoxStrArray (tmp)
            Header.Config2Form
        Case 3
            'tmp = Util.cvt2StrArr(Body.Parent.Config)
            'Util.msgBoxStrArray (tmp)
            Body.Config2Form
    End Select
End Sub

Private Sub btn_interiour_Click()
    Factory.Form_ColourSelect 0, tb_interiour.Value
End Sub

Private Sub btn_altInteriour_Click()
    Factory.Form_ColourSelect 1, tb_altInteriour.Value
End Sub

Private Sub btn_textColour_Click()
    Factory.Form_ColourSelect 2, tb_textColour.Value
End Sub


Private Sub btn_saveConfig_Click()
    SafeForm
    General.Parent.WriteConfig
    Column.Parent.WriteConfig
    Header.Parent.WriteConfig
    Body.Parent.WriteConfig
End Sub

Private Sub btn_start_Click()
    SafeForm
    Call Format.ClearOutput
    Call Format.Format(Column, Header, Body, General)
    wsOutput.Select
End Sub

Private Sub checkb_altRowColour_Change()
    If (CInt(checkb_altRowColour.Value) * -1) = 0 Then
        tb_altInteriour.Enabled = False
        btn_altInteriour.Enabled = False
        lbl_altInteriour.Enabled = False
    Else
        tb_altInteriour.Enabled = True
        btn_altInteriour.Enabled = True
        lbl_altInteriour.Enabled = True
    End If
End Sub

VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AutoFormat 
   OleObjectBlob   =   "AutoFormat.frx":0000
   Caption         =   "Auto Format"
   ClientHeight    =   5385
   ClientLeft      =   1260
   ClientTop       =   5025
   ClientWidth     =   9720
   StartUpPosition =   2  'Bildschirmmitte
   TypeInfoVer     =   2281
End
Attribute VB_Name = "AutoFormat"
Attribute VB_Base = "0{1DBAFB5D-933C-44C1-955B-BB8BD66E82EC}{65FADB7A-AC90-40C7-8D11-8EB6E08C24F3}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private selectedSetting As Integer
Private TableCaptions As Object
Private Header As Object
Private Body As Object
Private General As Object


Private Sub UserForm_Initialize()
    startcol = "AJ"
    endcol = "AS"
    With wsConsole
        'TableCaptions
        Set TableCaptions = Factory.NewVisualRow(.Range(startcol & "5", endcol & "5"), _
        .Range("AA3", "AD3"), "TableCaptions")
        'Header
        Set Header = Factory.NewVisualRow(.Range(startcol & "6", endcol & "6"), _
        .Range("AA4", "AD4"), "Header")
        'Body
        Set Body = Factory.NewVisualRow(.Range(startcol & "7", endcol & "7"), _
        .Range("AA5", "AD5"), "Body")
        
        Set General = Factory.NewNonVisualRow(.Range("AJ10", "AL10"), "General")
    End With
    selectedSetting = -1
    Fill_cb_setting_selector
    Fill_cb_borderThickness
    
    cb_allBorders_Change
End Sub

Private Sub Fill_cb_setting_selector()
    With cb_setting_selector
        .AddItem ("General")
        .AddItem ("Table Captions")
        .AddItem ("Header")
        .AddItem ("Body")
        .Text = "General"
    End With
End Sub

Private Sub Fill_cb_borderThickness()
    With cb_borderThickness
        .AddItem ("hairline")
        .AddItem ("thin")
        .AddItem ("medium")
        .AddItem ("thick")
    End With
End Sub

Private Sub SafeForm()
    'Avoid alt colour bug with header or list captions
    If selectedSetting = 1 Or selectedSetting = 2 Then tb_altInteriour.Value = tb_interiour.Value
    'Save Form to Config
    Select Case selectedSetting
        Case 0
            General.Form2Config
        Case 1
            TableCaptions.Form2Config
        Case 2
            Header.Form2Config
        Case 3
            Body.Form2Config
    End Select
End Sub

Private Sub Current_Cfg2Form()
'Save Form to Config
    Select Case selectedSetting
        Case 0
            General.Config2Form
        Case 1
            TableCaptions.Config2Form
        Case 2
            Header.Config2Form
        Case 3
            Body.Config2Form
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
            General.Config2Form
            lbl_useAltColour.Enabled = False
            checkb_altRowColour.Enabled = False
        Case 1
            TableCaptions.Config2Form
            lbl_useAltColour.Enabled = False
            checkb_altRowColour.Enabled = False
        Case 2
            Header.Config2Form
            lbl_useAltColour.Enabled = False
            checkb_altRowColour.Enabled = False
        Case 3
            Body.Config2Form
            lbl_useAltColour.Enabled = True
            checkb_altRowColour.Enabled = True
    End Select
End Sub

Private Sub cb_allBorders_Change()
    If cb_allBorders.Value = True Then
        frame_customBorder.Visible = False
    Else
        frame_customBorder.Visible = True
    End If
End Sub

Private Sub tb_interiour_Change()
    tb_interiour.BorderColor = tb_interiour.Value
    If selectedSetting = 1 Or selectedSetting = 2 Then
        tb_altInteriour.Value = tb_interiour.Value
        tb_altInteriour.BorderColor = tb_interiour.Value
    End If
End Sub

Private Sub btn_interiour_Click()
    Factory.Form_ColourSelect 0, tb_interiour.Value
    tb_interiour.BorderColor = tb_interiour.Value
End Sub

Private Sub btn_altInteriour_Click()
    Factory.Form_ColourSelect 1, tb_altInteriour.Value
    tb_altInteriour.BorderColor = tb_altInteriour.Value
End Sub

Private Sub btn_textColour_Click()
    Factory.Form_ColourSelect 2, tb_textColour.Value
    tb_textColour.BorderColor = tb_textColour.Value
End Sub

Private Sub btn_borderLocation_Click()
    Border_Location.Show
End Sub


'Header Buttons
Private Sub btn_inputFilePath_Click()
    answer = MsgBox("Do you want to load the input sheet data from another excel file?" & vbNewLine & "Current input sheet data will be lost.", vbQuestion + vbYesNo + vbDefaultButton1)
    If answer = vbYes Then
        Application.ScreenUpdating = False
        Dim wbSource As Workbook, wsSource As Worksheet
        Path = Util.openDialog
        Index = CInt(InputBox("Enter the index of the worksheet", "Select Worksheet"))
        
        Set wbSource = Workbooks.Open(Path)
        Set wsSource = wbSource.Worksheets(Index)
        wsSource.Cells.Copy Destination:=wsInput.Range("A1")
        wbSource.Close
        Application.ScreenUpdating = True
    End If
End Sub

Private Sub btn_clearForm_Click()
    answer = MsgBox("Do you really want to clear the form?" & vbNewLine & "All unsaved settings will be lost.", vbQuestion + vbYesNo + vbDefaultButton1)
    If answer = vbYes Then
        'Clear object config
        General.ClearConfig
        TableCaptions.ClearConfig
        Header.ClearConfig
        Body.ClearConfig
        
        'Load object config into current setting screen, so it gets saved corretly in the step after
        Current_Cfg2Form
        
        'Change setting screen to General
        cb_setting_selector.ListIndex = 0
    End If
End Sub


Private Sub btn_loadConfig_Click()
    answer = MsgBox("Loading the saved config will overwrite all unsaved settings." & vbNewLine & "Do you still wish to proceed?", vbQuestion + vbYesNo + vbDefaultButton1)
    If answer = vbYes Then
        'Load table config into object config
        General.Parent.ReadConfig
        TableCaptions.Parent.ReadConfig
        Header.Parent.ReadConfig
        Body.Parent.ReadConfig
        'Load object config into current setting screen, so it gets saved corretly in the step after
        Current_Cfg2Form
        
'        Util.msgBoxStrArray (Util.cvt2StrArr(TableCaptions.Parent.Config))
    End If
End Sub


Private Sub btn_saveConfig_Click()
    answer = MsgBox("Saving the form will overwrite the current saved config." & vbNewLine & "(Later on the option to save multiple configs will be added.)" & vbNewLine & vbNewLine & "Do you still wish to proceed?", vbQuestion + vbYesNo + vbDefaultButton1)
    If answer = vbYes Then
        SafeForm
        General.Parent.WriteConfig
        TableCaptions.Parent.WriteConfig
        Header.Parent.WriteConfig
        Body.Parent.WriteConfig
        
'        Util.msgBoxStrArray (Util.cvt2StrArr(TableCaptions.Parent.Config))
    End If
End Sub

Private Sub btn_start_Click()
    Application.ScreenUpdating = False
    SafeForm
    Call Format.ClearOutput
    Call Format.Format(TableCaptions, Header, Body, General)
    wsOutput.Select
    Application.ScreenUpdating = True
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

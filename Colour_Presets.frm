VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Colour_Presets 
   OleObjectBlob   =   "Colour_Presets.frx":0000
   Caption         =   "UserForm2"
   ClientHeight    =   3210
   ClientLeft      =   150
   ClientTop       =   570
   ClientWidth     =   5745
   StartUpPosition =   2  'Bildschirmmitte
   TypeInfoVer     =   304
End
Attribute VB_Name = "Colour_Presets"
Attribute VB_Base = "0{C60E70C2-7FDB-4296-A1B2-403CB9CC0F1D}{7D8871A0-779C-487C-BE42-164B1F42400B}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private ColCustImg As Collection

Private Sub UserForm_Initialize()
    Dim ctl As Control
    Dim CustomImage As CustomImage
    
    Set ColCustImg = New Collection
    
    For Each ctl In Me.Controls
    
        If TypeName(ctl) = "Image" Then
            Set CustomImage = New CustomImage
            CustomImage.InitialiseCustomImage ctl, 1, ctl.name
            ColCustImg.Add CustomImage
        End If
    Next ctl
End Sub

Private Sub btn_choose_Click()
    Dim rgb, colour
    colour = ColourPreview.BackColor
    rgb = Util.Colour2RGB(colour)
    
    'Pass to Parent Form
    With Colour_select
        .ColourPreview.BackColor = colour
        .tb_red.Text = rgb(0)
        .tb_green.Text = rgb(1)
        .tb_blue.Text = rgb(2)
        .tb_colourCode.Text = colour
        .btn_set.Enabled = False
        .btn_Ok.Enabled = True
    End With
    Unload Me
End Sub

Private Sub btn_cancel_Click()
    Unload Me
End Sub

    

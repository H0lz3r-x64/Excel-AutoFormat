VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Border_Location 
   OleObjectBlob   =   "Border_Location.frx":0000
   Caption         =   "Borders"
   ClientHeight    =   2235
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6840
   StartUpPosition =   1  'Fenstermitte
   TypeInfoVer     =   47
End
Attribute VB_Name = "Border_Location"
Attribute VB_Base = "0{41A445E5-56AD-4F47-9CD6-248DD2444108}{9072DE3F-F5B9-43E9-AEF3-9838FE42A1F1}"
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
    
    If InStr(AutoFormat.tb_borderLocation.Text, "bottom") Then
        img_bottom.BackColor = 0
    End If
    If InStr(AutoFormat.tb_borderLocation.Text, "top") Then
        img_top.BackColor = 0
    End If
    If InStr(AutoFormat.tb_borderLocation.Text, "left") Then
        img_left.BackColor = 0
    End If
    If InStr(AutoFormat.tb_borderLocation.Text, "right") Then
        img_right.BackColor = 0
    End If
    
    For Each ctl In Me.Controls
        If TypeName(ctl) = "Image" Then
            Set CustomImage = New CustomImage
            CustomImage.InitialiseCustomImage ctl, 2, ctl.name
            ColCustImg.Add CustomImage
        End If
    Next ctl
End Sub


Private Sub btn_cancel_Click()
    Unload Me
End Sub

Private Sub btn_choose_Click()
    Dim toggledBorders As String
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "Image" Then
            If ctrl.BackColor = 0 Then
                If toggledBorders <> "" Then
                    toggledBorders = toggledBorders & ", "
                End If
                toggledBorders = toggledBorders & Util.cut_string(ctrl.name, 5, Len(ctrl.name))
            End If
        End If
    Next ctrl
    AutoFormat.tb_borderLocation.Value = toggledBorders
    
    Unload Me
End Sub

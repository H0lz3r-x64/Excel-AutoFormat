VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CustomImage"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private WithEvents CustomImage As Image
Attribute CustomImage.VB_VarHelpID = -1
Private UserForm As Integer
Private ctrlName As String

Public Sub InitialiseCustomImage(imgToCustomise As Image, Form As Integer, name As String)
    ctrlName = name
    Set CustomImage = imgToCustomise
    UserForm = Form
End Sub

Private Sub customImage_Click()
    Select Case UserForm
        Case 1
        UpdatePreview_Colour_Presets
        Case 2
        UpdatePreview_Border_Location
    End Select
End Sub

Private Sub UpdatePreview_Colour_Presets()
    For Each ctrl In Colour_Presets.Controls
        If TypeName(ctrl) = "Image" Then
            If ctrl.name = "ColourPreview" Then
                ctrl.BackColor = CustomImage.BackColor
            End If
        End If
    Next ctrl
End Sub

Private Sub UpdatePreview_Border_Location()
    For Each ctrl In Border_Location.Controls
        If TypeName(ctrl) = "Image" Then
            If ctrl.name = ctrlName Then
                'Switch Colour
                If ctrl.BackColor = 0 Then
                    ctrl.BackColor = 14737632
                ElseIf ctrl.BackColor = 14737632 Then
                    ctrl.BackColor = 0
                End If
            End If
        End If
    Next ctrl
    
End Sub

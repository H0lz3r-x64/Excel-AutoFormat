VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CustomImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private WithEvents CustomImage As Image
Attribute CustomImage.VB_VarHelpID = -1

Public Sub InitialiseCustomImage(imgToCustomise As Image)
    Set CustomImage = imgToCustomise
End Sub

Private Sub customImage_Click()
    UpdatePreview
End Sub

Private Sub UpdatePreview()
    For Each ctrl In Colour_Presets.Controls
        If TypeName(ctrl) = "Image" Then
            If ctrl.Name = "ColourPreview" Then
                ctrl.BackColor = CustomImage.BackColor
            End If
        End If
    Next ctrl
End Sub

Attribute VB_Name = "Factory"
'Factory acting as constructor

'Userform pass values
Public ColourSelect_CallerID As Integer
Public ColourSelect_Colour As Double


Function NewGeneralRow(cfgRng As Range, childInstanceName As String)
    Set NewGeneralRow = New GeneralRow
    NewGeneralRow.Init cfgRng, childInstanceName
End Function

Function NewVisualRow(cfgRng As Range, previewRng As Range, Optional instanceName As String)
    Set NewVisualRow = New VisualRow
    If instanceName = "" Then
        instanceName = "VisualRow"
    End If
    NewVisualRow.Init cfgRng, previewRng, instanceName
End Function

Function NewNonVisualRow(cfgRng As Range, Optional instanceName As String)
    Set NewNonVisualRow = New NonVisualRow
    If instanceName = "" Then
        instanceName = "NonVisualRow"
    End If
    NewNonVisualRow.Init cfgRng, instanceName
End Function

Sub Form_ColourSelect(id As Integer, colour As Double)
    ColourSelect_CallerID = id
    ColourSelect_Colour = colour
    Colour_select.Show
End Sub


' Not in usage
Function NewRGBColour(r As Integer, g As Integer, b As Integer)
    Set NewRGB = New RGBColour
    NewRGB.Init r, g, b
End Function

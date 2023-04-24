VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GeneralRow"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Dim name As String
Dim rng As Range
Dim cfg() As Double

'Properties ---------
Public Property Get Config() As Double()
    Config = cfg
End Property

Public Property Let Config(tmpCfg() As Double)
    cfg = tmpCfg
End Property

Public Property Get ConfigRange() As Range
    Set ConfigRange = rng
End Property

Public Property Let ConfigRange(tmpRng As Range)
    rng = tmpRng
End Property

     
Public Sub Init(tmpRng As Range, ByVal childName As String)
    Set rng = tmpRng
    name = "Parent " & childName
    Debug.Print "   ------------------------" & vbNewLine & name & " - Init"
    'Declare & Initialize config
    ReDim cfg(rng.Columns.Count) As Double
    For Each cell In rng
        cfg(i) = 0
    Next cell
    Debug.Print "   ------------------------"
End Sub

' Reads the config from worksheet and saves in cfg array
Public Function ReadConfig() As Double()
    Debug.Print name & " - ReadConfig"
    ReDim cfg(rng.Columns.Count) As Double
    Dim i As Integer
    i = 0
    
    For Each cell In rng
        If Not IsEmpty(cell.Value2) Then
            cfg(i) = cell.Value2
        Else
            cfg(1) = 0
        End If
        i = i + 1
    Next cell
    ReadConfig = cfg
End Function

' Writes the config in worksheet from cfg array
Public Function WriteConfig(Optional ByVal vConfig As Variant)
    Debug.Print name & " - WriteConfig"
    Dim i As Integer
    i = 0

    If IsMissing(vConfig) Then
        For Each cell In rng
            cell.Value2 = cfg(i)
            i = i + 1
        Next cell
    Else
        For Each cell In rng
            cell.Value2 = vConfig(i)
            i = i + 1
        Next cell
    End If
    'Leave timestamp
    rng(0).Offset(0, -1) = Now()
End Function

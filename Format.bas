Attribute VB_Name = "Format"
Dim rowCount As Long, colCount As Long
Dim useTableCaptions As Boolean
Dim useHeaders As Boolean
Dim altRowColours As Boolean
Dim AutoCol As Boolean
Dim BorderThickness As Integer
Dim rng As Range


Public Sub ClearOutput()
    wsOutput.Cells.Clear
    wsOutput.Cells.ClearFormats
    MsgBox "Cleared Output Sheet"
End Sub

Public Sub ClearInput()
    answer = MsgBox("Do you really want to clear the input sheet?", vbQuestion + vbYesNo + vbDefaultButton2)
    If answer = vbYes Then
        wsInput.Cells.Clear
        wsInput.Cells.ClearFormats
        MsgBox "Cleared Input Sheet"
    End If
    
End Sub

Public Sub Format(ByVal TableCaptionConf As VisualRow, ByVal HeaderConf As VisualRow, ByVal BodyConf As VisualRow, ByVal GeneralConf As NonVisualRow)
    rowCount = wsInput.UsedRange.Rows(wsInput.UsedRange.Rows.Count).row
    colCount = wsInput.UsedRange.Columns(wsInput.UsedRange.Columns.Count).Column
    
    'Copy over cells
    wsInput.Cells.Copy Destination:=wsOutput.Range("A1")
    
    'General
    FormatNonVisual GeneralConf.Parent.Config
    
    'Row
    rowFromHeader = 0
    For row = 1 To rowCount
        'Get range of current row
        Set rng = Range(wsOutput.Cells(row, 1), wsOutput.Cells(row, colCount))
        'Determine type of current row
        If row = 1 And useTableCaptions = True Then
            'TableCaptions
            rowFromHeader = row
            FormatVisual rowFromHeader, rng, TableCaptionConf.Parent.Config
        ElseIf wsOutput.Cells(row, 1).Text <> "" And useHeaders = True Then
            'Heading
            FormatVisual rowFromHeader, rng, HeaderConf.Parent.Config
        Else
            'Body
            FormatVisual rowFromHeader, rng, BodyConf.Parent.Config
        End If
        rowFromHeader = rowFromHeader + 1
    Next row
    
    'AutoCol
    If AutoCol = True Then wsOutput.UsedRange.Columns.AutoFit
End Sub


Private Sub FormatVisual(ByVal row As Integer, rng As Range, Config As Variant)
    'Alternating Rows
    altRowColours = CBool(Config(8))
    'Bold
    rng.Font.Bold = CBool(Config(0))
    'Underlined
    rng.Font.Underline = CBool(Config(1))
    'Italic
    rng.Font.Italic = CBool(Config(2))
    'WordWrap
    rng.WrapText = CBool(Config(3))
    'Cell Interior
    rng.Interior.Color = Config(4)
    'Alt. Cell Interior
    If altRowColours = True Then
        If (row) Mod 2 = 1 Then rng.Interior.Color = Config(5)
    End If
    'Text Colour
    rng.Font.Color = Config(6)
    'Border thickness
    BorderThickness = ThicknessIndex2Enumeration(Config(7))
    'Border location
    If Config(9) = 9 Then DrawFullBorder BorderThickness Else: DrawCustomBorder Config(9), BorderThickness
End Sub

Private Sub FormatNonVisual(Config As Variant)
    'Use TableCaptions
    useTableCaptions = CBool(Config(0))
    'Use Headers
    useHeaders = CBool(Config(1))
    'AutoCol
    AutoCol = CBool(Config(2))
End Sub

Private Sub DrawFullBorder(thickness As Integer)
    For Each cell In rng
        cell.BorderAround LineStyle:=xlContinuous, Weight:=thickness
    Next cell
End Sub

Private Sub DrawCustomBorder(ByVal inp As Double, thickness As Integer)
    Dim edges(0 To 3) As Integer
    edges(0) = 9 'bottom edge
    edges(1) = 8 'top edge
    edges(2) = 7 'left edge
    edges(3) = 10 'right edge
    
    op_val = 1000
    For i = 0 To 3
        If CInt(inp / op_val) <> 0 Then
            'reduce inp for next iteration
            inp = inp Mod op_val
            'draw border
            For Each cell In rng
                With cell.Borders(edges(i))
                    .LineStyle = XlLineStyle.xlContinuous
                    .Weight = thickness
    '               .Color =
                End With
            Next cell
        End If
        'Next operator value
        op_val = op_val / 10
    Next i
End Sub

Private Function ThicknessIndex2Enumeration(ByVal inp As Double)
    Dim outp As Integer
    Select Case CInt(inp)
        Case 0
        outp = 1
        Case 1
        outp = 2
        Case 2
        outp = -4138
        Case 3
        outp = 4
    End Select
    ThicknessIndex2Enumeration = outp
End Function

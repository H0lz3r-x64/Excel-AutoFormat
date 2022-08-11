Attribute VB_Name = "Format"
Dim rowCount As Long, colCount As Long
Dim useColumns As Boolean
Dim useHeaders As Boolean
Dim altRowColours As Boolean
Dim AutoCol As Boolean
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

Public Sub Format(ColumnConf As VisualRow, HeaderConf As VisualRow, BodyConf As VisualRow, GeneralConf As NonVisualRow)
    rowCount = wsInput.UsedRange.Rows(wsInput.UsedRange.Rows.Count).row
    colCount = wsInput.UsedRange.Columns(wsInput.UsedRange.Columns.Count).Column
    
    'Copy over cells
    wsInput.Cells.Copy Destination:=wsOutput.Range("A1")

    'Start Formating
    FormatNonVisual (GeneralConf.Parent.Config)
    'Row
    For row = 1 To rowCount
        'Get range of current row
        Set rng = Range(wsOutput.Cells(row, 1), wsOutput.Cells(row, colCount))
        'Determine type of current row
        If row = 1 And useColumns = True Then
            'Column
            FormatVisual row, rng, ColumnConf.Parent.Config
        ElseIf wsOutput.Cells(row, 1).Text <> "" And useHeaders = True Then
            'Heading
            FormatVisual row, rng, HeaderConf.Parent.Config
        Else
            'Body
            FormatVisual row, rng, BodyConf.Parent.Config
        End If
    Next row
    'Column
    Set rng = Range(wsOutput.Cells(1, 1), wsOutput.Cells(rowCount, 1))
    If AutoCol = True Then
        rng.Columns.AutoFit
    End If
End Sub


Private Sub FormatVisual(ByVal row As Integer, rng As Range, Config As Variant)
    Dim i As Integer
    i = 0
    For Each Item In Config
        Select Case (i)
            Case 8
            'Alternating Rows
            altRowColours = CBool(Item)
            Case 0
            'Bold
            rng.Font.Bold = CBool(Item)
            Case 1
            'Underlined
            rng.Font.Underline = CBool(Item)
            Case 2
            'Italic
            rng.Font.Italic = CBool(Item)
            Case 3
            'WordWrap
            rng.WrapText = CBool(Item)
            Case 4
            'Cell Interior
            rng.Interior.Color = Item
            Case 5
            'Alt. Cell Interior
            If altRowColours = True Then
                If row Mod 2 = 1 Then
                    rng.Interior.Color = Item
                End If
            End If
            Case 6
            'Text Colour
            rng.Font.Color = Item
            Case 7
            'Border Style TODO
            'rng.BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
        End Select
        i = i + 1
    Next Item
End Sub

Private Sub FormatNonVisual(Config As Variant)
    Dim i As Integer
    i = 0
    For Each Item In Config
        Select Case (i)
            Case 0
            'Use Columns
            useColumns = CBool(Item)
            Case 1
            'Use Headers
            useHeaders = CBool(Item)
            Case 2
            'AutoCol
            AutoColumn = CBool(Item)
        End Select
        i = i + 1
    Next Item
End Sub

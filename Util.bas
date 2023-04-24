Attribute VB_Name = "Util"
Public Function Colour2RGB(ByVal colourVal) As Integer()
    Dim r, g, b
    Dim ret(2) As Integer
    r = Hex(colourVal - (colourVal \ 65536) * 65536 - _
        ((colourVal - (colourVal \ 65536) * 65536) \ 256) * 256)
     
    g = Hex((colourVal - (colourVal \ 65536) * 65536) \ 256)
     
    b = Hex(colourVal \ 65536)
     
    If Len(r) < 2 Then r = r & "0"
    If Len(g) < 2 Then g = g & "0"
    If Len(b) < 2 Then b = b & "0"
    
    ret(0) = CDec("&H" & r)
    ret(1) = CDec("&H" & g)
    ret(2) = CDec("&H" & b)
    Colour2RGB = ret
End Function

Public Function msgBoxStrArray(ByVal arr As Variant)
    Dim s As String, i As Double
    s = "Array values returned:" & vbCrLf
    For i = 0 To UBound(arr)
        s = (s & arr(i) & ", ")
    Next
    MsgBox s
End Function

Public Function cvt2StrArr(ByVal arr As Variant) As String()
    ReDim strArr(GetArrLength(arr)) As String
    
    For i = 0 To GetArrLength(arr) - 1
        strArr(i) = CStr(arr(i))
    Next i
    cvt2StrArr = strArr
End Function

Public Function GetArrLength(arr As Variant) As Integer
   If IsEmpty(arr) Then
      GetArrLength = 0
   Else
      GetArrLength = (UBound(arr) - LBound(arr))
   End If
End Function

Function cut_string(ByVal Text As String, StartPos As Integer, EndPos As Integer) As String
    cut_string = Mid(Text, StartPos, EndPos - StartPos + 1)
End Function

Function Count_NonBlank_Cells(ByVal sheet As Worksheet)
    Dim col As Integer, rng As Range, n#, b#
    col = Selection.Column

    Set rng = Intersect(sheet.Columns(1), sheet.UsedRange)
    On Error Resume Next
    b = rng.Cells.SpecialCells(xlCellTypeBlanks).Count
    n = rng.Cells.Count - b
    On Error GoTo 0
    Count_NonBlank_Cells = n
End Function

Function openDialog()
    Dim fd As Office.FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd
        .AllowMultiSelect = False
        
        ' Set the title of the dialog box.
        .Title = "Please select the file."
        
        ' Clear out the current filters, and add our own.
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx, *.xlsm, *.xls"
        
        ' Result
        If .Show = True Then openDialog = .SelectedItems(1)
    End With
End Function

'Function ReGexNumbers()
'    Dim regex As New RegExp
'End Function

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

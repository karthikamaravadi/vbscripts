Sub ExpandIP()
Dim ic As Range, oc As Range, arr As Variant, dic As Object, i As Long, j As Long, c As Long
Dim sloc As Long, v1 As String, v2 As String, v1a As String, s1 As Long, s2 As Long

    Set ic = Range("A2")
    Set oc = Range("B1")
    
    arr = Range(ic, ic.End(xlDown)).Value
    Set dic = CreateObject("Scripting.Dictionary")
    dic.Add 0, "Output"
    
    For i = 1 To UBound(arr)
        sloc = InStr(arr(i, 1), "-")
        If sloc > 0 Then
            v1 = Left(arr(i, 1), sloc - 1)
            v2 = Mid(arr(i, 1), sloc + 1)
        Else
            v1 = arr(i, 1)
            v2 = arr(i, 1)
        End If
        v1a = Left(v1, InStrRev(v1, "."))
        s1 = Mid(v1, InStrRev(v1, ".") + 1)
        s2 = Mid(v2, InStrRev(v2, ".") + 1)
        For j = s1 To s2
            c = c + 1
            dic.Add c, v1a & j
        Next j
    Next i
    
    oc.Resize(dic.Count).Value = WorksheetFunction.Transpose(dic.items)
            
End Sub
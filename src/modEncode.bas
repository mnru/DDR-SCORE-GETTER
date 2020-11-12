Attribute VB_Name = "modEncode"
Function urlEncode(strFrom As String, Optional bSpaceToPlus As Boolean = True, Optional asis As Boolean = False) As String
    Dim num As Long
    Dim char As String
    Dim code As Long
    num = Len(strFrom)
    If asis Then
        urlEncode = strFrom
    ElseIf num = 0 Then
        urlEncode = ""
    Else
        ReDim ret(1 To num)
        For i = 1 To num
            char = Mid(strFrom, i, 1)
            code = Asc(char)
            Select Case code
                Case 48 To 57, 65 To 90, 97 To 122, 45, 46, 95, 126
                    ret(i) = char
                Case 32
                    ret(i) = IIf(bSpaceToPlus, "+", "%20")
                Case 1 To 15
                    ret(i) = "%0" & Hex(code)
                Case Else
                    ret(i) = "%" & Hex(code)
            End Select
        Next i
        urlEncode = Join(ret, "")
    End If
End Function

Function encodeDic(dic, Optional bSpaceToPlus As Boolean = True, Optional asis As Boolean = False)
    Dim ret
    ret = ""
    For Each key In dic.keys
        If TypeName(dic(key)) = "Collection" Then
            For Each elm In dic(key)
                If ret <> "" Then ret = ret & "&"
                ret = ret & urlEncode(CStr(key), bSpaceToPlus, asis) & "=" & urlEncode(CStr(elm), bSpaceToPlus, asis)
            Next elm
        Else
            If ret <> "" Then ret = ret & "&"
            ret = ret & urlEncode(CStr(key), bSpaceToPlus, asis) & "=" & urlEncode(CStr(dic(key)), bSpaceToPlus, asis)
        End If
    Next key
    encodeDic = ret
End Function

Sub testdic()
    Set dic = CreateObject("Scripting.Dictionary")
    Set dic("z") = mkClc(1, 2, "c", "d", "e")
    dic("name") = "abc"
    dic("mail") = "aaa@bbb.com"
    dic("tel") = "(012)345-6789"
    dic("URL") = "https://www.com/a/b/c/"
    dic("x") = "_ a+b-c*d/e=f"
    Set dic("y") = mkClc(" ", "{", "}", "?", "%", "=")
    Data1 = encodeDic(dic)
    Data2 = encodeDic(dic, False)
    Debug.Print Data1
    Debug.Print Data2
End Sub

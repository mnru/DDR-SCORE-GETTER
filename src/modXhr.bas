Attribute VB_Name = "modXhr"
Public xhr

Sub mkXhr()
    Set xhr = CreateObject("msxml2.xmlhttp.6.0")
End Sub

Sub setUrlEncoded()
    Call xhr.setRequestHeader("Content-Type", "application/x-www-form-urlencoded")
End Sub

Sub resBodyToFile(spath)
    Set stm = CreateObject("adodb.stream")
    stm.Open
    stm.Type = 1 'adTypeBinary
    stm.Write xhr.responsebody
    stm.savetofile spath, 2 'adSaveCreateOverWrite
    stm.Close
End Sub

Sub fileToHtml(spath, Optional chr = "shift-jis", Optional bwithmeta = False)
    Set stm = CreateObject("adodb.stream")
    stm.Open
    stm.Type = 2 'adTypeText
    stm.charset = chr
    stm.LoadFromFile spath
    str0 = stm.readtext
    stm.Close
    If bwithmeta Then
        str0 = Replace(str0, "<head>", "<head><meta charset=""" & chr & """>")
    End If
    html.Write str0
End Sub

Attribute VB_Name = "sampleDL"
Function getPageNum(spath)
    Dim ret
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set stm = fso.OpenTextFile(spath)
    src = stm.readall
    Set html = CreateObject("htmlfile")
    html.Write (src)
    stm.Close
    ret = 0
    Set elms = html.getElementsByTagName("div")
    For Each elm In elms
        If elm.className = "page_num" Then
            ret = ret + 1
        End If
    Next elm
    getPageNum = ret
    Set stm = Nothing
    Set html = Nothing
End Function

Sub dlAllScores(Optional rival = "")
    Call dlScores("double", rival)
    Call dlScores("single", rival)
End Sub

Sub dlScores(Optional sd = "double", Optional rival = "")
    Dim t1, t2
    Dim stm
    t1 = Time
    Set fso = CreateObject("scripting.FileSystemObject")
    htmlDir = ThisWorkbook.path & "\html\"
    tsvDir = ThisWorkbook.path & "\tsv\"
    If rival = "" Then
        saveDir = htmlDir
    Else
        saveDir = htmlDir & rival & "\"
    End If
    If Not fso.FolderExists(htmlDir) Then fso.CreateFolder (htmlDir)
    If Not fso.FolderExists(tsvDir) Then fso.CreateFolder (tsvDir)
    If Not fso.FolderExists(saveDir) Then fso.CreateFolder (saveDir)
    If rival = "" Then
        url0 = "https://p.eagate.573.jp/game/ddr/ddra20/p/playdata/music_data_" & sd & ".html"
    Else
        url0 = "https://p.eagate.573.jp/game/ddr/ddra20/p/rival/rival_musicdata_" & sd & ".html"
    End If
    Dim num, i
    i = 1
    Do
        url = url0 & mkOffset(i - 1, rival)
        If rival <> "" Then url = url & "&rival_id=" & rival
        Call xhr.Open("get", url)
        xhr.send
        spath = saveDir & sd & Format(i, "00") & ".html"
        Call resBodyToFile(spath)
        DoEvents
        If i = 1 Then
            num = getPageNum(spath)
        End If
        On Error Resume Next
        frmLogin.llblInfo.Caption = "ダウンロード :" & sd & " " & i & "/" & num
        DoEvents
        On Error GoTo 0
        i = i + 1
        Sleep (3000)
    Loop Until i > num
    t2 = Time
    Debug.Print "dlScores", sd, Format(t2 - t1, "hh:mm:ss")
End Sub

Sub dlMusicData()
    Set fso = CreateObject("Scripting.FileSystemObject")
    url = "https://raw.githubusercontent.com/mnru/DDR-SCORE-GETTER/main/data/musicData.csv"
    dlFdr = ThisWorkbook.path & "\data"
    dlPath = dlFdr & "\musicData.csv"
    If Not fso.FolderExists(dlFdr) Then fso.CreateFolder (dlFdr)
    Call dlUrlToFile(url, dlPath)
End Sub


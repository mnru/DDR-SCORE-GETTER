Attribute VB_Name = "SampleParse"
Sub mkDirs()
    Set fso = CreateObject("Scripting.FileSystemObject")
    htmlDir = ThisWorkbook.path & "\html"
    tsvDir = ThisWorkbook.path & "\tsv"
    For Each fdr In Array(htmlDir, tsvDir)
        If Not fso.FolderExists(fdr) Then fso.CreateFolder (fdr)
    Next fdr
End Sub

Sub htmlsToTsv(sd, htmlFdr, ToFdr, Optional rival = "", Optional bFrm = False)
    Dim html As HTMLDocument
    Dim elm As HTMLObjectElement
    Dim A, imgs, tbl
    Dim stm0, stm1
    Dim fn, sdcnt
    Dim rankAry, comboAry, singleAry, doubleAry
    rankAry = Array("none", "e", "d", "d_p", "c_m", "c", "c_p", "b_m", "b", "b_p", "a_m", "a", "a_p", "aa_m", "aa", "aa_p", "aaa")
    comboAry = Array("none", "good", "great", "perfect", "mar")
    singleAry = Array("ID", "score0", "rank0", "combo0", "score1", "rank1", "combo1", _
    "score2", "rank2", "combo2", "score3", "rank3", "combo3", "score4", "rank4", "combo4", "title")
    doubleAry = Array("ID", "score5", "rank5", "combo5", "score6", "rank6", "combo6", _
    "score7", "rank7", "combo7", "score8", "rank8", "combo8", "title")
    Call mkDirs
    '
    Application.DisplayAlerts = False
    Set fso = CreateObject("Scripting.FileSystemObject")
    If rival = "" Then
        FromFdr = htmlFdr
        toPath = ToFdr & "\" & sd & ".txt"
    Else
        FromFdr = htmlFdr & "\" & rival
        toPath = ToFdr & "\" & rival & "_" & sd & ".txt"
    End If
    '
    Dim adoStm 'As ADODB.Stream
    Set adoStm = CreateObject("adodb.stream")
    adoStm.Type = 1
    adoStm.Open
    Select Case sd
        Case "single"
            Call adoStm.Write(utf8Join(singleAry, vbTab, True))
        Case "double"
            Call adoStm.Write(utf8Join(doubleAry, vbTab, True))
        Case Else
    End Select
    sdcnt = 0
    Set html = CreateObject("htmlfile")
    Set fromFiles = fso.GetFolder(FromFdr).Files
    For Each fromFile In fromFiles
        fn = fromFile.name
        If fn Like sd & "*.html" Then
            sdcnt = sdcnt + 1
            Application.StatusBar = rival & " " & sd & " " & sdcnt
            If bFrm Then
                frmLogin.llblInfo.Caption = "ƒtƒ@ƒCƒ‹•ÏŠ· : " & sd & " " & sdcnt
            End If
            Application.DisplayAlerts = True
            DoEvents
            Application.DisplayAlerts = False
            Set stm0 = fso.OpenTextFile(fromFile)
            txt0 = stm0.readall
            stm0.Close
            html.body.innerHTML = txt0
            stm0.Close
            d = Array(7, 5)
            Set tbl = html.getElementById("data_tbl")
            rNum = tbl.Rows.Length
            cNum = tbl.Rows(0).Cells.Length
            aNum = 2 + 3 * (cNum - 1)
            ReDim ary(0 To aNum - 1)
            For r = 1 To rNum - 1
                For c = 0 To cNum - 1
                    Set elm = tbl.Rows(r).Cells(c)
                    Select Case c
                        Case 0
                            Set A = elm.getElementsByTagName("a")(0)
                            aref = A.href
                            x = Split(aref, "=")(1)
                            y = Split(x, "&")(0)
                            ary(0) = y
                            title = A.innertext
                            ary(aNum - 1) = title
                        Case Else
                            Set DIV = elm.querySelector("div.data_score")
                            score = DIV.innertext
                            ary(3 * c - 2) = score
                            Set imgs = elm.getElementsByTagName("img")
                            For i = 0 To 1
                                fn = fso.GetBaseName(imgs(i).src)
                                txt = Right(fn, Len(fn) - d(i))
                                If i = 0 Then
                                    vl = Application.WorksheetFunction.Match(txt, rankAry, 0) - 1
                                ElseIf i = 1 Then
                                    vl = Application.WorksheetFunction.Match(txt, comboAry, 0) - 1
                                End If
                                ary(3 * c + i - 1) = vl
                            Next i
                    End Select
                Next c
                Call adoStm.Write(utf8Join(ary, vbTab, True))
                DoEvents
            Next r
        End If
        html.Close
        Call adoStm.savetofile(toPath, 2) '1:adSaveCreateNotExist  ,2:adSaveCreateOverWrite
        adoStm.Position = adoStm.Size
        DoEvents
        Application.StatusBar = " "
    Next fromFile
    adoStm.Close
    Application.DisplayAlerts = True
End Sub

Sub bothHtmlToTsv(Optional rival = "")
    Call sdHtmlToTsv("single", rival)
    Call sdHtmlToTsv("double", rival)
End Sub

Sub sdHtmlToTsv(sd, Optional rival = "")
    htmlDir = ThisWorkbook.path & "\html"
    tsvDir = ThisWorkbook.path & "\tsv"
    Call printTimeP("htmlsToTsv", sd, htmlDir, tsvDir, rival, True)
End Sub

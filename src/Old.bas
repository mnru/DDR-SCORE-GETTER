Attribute VB_Name = "Old"
Function mkDiffSQLOld(tblA, tblB, colsSelectA, colsJoinA, colsJoinB)
    Dim sql
    colsSA = eachJoinAry("A.", ",", "", Split(colsSelectA, ","))
    aryJA = Split(colsJoinA, ",")
    aryJB = Split(colsJoinB, ",")
    aryC = Array("A.", " = B.", "")
    aryJ = eachConcateAry(aryC, aryJA, aryJB)
    sql = "Select " & Join(colsSA, ",") & " From " & tblA & " A Left Join " & tblB & " B  On "
    sql = sql & Join(aryJ, " and ") & " Where A." & aryJA(0) & " Is Null"
    mkDiffSQL = sql
End Function

Sub createScoreViewOld()
    sql = "create View ScoreView as " & _
    "select A.ID,A.classID,C.verID,C.initID,ver,init,title,play,deg,lev,score,rank,combo,iif(score<=900000,0,cdbl((score-900000)*lev*2/100000+lev)) as skill " & _
    "from ((((Musiclevel A left join ScoreTbl B on A.ID=B.ID and A.classID=B.classID) " & _
    "left join MusicTbl C on A.ID=C.ID) " & _
    "left join classTbl D on A.classID=D.classID) " & _
    "left join verTbl E on C.verID=E.verID) " & _
    "left join initTbl F on C.initID=F.initID"
    mdbPath = ThisWorkbook.path & "\data.mdb"
    Debug.Print sql
    Call execSQL(sql, mdbPath)
End Sub

Sub testclc()
    Set clc = New Collection
    clc.Add "abc"
    clc(1) = 123
    Stop
End Sub

Sub dropTmpOld()
    tblAry = Array("tmp_s", "tmp_d", "musicData")
    mdbPath = ThisWorkbook.path & "\data.mdb"
    For Each tbl In tblAry
        sql = "Drop Table " & tbl
        Call execSQL(sql, mdbPath)
    Next
End Sub

Sub importTxtOld()
    Dim sdAry(1 To 2)
    sdAry(1) = "single"
    sdAry(2) = "double"
    Dim tblAry(1 To 2)
    tblAry(1) = "tmp_s"
    tblAry(2) = "tmp_d"
    Dim sqls(1 To 2)
    text = mkIniHead(False, "utf8")
    Call mkSchemaIniFile(ThisWorkbook.path & "\tsv\", Array("single.csv", "double.csv"), text)
    For i = 1 To 2
        sqls(i) = "select * into " & tblAry(i) & " from " & txtAsTable(ThisWorkbook.path & "\tsv\" & sdAry(i) & ".csv")
    Next i
    Call execSQLs(sqls, ThisWorkbook.path & "\data.mdb")
End Sub

Sub tmpToScoreTblOld()
    Set fso = CreateObject("Scripting.FileSystemObject")
    Call fso.CopyFile(ThisWorkbook.path & "\data\schemaCSV.ini", ThisWorkbook.path & "\tsv\schema.ini")
    tblD = txtAsTable(ThisWorkbook.path & "\tsv\double.csv")
    tblS = txtAsTable(ThisWorkbook.path & "\tsv\single.csv")
    Dim sSQLs(1 To 9)
    'sSQLs(1) = mkSelectIntoSQL("ScoreTbl", tblD, "ID,classID,score,rank,combo", "F1,5,F3,F4,F5", "F4<>'none'")
    sSQLs(1) = mkInsertIntoSQL("ScoreTbl", tblD, "ID,classID,score,rank,combo", "F1,5,F3,F4,F5", "F4<>'none'") '5 db
    sSQLs(2) = mkInsertIntoSQL("ScoreTbl", tblD, "ID,classID,score,rank,combo", "F1,6,F6,F7,F8", "F7<>'none'") '6 dd
    sSQLs(3) = mkInsertIntoSQL("ScoreTbl", tblD, "ID,classID,score,rank,combo", "F1,7,F9,F10,F11", "F10<>'none'") '7 de
    sSQLs(4) = mkInsertIntoSQL("ScoreTbl", tblD, "ID,classID,score,rank,combo", "F1,8,F12,F13,F14", "F13<>'none'") '8 dc
    sSQLs(5) = mkInsertIntoSQL("ScoreTbl", tblS, "ID,classID,score,rank,combo", "F1,0,F3,F4,F5", "F4<>'none'") '0 sg
    sSQLs(6) = mkInsertIntoSQL("ScoreTbl", tblS, "ID,classID,score,rank,combo", "F1,1,F6,F7,F8", "F7<>'none'") '1 sb
    sSQLs(7) = mkInsertIntoSQL("ScoreTbl", tblS, "ID,classID,score,rank,combo", "F1,2,F9,F10,F11", "F10<>'none'") '2 sd
    sSQLs(8) = mkInsertIntoSQL("ScoreTbl", tblS, "ID,classID,score,rank,combo", "F1,3,F12,F13,F14", "F13<>'none'") '3 se
    sSQLs(9) = mkInsertIntoSQL("ScoreTbl", tblS, "ID,classID,score,rank,combo", "F1,4,F15,F16,F17", "F16<>'none'") '4 sc
    Call execSQLs(sSQLs, ThisWorkbook.path & "\data.mdb")
End Sub

Sub importTxtOld2()
    Dim sdAry(1 To 2)
    sdAry(1) = "single"
    sdAry(2) = "double"
    Dim tblAry(1 To 2)
    tblAry(1) = "tmp_s"
    tblAry(2) = "tmp_d"
    Dim sqls(1 To 2)
    Set fso = CreateObject("Scripting.FileSystemObject")
    Call fso.CopyFile(ThisWorkbook.path & "\data\schemaCSV.ini", ThisWorkbook.path & "\tsv\schema.ini")
    For i = 1 To 2
        sqls(i) = mkInsertIntoSQL(tblAry(i), txtAsTable(ThisWorkbook.path & "\tsv\" & sdAry(i) & ".csv"))
    Next i
    Call execSQLs(sqls, ThisWorkbook.path & "\data.mdb")
End Sub

Sub tmpToScoreTblOld2(Optional toTbl = "ScoreTbl")
    Dim tmpTbl
    Set fso = CreateObject("Scripting.FileSystemObject")
    '    Call fso.CopyFile(ThisWorkbook.path & "\data\schemaCSV.ini", ThisWorkbook.path & "\tsv\schema.ini")
    '    tblD = txtAsTable(ThisWorkbook.path & "\tsv\double.csv")
    '    tblS = txtAsTable(ThisWorkbook.path & "\tsv\single.csv")
    Dim sSQLs(0 To 8)
    For i = 0 To 8
        tmpTbl = IIf(i <= 4, "tmp_s", "tmp_d")
        sSQLs(i) = mkInsertIntoSQL(toTbl, tmpTbl, "ID,classID,score,rankID,comboID", _
        Join(Array("id", i, "score" & i, "rank" & i, "combo" & i), ","), "rank" & i & ">0")
    Next i
    Call execSQLs(sSQLs, ThisWorkbook.path & "\data.mdb")
End Sub

Sub updateTblOld2()
    Dim sSQLs(0 To 2)
    sSQLs(0) = "update scoretbl set rank=replace(rank,'_m','-')"
    sSQLs(1) = "update scoretbl set rank=replace(rank,'_p','+')"
    sSQLs(2) = "update scoreTbl set rank=Ucase(rank)"
    Call execSQLs(sSQLs, ThisWorkbook.path & "\data.mdb")
End Sub

Sub dropTmpOld1()
    mdbPath = ThisWorkbook.path & "\data.mdb"
    sql = "Drop Table musicData"
    Call execSQL(sql, mdbPath)
End Sub

Sub dropMusicsOld()
    mdbPath = ThisWorkbook.path & "\data.mdb"
    ary = Array("musicData", "music1", "music2")
    On Error Resume Next
    For Each elm In ary
        Call execSQL("Drop Table " & elm, mdbPath)
    Next
    On Error GoTo 0
End Sub

Sub importMusicDataOld()
    Dim sql, tbl
    tbl = txtAsTable(ThisWorkbook.path & "\Data\musicData.csv")
    sql = mkSelectIntoSQL("musicData", tbl)
    Call execSQL(sql, ThisWorkbook.path & "\data.mdb")
End Sub

Sub dropTmpDataOld()
    mdbPath = ThisWorkbook.path & "\data.mdb"
    sql = "Drop Table musicData"
    Call execSQL(sql, mdbPath)
End Sub

Sub changeOldCsv()
    Set fso = CreateObject("Scripting.FileSystemObject")
    pn = Application.GetOpenFilename("csv file,*.csv")
    If pn = False Then Exit Sub
    fdrn = fso.GetParentFolderName(pn)
    sd = LCase(fso.GetBaseName(pn))
    Dim infoAry, titleAry, repNum
    Select Case sd
        Case "double"
            infoAry = Array(Array(1, 1), _
            Array(2, 2), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), _
            Array(9, 1), Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1), Array(14, 1))
            titleAry = Array("ID", "title", "score5", "rank5", "combo5", "score6", "rank6", "combo6", _
            "score7", "rank7", "combo7", "score8", "rank8", "combo8")
            repNum = 4
        Case "single"
            infoAry = Array(Array(1, 1), _
            Array(2, 2), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), _
            Array(9, 1), Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1), Array(14, 1), _
            Array(15, 1), Array(16, 1), Array(17, 1))
            titleAry = Array("ID", "title", "score0", "rank0", "combo0", "score1", "rank1", "combo1", _
            "score2", "rank2", "combo2", "score3", "rank3", "combo3", "score4", "rank4", "combo4")
            repNum = 5
        Case Else
    End Select
    Workbooks.OpenText Filename:=pn, Origin _
    :=65001, StartRow:=1, DataType:=xlDelimited, TextQualifier:= _
    xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, _
    Comma:=True, Space:=False, Other:=False, FieldInfo:=infoAry
    Rows(1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Call layAryAt(titleAry, 1, 1)
    Dim rankAry, comboAry
    rankAry = Array("none", "e", "d", "d_p", "c_m", "c", "c_p", "b_m", "b", "b_p", "a_m", "a", "a_p", "aa_m", "aa", "aa_p", "aaa")
    comboAry = Array("none", "good", "great", "perfect", "mar")
    addrow = 2
    rNum = Cells(1, 1).CurrentRegion.Rows.Count
    c = 4
    For rep = 1 To repNum
        For r = 2 To rNum
            txt1 = Cells(r, c).Value
            txt2 = Cells(r, c + 1).Value
            Cells(r, c) = Application.WorksheetFunction.Match(txt1, rankAry, 0) - 1
            Cells(r, c + 1) = Application.WorksheetFunction.Match(txt2, comboAry, 0) - 1
        Next r
        c = c + 3
    Next rep
    ActiveSheet.Copy
    Call ActiveWorkbook.SaveAs(fdrn & "\" & sd & "0.csv", xlCSVUTF8)
End Sub

Sub testold()
    fdrn = ThisWorkbook.path & "\html\"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fdr = fso.GetFolder(fdrn)
    For Each fl In fdr.Files
        x = fl.path
        y = fl.name
        If Len(fl.name) > 13 Then
            fl.name = Right(fl.name, Len(fl.name) - 9)
        End If
        'Stop
    Next
End Sub

Function xlTblAsTableOld(xlTbl, Optional bkn = "", Optional bPlusOneAbove = True, Optional bHdr = True)
    Dim xlsPath, adn, shn
    Dim rs, cs
    If bkn = "" Then bkn = ThisWorkbook.name
    Set fso = CreateObject("Scripting.FileSystemObject")
    Workbooks(bkn).Activate
    xlsPath = Workbooks(bkn).FullName
    shn = Range(xlTbl).Parent.name
    Sheets(shn).Activate
    xlTbl = Trim(xlTbl)
    Set rg = Range(xlTbl)
    rs = rg.Rows.Count
    cs = rg.Columns.Count
    If bPlusOneAbove Then
        adn = rg.offset(-1, 0).Resize(rs + 1, cs).Address(False, False)
    Else
        adn = rg.Address(False, False)
    End If
    If bAddTitle Then
        num = InStr(xlTbl, "[")
        If num = 0 Then
            xlTbl = Trim(xlTbl) & "[#All]"
        Else
            c = Mid(xlTbl, num + 1, 1)
            If c <> "#" Then
                If c = "[" Then
                    If Mid(xlTbl, num + 2, 1) <> "#" Then
                        xlTbl = Left(xlTbl, num) & "[#All],[" & Right(xlTbl, Len(xlTbl) - num - 1)
                    End If
                Else
                    xlTbl = Left(xlTbl, num) & "[#All],[" & Right(xlTbl, Len(xlTbl) - num) & "]"
                End If
            End If
        End If
    End If
    adn = Range(xlTbl).Address(False, False)
    xlTblAsTable = xlAsTable(xlsPath, shn, adn, bHdr)
End Function

Sub writeSkillOld(code, pwd, id, Optional sp = True, Optional dp = True, Optional mdbPath = "")
    cdpw = "&ddrcode=" & code & "&password=" & pwd
    Call mkXhr
    Call getCurMdb(mdbPath)
    url = "http://skillattack.com/sa4/dancer_profile.php?ddrcode=" & code
    Call xhr.Open("get", url)
    xhr.send
    '
    postdata = "_=" & cdpw
    Call xhr.Open("post", url)
    Call setUrlEncoded
    xhr.send (postdata)
    Call resBodyToFile(ThisWorkbook.path & "\tmp1.html")
    '
    url = " http://skillattack.com/sa4/dancer_input.php"
    Call xhr.Open("post", url)
    Call setUrlEncoded
    xhr.send (postdata)
    Call resBodyToFile(ThisWorkbook.path & "\tmp2.html")
    '
    url = " http://skillattack.com/sa4/dancer_input.php"
    Call xhr.Open("post", url)
    Call setUrlEncoded
    xhr.send (postdata)
    verID = getSqlVals("select verID from MusicTbl where id='" & id & "'")(0)
    ver = TLookup(verID, "verTbl", "val")
    postdata = "_=score" & cdpw & "&series=" & ver & "&initial=&music_data="
    Call xhr.Open("post", url)
    Call setUrlEncoded
    xhr.send (postdata)
    Call resBodyToFile(ThisWorkbook.path & "\tmp3.html")
    '
    url = " http://skillattack.com/sa4/dancer_input.php"
    Call xhr.Open("post", url)
    Call setUrlEncoded
    postdata = "_=score_submit" & "&ddrcode=" & code & "&password=" & pwd
    Call getCurMdb(mdbPath)
    postdata = postdata & "&" & getPostScoreData(id, sp, dp)
    Call xhr.Open("post", url)
    xhr.send (postdata)
    Call resBodyToFile(ThisWorkbook.path & "\tmp4.html")
End Sub

Sub htmlsToCsv0Old(sd, FromFdr, ToFdr)
    Dim html As HTMLDocument
    Dim elm As HTMLDTElement
    Dim fn, sdcnt
    Dim rankAry, comboAry, singleAry, doubleAry
    Dim A
    '
    rankAry = Array("none", "e", "d", "d_p", "c_m", "c", "c_p", "b_m", "b", "b_p", "a_m", "a", "a_p", "aa_m", "aa", "aa_p", "aaa")
    comboAry = Array("none", "good", "great", "perfect", "mar")
    singleAry = Array("ID", "title", "score0", "rank0", "combo0", "score1", "rank1", "combo1", _
    "score2", "rank2", "combo2", "score3", "rank3", "combo3", "score4", "rank4", "combo4")
    doubleAry = Array("ID", "title", "score5", "rank5", "combo5", "score6", "rank6", "combo6", _
    "score7", "rank7", "combo7", "score8", "rank8", "combo8")
    '
    Application.DisplayAlerts = False
    Set fso = CreateObject("Scripting.FileSystemObject")
    toPath = ToFdr & "\" & sd & ".csv"
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets.Add After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    sn = ActiveSheet.name
    ThisWorkbook.Sheets(sn).Columns(2).NumberFormat = "@"
    Select Case sd
        Case "single"
            Call layAryAt(singleAry, 1, 1)
        Case "double"
            Call layAryAt(doubleAry, 1, 1)
        Case Else
    End Select
    addrow = 2
    sdcnt = 0
    Set html = CreateObject("htmlfile")
    Set fromFiles = fso.GetFolder(FromFdr).Files
    For Each fromFile In fromFiles
        fn = fromFile.name
        If fn Like sd & "*.html" Then
            sdcnt = sdcnt + 1
            Application.StatusBar = sd & " " & sdcnt
            Application.DisplayAlerts = True
            DoEvents
            Application.DisplayAlerts = False
            Set stm0 = fso.OpenTextFile(fromFile)
            txt0 = stm0.readall
            stm0.Close
            html.Write (txt0)
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
                            ary(0) = Right(aref, Len(aref) - InStr(aref, "="))
                            ary(1) = A.innertext
                        Case Else
                            Set divs = elm.getElementsByTagName("div")
                            score = ""
                            For Each DIV In divs
                                If Not IsNull(DIV.className) Then
                                    If DIV.className = "data_score" Then
                                        score = DIV.innerHTML
                                        Exit For
                                    End If
                                End If
                            Next DIV
                            ary(3 * c - 1) = score
                            Set imgs = elm.getElementsByTagName("img")
                            For i = 0 To 1
                                fn = fso.GetBaseName(imgs(i).src)
                                txt = Right(fn, Len(fn) - d(i))
                                If i = 0 Then
                                    vl = Application.WorksheetFunction.Match(txt, rankAry, 0) - 1
                                ElseIf i = 1 Then
                                    vl = Application.WorksheetFunction.Match(txt, comboAry, 0) - 1
                                End If
                                ary(3 * c + i) = vl
                            Next i
                    End Select
                Next c
                ThisWorkbook.Sheets(sn).Cells(addrow, 1).Resize(1, aNum) = ary
                addrow = addrow + 1
            Next r
        End If
        html.Close
        DoEvents
        Application.StatusBar = " "
        ThisWorkbook.Sheets(sn).Cells(addrow, 1).Select
    Next fromFile
    ThisWorkbook.Sheets(sn).Move
    Call ActiveWorkbook.SaveAs(toPath, xlCSVUTF8)
    ActiveWorkbook.Close
    Application.DisplayAlerts = True
End Sub

Sub htmlsToCsvOld(sd, htmlFdr, ToFdr, Optional rival = "")
    Dim html
    Dim fn, sdcnt
    Dim rankAry, comboAry, singleAry, doubleAry
    Set fso = CreateObject("Scripting.FileSystemObject")
    toPath = ToFdr & "\" & sd & ".csv"
    If rival = "" Then
        FromFdr = htmlFdr
        toPath = ToFdr & "\" & sd & ".csv"
    Else
        FromFdr = htmlFdr & "\" & rival
        toPath = ToFdr & "\" & rival & "_" & sd & ".csv"
    End If
    '
    rankAry = Array("none", "e", "d", "d_p", "c_m", "c", "c_p", "b_m", "b", "b_p", "a_m", "a", "a_p", "aa_m", "aa", "aa_p", "aaa")
    comboAry = Array("none", "good", "great", "perfect", "mar")
    singleAry = Array("ID", "title", "score0", "rank0", "combo0", "score1", "rank1", "combo1", _
    "score2", "rank2", "combo2", "score3", "rank3", "combo3", "score4", "rank4", "combo4")
    doubleAry = Array("ID", "title", "score5", "rank5", "combo5", "score6", "rank6", "combo6", _
    "score7", "rank7", "combo7", "score8", "rank8", "combo8")
    '
    Application.DisplayAlerts = False
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets.Add After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    sn = ActiveSheet.name
    ThisWorkbook.Sheets(sn).Columns(2).NumberFormat = "@"
    Select Case sd
        Case "single"
            Call layAryAt(singleAry, 1, 1)
        Case "double"
            Call layAryAt(doubleAry, 1, 1)
        Case Else
    End Select
    addrow = 2
    sdcnt = 0
    Set html = CreateObject("htmlfile")
    Set fromFiles = fso.GetFolder(FromFdr).Files
    For Each fromFile In fromFiles
        fn = fromFile.name
        If fn Like sd & "*.html" Then
            sdcnt = sdcnt + 1
            Application.StatusBar = sd & " " & sdcnt
            Application.DisplayAlerts = True
            DoEvents
            Application.DisplayAlerts = False
            Set stm0 = fso.OpenTextFile(fromFile)
            txt0 = stm0.readall
            stm0.Close
            html.Write (txt0)
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
                            ary(0) = Right(aref, Len(aref) - InStr(aref, "="))
                            ary(1) = A.innertext
                        Case Else
                            Set divs = elm.getElementsByTagName("div")
                            score = ""
                            For Each DIV In divs
                                If Not IsNull(DIV.className) Then
                                    If DIV.className = "data_score" Then
                                        score = DIV.innerHTML
                                        Exit For
                                    End If
                                End If
                            Next DIV
                            ary(3 * c - 1) = score
                            Set imgs = elm.getElementsByTagName("img")
                            For i = 0 To 1
                                fn = fso.GetBaseName(imgs(i).src)
                                txt = Right(fn, Len(fn) - d(i))
                                If i = 0 Then
                                    vl = Application.WorksheetFunction.Match(txt, rankAry, 0) - 1
                                ElseIf i = 1 Then
                                    vl = Application.WorksheetFunction.Match(txt, comboAry, 0) - 1
                                End If
                                ary(3 * c + i) = vl
                            Next i
                    End Select
                Next c
                ThisWorkbook.Sheets(sn).Cells(addrow, 1).Resize(1, aNum) = ary
                addrow = addrow + 1
            Next r
        End If
        html.Close
        DoEvents
        Application.StatusBar = " "
        ThisWorkbook.Sheets(sn).Cells(addrow, 1).Select
    Next fromFile
    ThisWorkbook.Sheets(sn).Move
    Call ActiveWorkbook.SaveAs(toPath, xlCSVUTF8)
    ActiveWorkbook.Close
    Application.DisplayAlerts = True
End Sub

Sub bothHtmlToCsvOld(Optional rival = "")
    htmlDir = ThisWorkbook.path & "\html"
    csvDir = ThisWorkbook.path & "\csv"
    Call printTimeP("htmlsToCsv", "single", htmlDir, csvDir, rival)
    Call printTimeP("htmlsToCsv", "double", htmlDir, csvDir, rival)
    If rival = "" Then
        setSchema
        importTxt
        updateScoreTbl
    Else
        setSchema
        importRivalTxt (rival)
        updateRivalScoreTbl (rival)
    End If
    MsgBox "èIóπÇµÇ‹ÇµÇΩ"
End Sub

Sub allTsvToScoreDBOld()
    rivalAry = mapA("getfilepart", getFileAry(ThisWorkbook.path & "\html", "folder"), "base")
    Call tsvToScoreDB
    For Each rival In rivalAry
        tsvToScoreDB (rival)
    Next
End Sub

Sub appendUTF8AryToFileOld(adoStm, ary, dlm)
    Dim stm0, stm1, ret
    Set stm0 = CreateObject("ADODB.Stream")
    stm0.Type = 2 '1:adTypeBinary,2:adTypeText
    stm0.Open
    lb = LBound(ary)
    ub = UBound(ary)
    For i = lb To ub
        Call stm0.WriteText(ary(i), 0) '0:adWriteChar,1:adWriteLine
        If i < ub Then
            Call stm0.WriteText(dlm, 0)
        End If
    Next i
    stm0.Position = 0
    stm0.Type = 1 'adTypeBinary
    stm0.Position = 3
    stm0.CopyTo (adoStm)
    stm0.Close
    adoStm.Position = adoStm.Size
End Sub

Function getPostScoreDatasOld(id, Optional sp = True, Optional dp = True, Optional mdbPath = "")
    Dim ret
    Dim scoreAry
    clsAry = Array("gsp", "bsp", "dsp", "esp", "csp", "bdp", "ddp", "edp", "cdp")
    Call getCurMdb(mdbPath)
    ReDim scoreAry(0 To 8)
    num = getSqlVals("select num from musicTbl where id='" & id & "'")(0)
    For i = 0 To 8
        sql = "select score,comboID from scoreTbl where id='" & id & "' and classID=" & i
        tmp = getSqlVals(sql, 2)
        If tmp(0) <> "" Then
            If tmp(1) = 3 Then
                scoreAry(i) = tmp(0) & "%2A%2A"
            ElseIf tmp(1) = 2 Or tmp(1) = 1 Then
                scoreAry(i) = tmp(0) & "%2A"
            Else
                scoreAry(i) = tmp(0)
            End If
        End If
    Next i
    ret = "index%5B%5D=" & num
    For i = 0 To 4
        ret = ret & "&" & clsAry(i) & "%5B%5D=" & IIf(sp, scoreAry(i), "")
    Next i
    For i = 5 To 8
        ret = ret & "&" & clsAry(i) & "%5B%5D=" & IIf(dp, scoreAry(i), "")
    Next i
    getPostScoreDatas = ret
End Function

Function getPostScoreDataOld(id, classId, Optional sp = True, Optional dp = True, Optional mdbPath = "")
    Dim ret, play
    Dim score
    clsAry = Array("gsp", "bsp", "dsp", "esp", "csp", "bdp", "ddp", "edp", "cdp")
    Call getCurMdb(mdbPath)
    num = getSqlVals("select num from musicTbl where id='" & id & "'")(0)
    sql = "select score,comboID from scoreTbl where id='" & id & "' and classID=" & classId
    tmp = getSqlVals(sql, 2)
    If tmp(0) <> "" Then
        If tmp(1) = 3 Then
            score = tmp(0) & "%2A%2A"
        ElseIf tmp(1) = 2 Then
            score = tmp(0) & "%2A"
        Else
            score = tmp(0)
        End If
    End If
    play = IIf(i <= 4, sp, dp)
    ret = "index%5B%5D=" & num & "&" & clsAry(classId) & "%5B%5D=" & IIf(play, score, "")
    getPostScoreData = ret
End Function

Sub writeSkillDataOld(code, pwd, id, Optional sp = True, Optional dp = True, Optional mdbPath = "")
    Call mkXhr
    Call getCurMdb(mdbPath)
    '
    url = " http://skillattack.com/sa4/dancer_input.php"
    postdata = "_=" & "&ddrcode=" & code & "&password=" & pwd
    Call xhr.Open("post", url)
    Call setUrlEncoded
    xhr.send (postdata)
    '
    url = " http://skillattack.com/sa4/dancer_input.php"
    Call xhr.Open("post", url)
    Call setUrlEncoded
    postdata = "_=score_submit" & "&ddrcode=" & code & "&password=" & pwd
    Call getCurMdb(mdbPath)
    postdata = postdata & "&" & getPostScoreDatas(id, sp, dp)
    Call xhr.Open("post", url)
    Call setUrlEncoded
    xhr.send (postdata)
End Sub

Sub importTxt0Old(Optional rival = "", Optional bDirect = True, Optional tsvFdr = "")
    If tsvFdr = "" Then tsvFdr = ThisWorkbook.path & "\tsv\"
    toTbl = IIf(Not bDirect, "tmp", IIf(rival = "", "scoreTbl", "rivalScoreTbl"))
    If Right(Trim(tsvFdr), 1) <> "\" Then tsvFdr = Trim(tsvFdr) & "\"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim sSQLs(0 To 9)
    sSQLs(0) = "delete from " & toTbl
    If rival <> "" Then sqls(0) = sqls(0) & " where rivalID=" & rival
    For i = 0 To 8
        sd = IIf(i <= 4, "single", "double")
        If rival = "" Then
            fromTbl = txtAsTable(tsvFdr & sd & ".txt")
        Else
            fromTbl = txtAsTable(tsvFdr & rival & "_" & sd & ".txt")
        End If
        '
        sSQLs(i + 1) = mkInsertIntoSQL(toTbl, fromTbl, "ID,classID,score,rankID,comboID", _
        Join(Array("id", i, "score" & i, "rank" & i, "combo" & i), ","), "rank" & i & "<16")
    Next i
    Call execSQLs(sSQLs, ThisWorkbook.path & "\data.mdb")
End Sub

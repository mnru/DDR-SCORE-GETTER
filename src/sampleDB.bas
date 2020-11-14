Attribute VB_Name = "sampleDB"
Option Base 0

Sub mkScoreDB()
    mdbPath = getCurMdb
    mkMdb (mdbPath)
    mkScoreTbls
    importMasterData
    importMusicData
    createScoreView
    Sheets("menu").Activate
    MsgBox "終了しました"
End Sub

Sub htmlToScoreDB(Optional rival = "")
    If rival = "" Then
        bothHtmlToCsv
        setSchema
        importTxt
        updateScoreTbl
    Else
        bothHtmlToCsv (rival)
        setSchema
        importRivalTxt (rival)
        updateRivalScoreTbl (rival)
    End If
    MsgBox "終了しました"
End Sub

Sub allHtmlToScoreDB()
    Call allHtmlToTsv
    Call allTsvToScoreDB
End Sub

Sub allHtmlToTsv()
    rivalAry = mapA("getfilepart", getFileAry(ThisWorkbook.path & "\html", "folder"), "base")
    Call bothHtmlToTsv
    For Each rival In rivalAry
        bothHtmlToTsv (rival)
    Next
End Sub

Sub allTsvToScoreDB()
    Application.StatusBar = "import my score"
    Call printTime("tsvToScoreDB")
    DoEvents
    Call allTsvToRivalScoreDB
    Application.StatusBar = ""
    MsgBox "終了しました"
End Sub

Sub allTsvToRivalScoreDB()
    rivalAry = mapA("getfilepart", getFileAry(ThisWorkbook.path & "\html", "folder"), "base")
    For Each rival In rivalAry
        Application.StatusBar = "import rival score : " & rival
        Call printTimeP("tsvToRivalScoreDB", rival)
        DoEvents
    Next
    Application.StatusBar = ""
End Sub

Sub importTsvToScoreDB(Optional rival = "")
    frmLogin.llblInfo.Caption = "インポート : " & sd & " " & sdcnt
    If rival = "" Then
        Call printTime("tsvToScoreDB")
    Else
        Call printTimeP("tsvToRivalScoreDB", rival)
    End If
    frmLogin.llblInfo.Caption = "終了しました"
End Sub

Sub tsvToScoreDB()
    Dim num
    setSchema
    num = getSqlVals("select count(*) from ScoreTbl")(0)
    If num = 0 Then
        importTxt ("ScoreTbl")
        setSkill
    Else
        importTxt
        updateScoreTbl
    End If
End Sub

Sub tsvToRivalScoreDB(rival)
    Dim num
    setSchema
    num = getSqlVals("select count(*) from rivalScoreTbl where rivalId=" & rival)(0)
    If num = 0 Then
        Call importRivalTxt(rival, "rivalScoreTbl")
        Call setRivalSkill(rival)
    Else
        Call importTxt(, rival)
        updateRivalScoreTbl (rival)
    End If
End Sub

Sub mkScoreTbls()
    Call getCurMdb
    Call mkTable("tmp")
    Call mkTable("ScoreTbl")
    Call mkTable("MusicLevel")
    Call mkTable("MusicTbl")
    Call mkTable("ClassTbl")
    Call mkTable("verTbl")
    Call mkTable("initTbl")
    Call mkTable("rankTbl")
    Call mkTable("comboTbl")
    Call mkTable("updateTbl")
    Call mkTable("rivalScoreTbl")
    Call mkTable("previousScore")
    Call mkTable("rivalPreviousScore")
End Sub

Sub execDelete(tbl, Optional mdbPath = "")
    If mdbPath = "" Then mdbPath = ThisWorkbook.path & "\data.mdb"
    sql = "delete from " & tbl
    Call execSQL(sql, mdbPath)
End Sub

Sub setSchema()
    Set fso = CreateObject("Scripting.FileSystemObject")
    tsvDir = ThisWorkbook.path & "\tsv"
    '  If notfso.FolderExists(tsvDir) Then fso.CreateFolder (tsvDir)
    schemaPath = tsvDir & "\schema.ini"
    Set stm = fso.CreateTextFile(schemaPath)
    pns = getFileAry(tsvDir)
    For Each pn In pns
        fn = fso.getfilename(pn)
        If fn Like "*single.txt" Then
            stm.WriteLine ("[" & fn & "]")
            stm.WriteLine (TLookup("single", "schemaDef", "def"))
        ElseIf fn Like "*double.txt" Then
            stm.WriteLine ("[" & fn & "]")
            stm.WriteLine (TLookup("double", "schemaDef", "def"))
        End If
    Next pn
End Sub

Sub importTxt(Optional toTbl = "tmp", Optional rival = "", Optional tsvFdr = "")
    If tsvFdr = "" Then tsvFdr = ThisWorkbook.path & "\tsv\"
    If Right(Trim(tsvFdr), 1) <> "\" Then tsvFdr = Trim(tsvFdr) & "\"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim sSQLs(0 To 9)
    sSQLs(0) = "delete from " & toTbl
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

Sub importRivalTxt(rival, Optional toTbl = "rivalScoreTbl", Optional tsvFdr = "")
    If tsvFdr = "" Then tsvFdr = ThisWorkbook.path & "\tsv\"
    If Right(Trim(tsvFdr), 1) <> "\" Then tsvFdr = Trim(tsvFdr) & "\"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim sSQLs(0 To 8)
    For i = 0 To 8
        sd = IIf(i <= 4, "single", "double")
        fromTbl = txtAsTable(tsvFdr & rival & "_" & sd & ".txt")
        '
        sSQLs(i) = mkInsertIntoSQL(toTbl, fromTbl, "rivalID,ID,classID,score,rankID,comboID", _
        Join(Array(rival, "id", i, "score" & i, "rank" & i, "combo" & i), ","), "rank" & i & "<16")
    Next i
    Call execSQLs(sSQLs, ThisWorkbook.path & "\data.mdb")
End Sub

Sub setTestForimportTxt()
    Call importTxt("scoreTbl", ThisWorkbook.path & "\tmp")
    Call importTxt("scoreHistory", ThisWorkbook.path & "\tmp")
    Call importTxt("tmp")
End Sub

Sub updateScoreTbl()
    Dim sqls1(1 To 5), sqls2(1 To 5)
    Dim num
    sqls1(1) = "update (select * from scoreTbl A right join tmp B on A.ID=B.ID and A.classID=B.classID where A.ID is null) set B.updateFlg=1"
    sqls1(2) = "update (select * from scoreTbl A inner join tmp B on A.ID=B.ID and A.classID=B.classID where A.score<B.score) set B.updateFlg=2"
    sqls1(3) = "update (select * from scoreTbl A inner join tmp B on A.ID=B.ID and A.classID=B.classID where A.score>=B.score and A.comboID>B.comboID) set B.updateFlg=3"
    sqls1(4) = "update (select * from MusicLevel A inner join tmp B on A.ID=B.ID and A.classID=B.ClassID where updateFlg>0) " & _
    "Set skill=iif(score<=900000,0,cdbl((score-900000)*lev*2/100000+lev))"
    sqls1(5) = "update tmp set skill=int(skill*100)/100.0"
    sqls2(1) = "delete from previousScore"
    sqls2(2) = "insert into previousScore(updateFlg,ID,classID,score,rankID,comboID) select updateFlg,ID,classID,score,rankID,comboID from tmp where updateFlg=1"
    sqls2(3) = "insert into previousScore(updateFlg,ID,classID,score,rankID,comboID,previousScore,previousRankID,previousComboID) " & _
    "select B.updateFlg,B.ID,B.classID,B.score,B.rankID,B.comboID,A.score,A.rankID,A.comboID from  ScoreTbl A inner join tmp B on A.ID=B.ID and A.classID=B.classID where B.updateFlg>=2"
    sqls2(4) = "delete from scoreTbl A where exists(select * from tmp B where A.ID=B.ID and A.classID=B.classID and B.updateFlg>=2) "
    sqls2(5) = "insert into scoreTbl(ID,classID,score,rankID,comboID,skill) select ID,classID,score,rankID,comboID,skill from tmp B where B.updateFlg >0 "
    Call getCurMdb
    Call execSQLs(sqls1)
    num = getSqlVals("select count(*) from tmp where updateFlg>0")(0)
    If num > 0 Then
        Call execSQLs(sqls2)
    End If
End Sub

Sub updateRivalScoreTbl(rival)
    Dim sqls1(1 To 5), sqls2(1 To 5)
    Dim num
    sqls1(1) = "update (select * from rivalScoreTbl A right join tmp B on A.ID=B.ID and A.classID=B.classID where A.rivalID=" & rival & " and A.ID is null) set B.updateFlg=1"
    sqls1(2) = "update (select * from rivalScoreTbl A inner join tmp B on A.ID=B.ID and A.classID=B.classID where A.rivalID=" & rival & " and  A.score<B.score) set B.updateFlg=2"
    sqls1(3) = "update (select * from rivalScoreTbl A inner join tmp B on A.ID=B.ID and A.classID=B.classID where A.rivalID=" & rival & " and  A.score>=B.score and A.comboID>B.comboID) set B.updateFlg=3"
    sqls1(4) = "update (select * from MusicLevel A inner join tmp B on A.ID=B.ID and A.classID=B.ClassID where updateFlg>0) " & _
    "Set skill=iif(score<=900000,0,cdbl((score-900000)*lev*2/100000+lev))"
    sqls1(5) = "update tmp set skill=int(skill*100)/100.0"
    sqls2(1) = "delete from rivalPreviousScore where rivalID=" & rival
    sqls2(2) = "insert into rivalPreviousScore(rivalID,updateFlg,ID,classID,score,rankID,comboID) select " & rival & ",updateFlg,ID,classID,score,rankID,comboID from tmp where updateFlg=1"
    sqls2(3) = "insert into rivalPreviousScore(rivalID,updateFlg,ID,classID,score,rankID,comboID,previousScore,previousRankID,previousComboID) " & _
    "select " & rival & ",B.updateFlg,B.ID,B.classID,B.score,B.rankID,B.comboID,A.score,A.rankID,A.comboID from  rivalScoreTbl A inner join tmp B on A.ID=B.ID and A.classID=B.classID where B.updateFlg>=2 and A.rivalID=" & rival
    sqls2(4) = "delete from rivalScoreTbl A where exists(select * from tmp B where A.rivalID=" & rival & " and A.ID=B.ID and A.classID=B.classID and B.updateFlg>=2) "
    sqls2(5) = "insert into rivalScoreTbl(rivalID,ID,classID,score,rankID,comboID,skill) select " & rival & ",ID,classID,score,rankID,comboID,skill from tmp B where B.updateFlg >0 "
    Call getCurMdb
    Call execSQLs(sqls1)
    num = getSqlVals("select count(*) from tmp where updateFlg>0")(0)
    If num > 0 Then
        Call execSQLs(sqls2)
    End If
End Sub

Sub importMasterData()
    Call getCurMdb
    Call xlTblToMdb("classTbl", "insert")
    Call xlTblToMdb("initTbl", "insert")
    Call xlTblToMdb("verTbl", "insert")
    Call xlTblToMdb("rankTbl", "insert")
    Call xlTblToMdb("comboTbl", "insert")
    Call xlTblToMdb("updateTbl", "insert")
End Sub

Sub importMusicData()
    mdata = txtAsTable(ThisWorkbook.path & "\data\musicData.csv", "YES")
    Call writeSchema(ThisWorkbook.path & "\data", Array("musicData.csv"))
    Call getCurMdb
    
    sSQL = "delete from MusicTbl"
    Call execSQL(sSQL)
    sSQL = "delete from MusicLevel"
    Call execSQL(sSQL)
    '
    sSQL = mkInsertIntoSQL("MusicTbl", mdata, "ID,num,title,verID,initID", "ID,num,title,verID,initID")
    Call execSQL(sSQL)
    '
    ary = Array("sg", "sb", "sd", "se", "sc", "db", "dd", "de", "dc")
    For i = 0 To 8
        cols = "ID," & i & "," & ary(i)
        sSQL = mkInsertIntoSQL("MusicLevel", mdata, "ID,classID,lev", cols, ary(i) & ">0")
        Call execSQL(sSQL)
    Next
    '
End Sub

Sub createScoreView()
    mdbPath = ThisWorkbook.path & "\data.mdb"
    Call mkView("Score0")
    Call mkView("MusicLevel0")
    Call mkView("ScoreView0")
    Call mkView("ScoreView")
    Call mkView("tmp0")
    Call mkView("horizontalData")
    Call mkView("rivalScoreView0")
    Call mkView("rivalScoreView")
    Call mkView("rivalScoreView1")
    Call mkView("previousScore0")
    Call mkView("rivalPreviousScore0")
    Call mkProc("horizontalData1")
    Call mkView("skillAttackData0")
    Call mkView("skillAttackData1")
    Call mkProc("p_rivalScoreView0")
    Call mkProc("p_rivalScoreView1")
    Call mkProc("p_rivalScoreView2")
    Call mkProc("p_rivalScoreView3")
    Call mkProc("p_rivalScoreView4")
    '
End Sub

Sub setSkill()
    Dim sqls(1 To 2)
    sqls(1) = "update (MusicLevel A inner join ScoreTbl B on A.ID=B.ID and A.classID=B.ClassID) " & _
    "Set skill=iif(score<=900000,0,cdbl((score-900000)*lev*2/100000+lev))"
    sqls(2) = "update ScoreTbl set skill=int(skill*100)/100.0"
    mdbPath = ThisWorkbook.path & "\data.mdb"
    Call execSQLs(sqls, mdbPath)
End Sub

Sub setRivalSkill(rival)
    Dim sqls(1 To 2)
    sqls(1) = "update (select * from MusicLevel A inner join rivalScoreTbl B on A.ID=B.ID and A.classID=B.ClassID where rivalId=" & rival & ") " & _
    "Set skill=iif(score<=900000,0,cdbl((score-900000)*lev*2/100000+lev))"
    sqls(2) = "update rivalScoreTbl set skill=int(skill*100)/100.0"
    mdbPath = ThisWorkbook.path & "\data.mdb"
    Call execSQLs(sqls, mdbPath)
End Sub

Sub prepareMusicData()
    Call getCurMdb
    Call xlTblToMdb("initTbl")
    Call xlTblToMdb("verTbl")
    Call xlTblToMdb("initData")
    Call xlTblToMdb("verData")
End Sub

Function mkHorizontalSQL()
    ary0 = Array("gsp", "bsp", "dsp", "esp", "csp", "bdp", "ddp", "edp", "cdp")
    Dim ary1(0 To 8)
    For i = 0 To 8
        ary1(i) = aggIf("data", "classID=" & i) & " as " & ary0(i)
    Next i
    sql = "select id ," & Join(ary1, "," & vbLf) & " " & vbLf & _
    " from scoreview0 group by id"
    mkHorizontalSQL = sql
End Function

Sub setHorizontalSQL()
    sql = mkHorizontalSQL
    Call TSetUp(sql, "horizontalData", "viewDef", "def")
    Call mkView("horizontalData")
End Sub

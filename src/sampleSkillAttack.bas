Attribute VB_Name = "sampleSkillAttack"
Option Base 0

Sub updateSkillData(code, pwd, Optional bSP = True, Optional bDP = True)
    frmSkillAttack.llblInfo = "送信データ作成"
    DoEvents
    Dim data
    data = getUpdateSkillData(bSP, bDP)
    Call writeSkill(code, pwd, data)
End Sub

Sub updateWholeSkillData(code, pwd, Optional bSP = True, Optional bDP = True)
    frmSkillAttack.llblInfo = "送信データ作成"
    DoEvents
    Dim data
    data = getWholeSkillData(bSP, bDP)
    Call writeSkill(code, pwd, data)
End Sub

Sub writeSkill(code, pwd, data)
    Call mkXhr
    Dim url, postdata
    url = " http://skillattack.com/sa4/dancer_input.php"
    postdata = "_=" & "&ddrcode=" & code & "&password=" & pwd
    Call xhr.Open("post", url)
    Call setUrlEncoded
    xhr.send (postdata)
    If xhr.responseText Like "<!DOCTYPE HTML*" Then
        frmSkillAttack.llblInfo = "ログイン成功"
    Else
        frmSkillAttack.llblInfo = xhr.responseText & vbCrLf & "処理を終了します"
        Exit Sub
    End If
    DoEvents
    '
    url = " http://skillattack.com/sa4/dancer_input.php"
    Call xhr.Open("post", url)
    Call setUrlEncoded
    postdata = "_=score_submit" & "&ddrcode=" & code & "&password=" & pwd & "&" & data
    Call xhr.Open("post", url)
    Call setUrlEncoded
    xhr.send (postdata)
    frmSkillAttack.llblInfo = "終了しました"
End Sub

Function getUpdateSkillData(Optional bSP = True, Optional bDP = True)
    Dim ret, rst, num, cid, data, clc, ary
    Dim sView
    sView = "skillAttackData0"
    ary0 = Array("index", "gsp", "bsp", "dsp", "esp", "csp", "bdp", "ddp", "edp", "cdp")
    ary1 = mapA("addstr", ary0, "", "%5B%5D=")
    Set clc = New Collection
    play = LCase(play)
    Set rst = CreateObject("ADODB.recordset")
    Call getCurMdb
    Call openMdbCon
    Call rst.Open(sView, adoCon)
    Do Until rst.EOF
        num = RTrim(rst("num"))
        cid = RTrim(rst("classID"))
        data = RTrim(rst("data"))
        If (cid < 5 And Not bSP) Or (cid >= 5 And Not bDP) Then GoTo skip
        ary = ary1
        ary(0) = ary(0) & num
        ary(cid + 1) = ary(cid + 1) & Trim(data)
        tmp = Join(ary, "&")
        clc.Add tmp
skip:
        rst.MoveNext
    Loop
    ret = Join(clcToAry(clc), "&")
    getUpdateSkillData = ret
End Function

Function getSkillDataById(id, Optional bSP = True, Optional bDP = True, Optional mdbPath = "")
    Dim ret, rst, num, cid, data, clc, ary, sql
    sql = "select * from score0 where id='" & id & "'"
    ary0 = Array("index", "gsp", "bsp", "dsp", "esp", "csp", "bdp", "ddp", "edp", "cdp")
    ary1 = mapA("addstr", ary0, "", "%5B%5D=")
    Set clc = New Collection
    play = LCase(play)
    Set rst = CreateObject("ADODB.recordset")
    Call getCurMdb
    num = getSqlVals("select num from musictbl where id='" & id & "'")(0)
    If IsNull(num) Then
        ret = ""
    Else
        ary = ary1
        ary(0) = ary(0) & num
        cnt = 0
        Call openMdbCon
        Call rst.Open(sql, adoCon)
        Do Until rst.EOF
            cid = CInt(RTrim(rst("classID")))
            data = RTrim(rst("data"))
            If (cid < 5 And bSP) Then
                ary(cid + 1) = ary(cid + 1) & Trim(data)
                cnt = cnt + 1
            End If
            If (cid >= 5 And bDP) Then
                ary(cid + 1) = ary(cid + 1) & Trim(data)
                cnt = cnt + 1
            End If
            rst.MoveNext
        Loop
        If cnt = 0 Then
            rer = ""
        Else
            ret = Join(ary, "&")
        End If
    End If
    getSkillDataById = ret
End Function

Function getWholeSkillData(Optional bSP = True, Optional bDP = True, Optional mdbPath = "")
    Dim ret, rst, num, cid, data, clc, ary
    Dim sView
    sTbl = "musicTbl"
    ary0 = Array("index", "gsp", "bsp", "dsp", "esp", "csp", "bdp", "ddp", "edp", "cdp")
    ary1 = mapA("addstr", ary0, "", "%5B%5D=")
    Set clc = New Collection
    play = LCase(play)
    Set rst = CreateObject("ADODB.recordset")
    Call getCurMdb(mdbPath)
    Call openMdbCon
    Call rst.Open(sTbl, adoCon)
    Do Until rst.EOF
        num = RTrim(rst("num"))
        id = RTrim(rst("ID"))
        If IsNull(num) Then GoTo skip1
        tmp = getSkillDataById(id, bSP, bDP)
        clc.Add tmp
skip1:
        rst.MoveNext
    Loop
    ret = Join(clcToAry(clc), "&")
    getWholeSkillData = ret
End Function

Sub importSkillNum()
    Call getCurMdb
    mdata = txtAsTable(ThisWorkbook.path & "\tmp\musicnum.csv", "YES")
    mdbPath = ThisWorkbook.path & "\data.mdb"
    '
    sSQL = mkSelectIntoSQL("MusicNumTbl", mdata)
    Call execSQL(sSQL)
    '
End Sub

Sub importMusicMasterData()
    Call getCurMdb
    mdata = txtAsTable(ThisWorkbook.path & "\data\musicData.csv", "YES")
    Call writeSchema(ThisWorkbook.path & "\data", Array("musicData.csv"))
    '
    sSQL = mkSelectIntoSQL("MusicMasterTbl", mdata)
    Call execSQL(sSQL)
    '
End Sub

Sub importSkillTxt(Optional toTbl = "SCoreTbl", Optional tsvFdr = "")
    If tsvFdr = "" Then tsvFdr = ThisWorkbook.path & "\tsv\"
    If Right(Trim(tsvFdr), 1) <> "\" Then tsvFdr = Trim(tsvFdr) & "\"
    Dim sqls(0 To 1)
    Set fso = CreateObject("Scripting.FileSystemObject")
    'Call fso.CopyFile(ThisWorkbook.path & "\data\schemaCSV.ini", tsvFdr & "schema.ini")
    sSQLs(0) = "delete from " & toTbl
    sd = IIf(i <= 4, "single", "double")
    fromTbl = txtAsTable(ThisWorkbook.path & "\skillAttack.txt")
    '
    sSQLs(1) = mkInsertIntoSQL(toTbl, fromTbl, "ID,classID,score,comboID", _
    Join(Array("id", i, "score" & i, "rank" & i, "combo" & i), ","), "rank" & i & ">0")
    sql = "insert into ScoreTbl(id,classId,score,comboId) select id, play*4+deg,score,iif(combo)"
    Call execSQLs(sSQLs, ThisWorkbook.path & "\data.mdb")
End Sub

Sub testgsd()
    data = getSkillData("skillAttackData0", "double")
    Debug.Print data
End Sub

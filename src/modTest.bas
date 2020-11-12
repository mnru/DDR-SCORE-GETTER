Attribute VB_Name = "modTest"
Sub testcss()
    mkIe
    Ie.navigate "https://p.eagate.573.jp/game/ddr/ddra20/p/playdata/music_detail.html?index=QPd01OQqbOIiDoO1dbdo1IIbb60bqPdl&diff=5"
    wait1
    wait2
    Set x = Ie.document.getElementById("difficulty")
    Set x0 = x.Children
    num = x0.Length
    For i = 0 To num - 1
        Debug.Print x0.Item(i).tagname
        Debug.Print x0.Item(i).innertext
    Next
    Stop
    For i = 0 To num - 1
        Debug.Print x0.Item(i).tagname
        Debug.Print x0.Item(i).innertext
    Next
    Stop
    Set y1 = x.getElementsByTagName("a")
    Set y2 = x.getElementsByTagName("img")
    n1 = y1.Length
    n2 = y2.Length
    Stop
    For i = 1 To n2
        y1.Add y2.Item(i)
    Next
    Stop
    Set Z = Ie.document.querySelectorAll("#difficulty > a")
    Stop
End Sub

Sub testcsv1()
    dbpath = ThisWorkbook.path & "\data.mdb"
    csvPath = ThisWorkbook.path & "\tmp\emoji.csv"
    tbl = txtAsTable(csvPath)
    sql = "select * into dscore from " & tbl
    Call execSQL(sql, dbpath)
End Sub

Sub testcsv2()
    dbpath = ThisWorkbook.path & "\data.mdb"
    csvPath = ThisWorkbook.path & "\out.csv"
    tbl = txtAsTable(csvPath)
    sql = "select * into " & tbl & " from dscore"
    Call execSQL(sql, dbpath)
End Sub

Sub testtb()
    mdbPath = ThisWorkbook.path & "\data.mdb"
    Call xlTblToMdb("song1", "song")
    Call xlTblToMdb("song2", "song[[F2]:[F3]]")
    Call xlTblToMdb("song3", "song[F2]")
End Sub

Sub testrg()
    mdbPath = ThisWorkbook.path & "\data.mdb"
    Call rangeToMdb(mdbPath, "fnc", ThisWorkbook.FullName, "Sheet1", "G3:M204")
    'Call xlTblToMdb(mdbpath, "procedure[[name]:[arg]]", "fnc")
    'Call xlTblToMdb(ThisWorkbook.path & "\data.mdb", "procedure", "fnc")
End Sub

Sub testInto()
    x0 = mkInsertIntoSQL("tblA", "tblB")
    x1 = mkInsertIntoSQL("tblA", "tblB", "a1,a2", "b1,b2", "b3='abc'")
    y0 = mkSelectIntoSQL("tblA", "tblB")
    y1 = mkSelectIntoSQL("tblA", "tblB", "a1,a2", "b1,b2", "b3='abc'")
    Debug.Print x0
    Debug.Print x1
    Debug.Print y0
    Debug.Print y1
End Sub

Sub testit()
    Set dic = CreateObject("Scripting.Dictionary")
    dic(2) = "name"
    x = mkIniTail(Array("Integer", "Char", "Date"))
    y = mkIniTail(Array("Integer", "Char", "Date"), dic)
    Debug.Print x
    Debug.Print
    Debug.Print y
End Sub

Sub testej()
    Dim A, B
    A = Array(3, 5)
    B = Array("a", "b")
    c = Array("7", "8")
    Dim r(1 To 21)
    r(1) = eachJoinAry("(", "", "", A)
    r(2) = eachJoinAry("(", "", "", A, B)
    r(3) = eachJoinAry("(", "", "", A, B, c)
    r(4) = eachJoinAry("(", "", ")", A)
    r(5) = eachJoinAry("(", "", ")", A, B)
    r(6) = eachJoinAry("(", "", ")", A, B, c)
    r(7) = eachJoinAry("", "", ")", A)
    r(8) = eachJoinAry("", "", ")", A, B)
    r(9) = eachJoinAry("", "", ")", A, B, c)
    r(10) = eachJoinAry("(", "+", "", A)
    r(11) = eachJoinAry("(", "+", "", A, B)
    r(12) = eachJoinAry("(", "+", "", A, B, c)
    r(13) = eachJoinAry("(", "+", ")", A)
    r(14) = eachJoinAry("(", "+", ")", A, B)
    r(15) = eachJoinAry("(", "+", ")", A, B, c)
    r(16) = eachJoinAry("", "+", ")", A)
    r(17) = eachJoinAry("", "+", ")", A, B)
    r(18) = eachJoinAry("", "+", ")", A, B, c)
    r(19) = eachJoinAry("", "+", "", A)
    r(20) = eachJoinAry("", "+", "", A, B)
    r(21) = eachJoinAry("", "+", "", A, B, c)
    For i = 1 To 21
        Debug.Print i;
        printAry r(i)
    Next
    Stop
End Sub

Sub testqt()
    Call getCurMdb
    Call displayQueryTable("dscore")
End Sub

Sub testForXlTbl()
    xlTbl = "abcd[werf]"
    'xlTbl = "werf"
    If InStr(xlTbl, "[") = 0 Then
        mdbTbl = xlTbl
    Else
        mdbTbl = Left(xlTbl, InStr(xlTbl, "[") - 1)
    End If
    Debug.Print mdbTbl
End Sub

Sub testsql()
    Dim sqls(1 To 2)
    'sqls(1) = "create table tmp2(name char,num integer) "
    sqls(1) = "insert into tmp2 values('zab',11230)"
    sqls(2) = "insert into tmp2 values('wcd',231110)"
    spath = ThisWorkbook.path & "\data2.mdb"
    Call execSQLs(sqls, spath)
    MsgBox "finished"
End Sub

Sub testsql2()
    pth1 = ThisWorkbook.path & "\data1.mdb"
    pth2 = ThisWorkbook.path & "\data2.mdb"
    sql = "insert into tmp1 select * from [" & pth2 & "].tmp2"
    Call execSQL(sql, pth1)
End Sub

Sub testdic()
    Set dic = CreateObject("Scripting.Dictionary")
    dic("a") = 1
    dic("b") = 2
    dic("c") = "none"
    Debug.Print dicToStr(dic, "=", ";")
End Sub

Sub testAryCon()
    x = Array("(", "+", "=", ")")
    y = Array(1, 2, 3, 4, 5)
    Z = Array(6, 7, 8, 9, 10)
    w = Array(7, 9, 11, 13, 15)
    A = eachConcateAry(x, y, Z, w)
    printAry A
End Sub

Sub testeja()
    'Z = Array()
    x = eachJoinAry("(", ",", ")", Split("", ","))
    y = eachConcateAry(Array("(", ")"), Split("", ","))
    Stop
End Sub

Sub testDiff()
    mdbPath = ThisWorkbook.path & "\data.mdb"
    Call mkDiffView(mdbPath, "tmp_d", "music2", "f1,f2", "f2,f12", "f1", "f2")
    Call mkDiffView(mdbPath, "tmp_d", "music1", "f1,f2", "f1,f12", "f1", "f1")
    Call mkDiffView(mdbPath, "music1", "music2", "f1,f12", "f2,f12", "f1", "f2")
End Sub

Sub testJoin()
    sql1 = mkJoinSQL("music1", "music2", "*", "*", "f1", "f2", "left")
    sql2 = "Create View left_music1_music2 as " & sql1
    Debug.Print sql1
    Debug.Print sql2
    Call execSQL(sql2, ThisWorkbook.path & "\data.mdb")
End Sub

Sub testDrop()
    ary = Array("tmp_d", "music1", "music2")
    mdbPath = ThisWorkbook.path & "\data.mdb"
    For i = 0 To 2
        For j = 0 To 2
            If i <> j Then
                sSQL = "Drop view " & Join(Array("Diff", ary(i), ary(j)), "_")
                Call execSQL(sSQL, mdbPath)
            End If
        Next
    Next
End Sub

Function testtmp()
    Dim sql
    csvTbl = txtAsTable(ThisWorkbook.path & "\tsv\double.csv")
    sql = mkInsertIntoSQL("ScoreTbl", csvTbl, "ID,classID,score,rank,combo", "F1,5,F3,F4,F5", "F4<>'none'")
    Debug.Print sql
    Call execSQL(sql, ThisWorkbook.path & "\data.mdb")
End Function

Sub testDel()
    sSQL = "delete from ScoreTbl"
    Call execSQL(sSQL, ThisWorkbook.path & "\data.mdb")
End Sub

Sub testimportMusic()
    For i = 0 To 2
        importMusic (i)
    Next
End Sub

Sub testTable()
    getCurMdb
    Call mkTable("tmp")
End Sub

Sub tstmkv()
    getCurMdb
    Call mkView("tmp")
End Sub

Sub testsch()
    Call writeSchema(ThisWorkbook.path & "\data", Array("musicData.csv"))
End Sub

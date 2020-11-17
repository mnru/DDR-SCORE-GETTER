Attribute VB_Name = "sampleIE"
Sub login()
    mkIe
End Sub

Sub gotoScorePage()
    Call mkIe
    url = "https://p.eagate.573.jp/game/ddr/ddra20/p/playdata/music_data_single.html"
    Ie.navigate (url)
End Sub

Sub getScoreIe(Optional sd = "double")
    Dim addrow, rNum, cNum, aNum
    Dim elm, elmNext
    Dim score
    Dim r, c
    url = "https://p.eagate.573.jp/game/ddr/ddra20/p/playdata/music_data_" & sd & ".html"
    Ie.navigate (url)
    wait1
    wait2
    ThisWorkbook.Sheets.Add After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    sn = ActiveSheet.name
    ThisWorkbook.Sheets(sn).Columns(2).NumberFormat = "@"
    addrow = 1
    d = Array(7, 5)
    Set fso = CreateObject("scripting.filesystemobject")
    Do
        Set tbl = Ie.document.getElementById("data_tbl")
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
                            ary(3 * c + i) = Right(fn, Len(fn) - d(i))
                        Next i
                End Select
            Next c
            ThisWorkbook.Sheets(sn).Cells(addrow, 1).Resize(1, aNum) = ary
            addrow = addrow + 1
        Next r
        'printAry (ary)
        If IsNull(Ie.document.getElementById("next")) Then
            Exit Do
        Else
            Set elmNext = Ie.document.getElementById("next")
            url = elmNext.getElementsByTagName("a")(0).href
            Ie.navigate (url)
            Call wait1
            Call wait2
        End If
        Sleep (3000)
    Loop
    ' ie.Quit
    '   MsgBox "èIóπÇµÇ‹ÇµÇΩ"
End Sub

Sub getTitleCatIe(filter, filterTypeTo, Optional sd = "double")
    Dim addrow, rNum, cNum, aNum
    Dim elm, elmNext
    Call mkIe
    ThisWorkbook.Sheets.Add After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    sn = ActiveSheet.name
    ThisWorkbook.Sheets(sn).Columns(2).NumberFormat = "@"
    addrow = 1
    Set fso = CreateObject("scripting.filesystemobject")
    For ft = 0 To filterTypeTo
        ost = 0
        Do
            url = "https://p.eagate.573.jp/game/ddr/ddra20/p/playdata/music_data_" & sd & ".html" & filterParam(ost, filter, ft)
            Ie.navigate (url)
            wait1
            wait2
            If ost = 0 Then
                Set elm = Ie.document.querySelector(".filter" & filter & " .filtertype" & ft)
                Cells(addrow, 1) = elm.innertext
            End If
            Set tbl = Ie.document.getElementById("data_tbl")
            rNum = tbl.Rows.Length
            cNum = tbl.Rows(0).Cells.Length
            ReDim ary(0 To 2)
            For r = 1 To rNum - 1
                Set elm = tbl.Rows(r).Cells(0)
                Set A = elm.getElementsByTagName("a")(0)
                aref = A.href
                ary(0) = Right(aref, Len(aref) - InStr(aref, "="))
                ary(1) = ft
                ary(2) = A.innertext
                ThisWorkbook.Sheets(sn).Cells(addrow, 2).Resize(1, 3) = ary
                addrow = addrow + 1
            Next r
            'printAry (ary)
            If IsNull(Ie.document.getElementById("next")) Then
                Exit Do
            Else
                ost = ost + 1
            End If
            Sleep (3000)
        Loop
    Next ft
    ' ie.Quit
    '   MsgBox "èIóπÇµÇ‹ÇµÇΩ"
End Sub

Function filterParam(offset, filter, filtertype)
    filterParam = "?offset=" & offset & "&filter=" & filter & "&filtertype=" & filtertype
End Function

Sub testft()
    Call getTitleCatIe(7, 17)
    Call getTitleCatIe(8, 36)
End Sub

Function getPageNumIE()
    Dim ret
    ret = Ie.getElementsByClassName("page_num").Count
    getPageNumIE = ret
End Function

Sub goAroundScoreSite()
    Call printTime("getScoreIe", "single")
    Call printTime("getScoreIe", "double")
    Ie.Quit
End Sub

Function getLevAry(id)
    Call mkIe
    url = "https://p.eagate.573.jp/game/ddr/ddra20/p/playdata/music_detail.html?index=" & id
    Ie.navigate url
    wait1
    wait2
    Set elms = Ie.document.querySelectorAll("#difficulty li.step img")
    n = elms.Length
    ReDim ret(0 To n + 1)
    ret(0) = id
    For i = 0 To n - 1
        pn = elms.Item(i).src
        tmp = Split(Right(pn, Len(pn) - InStrRev(pn, "_")), ".")(0)
        ret(i + 2) = IIf(tmp = "", 0, tmp)
    Next
    Set elm = Ie.document.querySelector("#music_info")
    ary = Split(elm.Rows(0).Cells(1).innertext, vbCrLf)
    ret(1) = ary(0)
    ' ret(2) = ary(1)
    printAry (ret)
    getLevAry = ret
    Ie.Quit
End Function

Sub getLevArySheet(ParamArray ids())
    Dim idAry
    idAry = ids
    Call mkIe
    ThisWorkbook.Sheets.Add
    sn = ActiveSheet.name
    Sheets(sn).Columns(2).NumberFormat = "@"
    rNum = 1
    Dim ret
    For Each id In idAry
        url = "https://p.eagate.573.jp/game/ddr/ddra20/p/playdata/music_detail.html?index=" & id
        Ie.navigate url
        wait1
        wait2
        Set elms = Ie.document.querySelectorAll("#difficulty li.step img")
        n = elms.Length
        ReDim ret(0 To n + 1)
        ret(0) = id
        For i = 0 To n - 1
            pn = elms.Item(i).src
            tmp = Split(Right(pn, Len(pn) - InStrRev(pn, "_")), ".")(0)
            ret(i + 2) = IIf(tmp = "", 0, tmp)
        Next
        Set elm = Ie.document.querySelector("#music_info")
        ary = Split(elm.Rows(0).Cells(1).innertext, vbCrLf)
        ret(1) = ary(0)
        ' ret(2) = ary(1)
        printAry (ret)
        Call layAryAt(ret, rNum, 1)
        rNum = rNum + 1
    Next id
    Ie.Quit
End Sub

Sub testLevAry()
    x = getLevAry("61QQi8i9Iliq66IOq1ib888b666o08O8")
    printAry (x)
    ary = Array("Dooii0960OioP6Q1l0qi68Q8Dbd9OO91", "dQbolPiDiiObioli96DlOo9OlP9DoDIi", "i6bl89d6OOi1O00qlQOIl8b8Qld6IiQP", "61Q6Q8OOiiQIbIi0l6l10qQ0Ii8P0Qb6")
    getLevArySheet (ary)
End Sub

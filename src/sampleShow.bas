Attribute VB_Name = "sampleShow"
Sub displayScoreView(Optional mdbPath = "")
    Call setCurMdb(mdbPath)
    Call delThisBookSheets("myScore", "myPivot")
    Call displayQueryTable(, "Select ver,init,title,play,deg,lev,score,rank,combo,skill From ScoreView")
    ActiveSheet.name = "myScore"
    Call mkMyPivot
End Sub

Sub mkMyPivot(Optional top = "A3")
    Call delThisBookSheets("myPivot")
    Set tbl = ThisWorkbook.Sheets("myScore").ListObjects(1)
    tbln = tbl.name
    ThisWorkbook.Sheets.Add
    ThisWorkbook.ActiveSheet.name = "myPivot"
    Set pvt = ThisWorkbook.PivotCaches.Create(xlDatabase, tbln).CreatePivotTable(Range(top))
    '
    Call pvt.AddDataField(pvt.PivotFields("title"), "ŒÂ” / title", xlCount)
    '
    pvt.PivotFields("play").Orientation = xlPageField
    pvt.PivotFields("rank").Orientation = xlColumnField
    pvt.PivotFields("lev").Orientation = xlRowField
End Sub

Sub displayRivalScoreView(Optional mdbPath = "")
    Call setCurMdb(mdbPath)
    Call delThisBookSheets("rivalScore", "rivalPivot")
    Call displayQueryTable(, "Select * From RivalScoreView1")
    ActiveSheet.name = "rivalScore"
    Call mkRivalPivot
End Sub

Sub mkRivalPivot(Optional top = "A3")
    Call delThisBookSheets("rivalPivot")
    Set tbl = ThisWorkbook.Sheets("rivalScore").ListObjects(1)
    tbln = tbl.name
    ThisWorkbook.Sheets.Add
    ThisWorkbook.ActiveSheet.name = "rivalPivot"
    Set pvt = ThisWorkbook.PivotCaches.Create(xlDatabase, tbln).CreatePivotTable(Range(top))
    '
    Call pvt.AddDataField(pvt.PivotFields("title"), "ŒÂ” / title", xlCount)
    '
    pvt.PivotFields("play").Orientation = xlPageField
    pvt.PivotFields("rivalID").Orientation = xlPageField
    pvt.PivotFields("lev").Orientation = xlColumnField
    pvt.PivotFields("diff").Orientation = xlRowField
    pvt.PivotFields("diff").DataRange.Cells(1).Group Start:=-1000000, End:=1000000, By:=50000
End Sub

Sub displayVsRivalScoreView()
    Dim rivalAry
    rivalAry = getRivalAry
    Dim num, c
    num = lenAry(rivalAry)
    Call setCurMdb
    sProc = "p_rivalScoreView" & num
    Call displayParamQuery(sProc, , rivalAry, , "A2", False)
    sn = ActiveSheet.name
    Columns(3).ColumnWidth = 40
    c = 6
    For i = 0 To num
        If i > 0 Then
            Cells(1, c) = rivalAry(i - 1)
        End If
        Columns(c).ColumnWidth = 8
        Columns(c + 1).ColumnWidth = 7
        Columns(c + 2).ColumnWidth = 4
        Columns(c + 3).ColumnWidth = 2
        Columns(c + 4).ColumnWidth = 5
        c = c + 5
    Next i
    For i = 1 To 2
        Columns(i).ColumnWidth = 5.5
    Next i
    For i = 4 To 6
        Columns(i).ColumnWidth = 3
    Next i
End Sub

Function getRivalAry()
    Dim tmp, clc, ret
    Set clc = New Collection
    For i = 1 To 4
        tmp = TLookup(i, "ƒ‰ƒCƒoƒ‹", "DDRCode", "num")
        If Not IsEmpty(tmp) Then
            clc.Add tmp
        End If
    Next i
    ret = clcToAry(clc)
    getRivalAry = ret
End Function

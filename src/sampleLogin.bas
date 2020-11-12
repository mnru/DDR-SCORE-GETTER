Attribute VB_Name = "sampleLogin"
Public captchaSheet
Public SelectedPicInfo
Public bCancel

Function getJsonStr(Optional bSave = False)
    Dim t1, t2
    t1 = Time
    Call mkXhr
    Dim ret
    On Error GoTo errDispose
    url = "https://p.eagate.573.jp/gate/p/login.html"
    Call xhr.Open("get", url)
    Call xhr.send
    url = "https://p.eagate.573.jp/gate/p/common/login/api/kcaptcha_generate.html"
    Call xhr.Open("post", url)
    Call xhr.send
    ret = xhr.responseText
    If bSave Then
        Set fso = CreateObject("Scripting.FileSystemObject")
        sfdr = ThisWorkbook.path & "\tmp\"
        If Not fso.FolderExists(sfdr) Then fso.CreateFolder (sfdr)
        spath = sfdr & "\json.txt"
        Set stm = fso.CreateTextFile(spath)
        stm.Write (ret)
        stm.Close
    End If
    getJsonStr = ret
    t2 = Time
    'Debug.Print "getJsonStr", Format(t2 - t1, "hh:mm:ss")
    Exit Function
errDispose:
    If xhr.Status = 12007 Then
        MsgBox "í êMÇ™Ç¬Ç»Ç™Ç¡ÇƒÇ¢Ç‹ÇπÇÒ"
        Set xhr = Nothing
        bCancel = True
    End If
End Function

Function mkKeyList(Optional jsonStr = "")
    Dim json
    Dim keylist(0 To 7)
    Set fso = CreateObject("Scripting.FileSystemObject")
    If jsonStr = "" Then
        spath = ThisWorkbook.path & "\tmp\" & "json.txt"
        Set stm = fso.OpenTextFile(spath)
        jsonStr = stm.readall
        stm.Close
    End If
    Set json = JsonConverter.ParseJson(jsonStr)
    cpic = json("data")("correct_pic")
    keylist(0) = fso.getfilename(cpic)
    keylist(6) = Left(cpic, InStrRev(cpic, "/"))
    keylist(7) = json("data")("kcsess")
    Set choicelist = json("data")("choicelist")
    For i = 1 To 5
        keylist(i) = choicelist(i)("key")
    Next i
    'printAry (keylist)
    mkKeyList = keylist
End Function

Sub dlImgs(keylist, Optional bDelFiles = True)
    Set fso = CreateObject("Scripting.FileSystemObject")
    imgDir = ThisWorkbook.path & "\img\"
    If Not fso.FolderExists(imgDir) Then
        fso.CreateFolder (imgDir)
    End If
    Set imgFdr = fso.GetFolder(imgDir)
    If bDelFiles Then
        For Each fl In imgFdr.Files
            fl.Delete
        Next fl
    End If
    Call mkXhr
    For i = 0 To 5
        url = keylist(6) & keylist(i)
        savepath = imgDir & keylist(i) & ".png"
        Call xhr.Open("get", url)
        xhr.send
        resBodyToFile (savepath)
        ' Sleep (1000)
    Next
End Sub

Sub prepareLoginSheet()
    bCancel = False
    Dim t1, t2
    t1 = Time
    str0 = ""
    str0 = getJsonStr
    If bCancel Then Exit Sub
    keylist = mkKeyList(str0)
    Call dlImgs(keylist)
    Call addLoginSheet(keylist, False)
    t2 = Time
    Debug.Print "prepareLoginSheet", Format(t2 - t1, "hh:mm:ss")
End Sub

Sub addLoginSheet(keylist, Optional bDlImg = True)
    Dim imgUrl, imgDir, picPath
    Dim i
    imgUrl = keylist(6)
    imgDir = ThisWorkbook.path & "\img\"
    If bDlImg Then
        dlImgs (keylist)
    End If
    ThisWorkbook.Sheets.Add After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    sn = ActiveSheet.name
    Sheets(sn).Cells(1, 1).Resize(1, 8).Font.ThemeColor = xlThemeColorDark1
    Sheets(sn).Rows(1).NumberFormat = "@"
    For i = 0 To 7
        Sheets(sn).Cells(1, i + 1).Value = keylist(i)
    Next i
    cellwidth = 10
    cellHight = 60
    Cells.ColumnWidth = cellwidth
    Cells.RowHeight = cellHight
    For i = 0 To 5
        picPath = imgDir & keylist(i) & ".png"
        If i = 0 Then
            Cells(2, 2).Select
        Else
            Set rg = Cells(4, 2 * i)
            ActiveSheet.CheckBoxes.Add(rg.Left, rg.top, rg.width, rg.Height).Select
            rg.offset(1, 0).Select
        End If
        Sheets(sn).Pictures.Insert(picPath).Select
        With Selection.ShapeRange
            .LockAspectRatio = msoTrue
            .width = 60
        End With
    Next i
    Set rg = Cells(2, 4)
    ActiveSheet.Buttons.Add(rg.Left, rg.top, rg.width, rg.Height / 2).Select
    Selection.OnAction = "showLoginForm"
    Selection.Characters.text = "ëIë"
    ActiveSheet.Range("B2").Select
End Sub

Sub showLoginForm()
    captchaSheet = ActiveSheet.name
    Dim str0
    str0 = ""
    For i = 1 To 5
        If ThisWorkbook.Sheets(captchaSheet).CheckBoxes(i).Value = 1 Then
            If str0 <> "" Then str0 = str0 & ","
            str0 = str0 & i
        End If
    Next
    If str0 = "" Then str0 = "Ç»Çµ"
    SelectedPicInfo = str0
    frmLogin.Show
End Sub

Function mkCaptchaVal(captchaSheet)
    Dim ret
    Dim tmplist(0 To 7)
    For i = 0 To 7
        tmplist(i) = ThisWorkbook.Sheets(captchaSheet).Cells(1, i + 1)
    Next i
    ret = "k_" & tmplist(7)
    For i = 1 To 5
        ret = ret & "_"
        If ThisWorkbook.Sheets(captchaSheet).CheckBoxes(i).Value = 1 Then
            ret = ret & tmplist(i)
        End If
    Next i
    mkCaptchaVal = ret
End Function

Function execLogin(sLoginName, sPassword, captchaSheet)
    Dim t1, t2
    t1 = Time
    Dim json
    Dim ret, res, code, rescode
    data = mkLoginData(sLoginName, sPassword, mkCaptchaVal(captchaSheet))
    'Debug.Print data
    '
    If IsEmpty(xhr) Then Call mkXhr
    url = "https://p.eagate.573.jp/gate/p/common/login/api/login_auth.html"
    Call xhr.Open("post", url)
    Call xhr.setRequestHeader("Content-Type", "application/x-www-form-urlencoded; charset=UTF-8")
    Call xhr.send(data)
    res = xhr.responseText
    res = Trim(res)
    'frmLogin.Hide
    If LCase(Left(res, 6)) = "<html>" Then
        rescode = -1
    Else
        Set json = JsonConverter.ParseJson(res)
        rescode = json("fail_code")
    End If
    Debug.Print rescode
    execLogin = rescode
    t2 = Time
    Debug.Print "execLogin", Format(t2 - t1, "hh:mm:ss")
    DoEvents
End Function

Function mkLoginData(sLoginName, sPassword, sCaptchaVal)
    Dim ret
    Set dic = CreateObject("scripting.dictionary")
    dic("login_id") = sLoginName
    dic("pass_word") = sPassword
    dic("otp") = ""
    dic("resrv_url") = "/gate/p/login_complete.html"
    dic("captcha") = sCaptchaVal
    ret = encodeDic(dic)
    'Debug.Print ret
    mkLoginData = ret
    Set dic = Nothing
End Function

Function mkOffset(i, rival)
    Dim ret
    If i = 0 And rival = "" Then
        ret = ""
    Else
        ret = "?offset=" & i & "&filter=0&filtertype=0&sorttype=0"
    End If
    mkOffset = ret
End Function

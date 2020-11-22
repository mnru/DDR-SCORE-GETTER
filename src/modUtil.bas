Attribute VB_Name = "modUtil"
Sub CheckGUID()
    Sheets.Add
    sn = ActiveSheet.name
    i = 1
    Dim ref
    For Each ref In ActiveWorkbook.VBProject.references
        ary = Array(ref.name, ref.GUID, ref.major, ref.minor)
        Sheets(sn).Cells(i, 1).Resize(1, 4) = ary
        i = i + 1
    Next
End Sub

Sub delTmpSheet(Optional bn = "")
    Application.DisplayAlerts = False
    If bn = "" Then bn = ThisWorkbook.name
    For Each sh In Workbooks(bn).Sheets
        If sh.name Like "Sheet*" Then
            sh.Delete
        End If
    Next
    ThisWorkbook.Save
    Application.DisplayAlerts = False
End Sub

Sub delThisBookSheets(ParamArray shts())
    On Error Resume Next
    Call delTmpSheet
    shtAry = shts
    Application.DisplayAlerts = False
    If bn = "" Then bn = ThisWorkbook.name
    For Each sh1 In Workbooks(bn).Sheets
        For Each sh2 In shtAry
            If sh1.name = sh2 Then
                sh1.Delete
            End If
        Next sh2
    Next sh1
    ThisWorkbook.Save
    Application.DisplayAlerts = False
    On Error GoTo 0
End Sub

Sub delQueryTables()
    For Each sh In ThisWorkbook.Sheets
        For Each qt In sh.QueryTables
            qt.Delete
        Next
    Next
End Sub

Sub refreshThisWorkbook(Optional fdrn = "tmp0")
    ToFdr = ThisWorkbook.path & "\" & fdrn & "\"
    tofn = fdrn & ".xlsm"
    toPath = ToFdr & tofn
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CreateFolder (ToFdr)
    fso.CreateFolder (ToFdr & "\src\")
    Call fso.CopyFile(ThisWorkbook.path & "\src\*", ToFdr & "src\")
    Call fso.CopyFile(ThisWorkbook.path & "\Fix_xlsm_compose.vbs", ToFdr)
    Call fso.CopyFile(ThisWorkbook.path & "\Fix_xlsm_decompose.vbs", ToFdr)
    num = ThisWorkbook.Sheets.Count
    ReDim shs(1 To num)
    For i = 1 To num
        shs(i) = ThisWorkbook.Sheets(i).name
    Next i
    ThisWorkbook.Sheets(shs).Copy
    ActiveWorkbook.SaveAs toPath, xlOpenXMLWorkbookMacroEnabled
    Workbooks(tofn).Close
End Sub

Function utf8Join(ary, Optional dlm = "", Optional withCrlf = False) As Byte()
    Dim stm, ret
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2 'adTypeText
    stm.charset = "UTF-8"
    stm.Open
    lb = LBound(ary)
    ub = UBound(ary)
    For i = lb To ub
        Call stm.WriteText(ary(i), 0) '0:adWriteChar,1:adWriteLine
        If i < ub Then
            Call stm.WriteText(dlm, 0)
        Else
            If withCrlf Then
                Call stm.WriteText("", 1)
            End If
        End If
    Next i
    stm.Position = 0
    stm.Type = 1
    stm.Position = 3
    ret = stm.Read
    'stm.Position = stm.Size
    stm.Close
    utf8Join = ret
End Function

Sub backUpDef()
    Set fso = CreateObject("Scripting.FileSystemObject")
    defAry = Array("ProcDef", "TblDef", "ViewDef", "SchemaDef")
    saveDir = ThisWorkbook.path & "\def"
    If Not fso.FolderExists(saveDir) Then fso.CreateFolder (saveDir)
    For Each tbl In defAry
        Call xlTblToCsv(tbl, saveDir)
    Next tbl
End Sub

Sub registerOrder()
    On Error Resume Next
    Application.AddCustomList ListArray:=Sheets("data").Range("rankTbl[rank]")
    Application.AddCustomList ListArray:=Sheets("data").Range("comboTbl[combo]")
    'Application.AddCustomList ListArray:=Sheets("data").Range("verTbl[ver]")
    On Error GoTo 0
End Sub


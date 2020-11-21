Attribute VB_Name = "Old"
Sub htmlToScoreDBOld(Optional rival = "")
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
    MsgBox "èIóπÇµÇ‹ÇµÇΩ"
End Sub

Sub mkImportTxtSQLsOld(Optional toTbl = "tmp", Optional rival = "", Optional bSP = True, Optional bDP = True, Optional bDel = True, Optional tsvFdr = "")
    Dim clc
    Dim toTbl
    Set clc = New Collection
    If tsvFdr = "" Then tsvFdr = ThisWorkbook.path & "\tsv\"
    If Right(Trim(tsvFdr), 1) <> "\" Then tsvFdr = Trim(tsvFdr) & "\"
    Set fso = CreateObject("Scripting.FileSystemObject")
End Sub

Sub updateMusicDataOld()
    Call dlMusicData
    Call getCurMdb
    Call importMusicData
    MsgBox "èIóπÇµÇ‹ÇµÇΩ"
End Sub


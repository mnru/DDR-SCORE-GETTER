VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLogin 
   Caption         =   "ログイン・ダウンロード"
   ClientHeight    =   4740
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4200
   OleObjectBlob   =   "frmLogin.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdLogin_Click()
    Me.llblInfo.Caption = ""
    DoEvents
    res = execLogin(tbxName.Value, tbxPwd.Value, captchaSheet)
    Me.llblInfo.Caption = IIf(res = 0, "ログイン成功", "ログイン失敗")
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(captchaSheet).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    ThisWorkbook.Sheets("menu").Activate
    DoEvents
    If res = 0 Then
        If cbxSP Then
            Call dlScores("single", tbxRival.Value)
        End If
        If cbxDP Then
            Call dlScores("double", tbxRival.Value)
        End If
        If cbxSP Then
            Call sdHtmlToTsv("single", tbxRival.Value)
        End If
        If cbxDP Then
            Call sdHtmlToTsv("double", tbxRival.Value)
        End If
        Call importTsvToScoreDB(tbxRival.Value)
    End If
End Sub

Private Sub UserForm_Initialize()
    Me.tbxName = TLookup("ログイン名", "認証", "値")
    Me.tbxPwd = TLookup("パスワード", "認証", "値")
    Me.lblSelectedPic.Caption = SelectedPicInfo
    cbxSP = True
    cbxDP = True
End Sub

VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSkillAttack 
   Caption         =   "SkillAttack"
   ClientHeight    =   4590
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmSkillAttack.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmSkillAttack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdLogin_Click()
    If opgUpdate Then
        Call updateSkillData(tbxCode, tbxPwd, cbxSP, cbxDP)
    ElseIf opgWhole Then
        Call updateWholeSkillData(tbxCode, tbxPwd, cbxSP, cbxDP)
    End If
End Sub

Private Sub UserForm_Initialize()
    Me.tbxCode = TLookup("DDRコード(8桁)", "認証", "値")
    Me.tbxPwd = TLookup("SkillAttackパスワード", "認証", "値")
    cbxSP = True
    cbxDP = True
    opgUpdate = True
End Sub

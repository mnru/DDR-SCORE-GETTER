VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmUpdateMusicData 
   Caption         =   "楽曲データ更新"
   ClientHeight    =   1740
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3720
   OleObjectBlob   =   "frmUpdateMusicData.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmUpdateMusicData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdUpdateMusicData_Click()
    If chkDownLoad Then
        Call dlMusicData
    End If
    Call importMusicData
    
    Me.lblInfo.Caption = "終了しました"
End Sub
    

Private Sub UserForm_Initialize()
    Me.chkDownLoad.Value = True
    Me.lblInfo.Caption = ""
End Sub

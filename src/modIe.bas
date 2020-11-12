Attribute VB_Name = "modIe"
#If VBA7 And Win64 Then
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal msec As Long)
#Else
Public Declare Sub Sleep Lib "kernel32" (ByVal msec As Long)
#End If
Public Ie

Sub mkIe(Optional bVisible = True)
    Set Ie = CreateObject("InternetExplorer.Application")
    Ie.Visible = bVisible
End Sub

Sub wait1()
    Do While Ie.busy Or Ie.readyState < 4
        DoEvents
    Loop
End Sub

Sub wait2()
    Dim elm
    Do
        Err.Clear
        Set elm = Ie.document.getElementsByTagName("html")(0)
        DoEvents
    Loop Until Err.Number = 0
    Set elm = Nothing
End Sub

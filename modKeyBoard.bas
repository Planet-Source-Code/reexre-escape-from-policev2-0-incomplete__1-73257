Attribute VB_Name = "modKeyBoard"
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer


Public Sub KEYBOARD()


    If GetAsyncKeyState(vbKeyA) <> 0 Then B(1).cAccellerate
    If GetAsyncKeyState(vbKeyZ) <> 0 Then B(1).cBrake
    If GetAsyncKeyState(vbKeyN) <> 0 Then B(1).cDoSteer -1
    If GetAsyncKeyState(vbKeyM) <> 0 Then B(1).cDoSteer 1

    If GetAsyncKeyState(vbKeyEscape) <> 0 Then Unload frmMAIN
End Sub

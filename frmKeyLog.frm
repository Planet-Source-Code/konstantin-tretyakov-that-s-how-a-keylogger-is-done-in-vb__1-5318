VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VB Keyboard Hook Example - Konstantin Tretyakov"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   Icon            =   "frmKeyLog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   5535
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Dummy 
      Caption         =   "Dumb button to receive messages"
      Height          =   495
      Left            =   1560
      TabIndex        =   4
      Top             =   2760
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtLog 
      Height          =   3735
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   840
      Width           =   5295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Exit (Alt+E)"
      Height          =   615
      Left            =   3720
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdRemoveHk 
      Caption         =   "&Unhook (Alt+U)"
      Height          =   615
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdSetHk 
      Caption         =   "&Hook (Alt+H)"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdRemoveHk_Click()
    UnhookKeyboard
    cmdSetHk.Enabled = True
End Sub

Private Sub cmdSetHk_Click()
    'Start the hook, and set the Receiver:txtReceiver, Event:KeyDown
    'Now every systemwide key-event will be also direced
    'to the textbox.
    If HookKeyboard(Dummy.hWnd, MyOwnMessage) = 0 Then MsgBox "Failed to hook keyboard, probably another application is using the hook." & vbCrLf & "Press Alt+U, then try again.", vbCritical, "Error"
    cmdSetHk.Enabled = False
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'Get control over form's WindowProc
    PrevFuncPointer = SetWindowLong(Dummy.hWnd, GWL_WNDPROC, AddressOf Window_OnMessage)
    If PrevFuncPointer = 0 Then
        MsgBox "Something is wrong in your system. Please, format your drive C, reinstall the system and try again", vbCritical, "Error"
        Unload Me
    End If
   'Start the hook from the very beginning
   cmdSetHk_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Remove the hook
    UnhookKeyboard
End Sub


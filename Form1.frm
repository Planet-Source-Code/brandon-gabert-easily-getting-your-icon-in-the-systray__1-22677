VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuPopupmenu 
      Caption         =   "PopupMenu"
      Begin VB.Menu mnuShow 
         Caption         =   "&Show"
      End
      Begin VB.Menu mnuHide 
         Caption         =   "&Hide"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "&Quit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'constants required by Shell_NotifyIcon API call:
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201     'Button down
Private Const WM_LBUTTONUP = &H202       'Button up
Private Const WM_LBUTTONDBLCLK = &H203    'Double-click
Private Const WM_RBUTTONDOWN = &H204     'Button down
Private Const WM_RBUTTONUP = &H205       'Button up
Private Const WM_RBUTTONDBLCLK = &H206   'Double-click

Private Declare Function SetForegroundWindow Lib "user32" _
    (ByVal hwnd As Long) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32" _
    Alias "Shell_NotifyIconA" _
    (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Private nid As NOTIFYICONDATA

'user defined type required by Shell_NotifyIcon API call
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Sub Form_Load()
    Me.Show
    Me.Refresh
    With nid
        .cbSize = Len(nid)
        .hwnd = Me.hwnd
        .uId = vbNull
        ''''''
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        ''''''The callback should be the mousemove event
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        ''''''Heres the tooltip in the taskbar'''''
        .szTip = "Your app name" & vbNullChar
    End With
    Shell_NotifyIcon NIM_ADD, nid
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim Result As Long
    Dim msg As Long

    'set up the msg handler
    If Me.ScaleMode = vbPixels Then
        msg = x
    Else
        msg = x / Screen.TwipsPerPixelX
    End If

    Select Case msg
        'respond to the right mouse button
        Case WM_RBUTTONUP
            'bring our window to the foreground
            Result = SetForegroundWindow(Me.hwnd)
            'mnuPopupMenu is in the menu editor for this form
            'it is set to be invisible
            Me.PopupMenu mnuPopupmenu
        'respond to the left double click
        Case WM_LBUTTONDBLCLK
            Me.Show
            Result = SetForegroundWindow(Me.hwnd)
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'remove the icon
    Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub mnuHide_Click()
    Me.Hide
End Sub

Private Sub mnuQuit_Click()
    Unload Me
End Sub

Private Sub mnuShow_Click()
    Me.Show
End Sub

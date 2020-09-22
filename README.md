﻿<div align="center">

## Easily Getting Your Icon In The Systray

<img src="PIC2001425245315265.jpg">
</div>

### Description

Easily add and delete your applications icon from the System Tray (systray), and the icon responds to all mouse clicks. lost all the formatting when i uploaded this article, very disappointing, the readme.txt contains the formatted article, sorry
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2001-04-25 01:01:02
**By**             |[Brandon Gabert](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/brandon-gabert.md)
**Level**          |Beginner
**User Rating**    |4.3 (13 globes from 3 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Easily Get188094252001\.zip](https://github.com/Planet-Source-Code/brandon-gabert-easily-getting-your-icon-in-the-systray__1-22677/archive/master.zip)





### Source Code

```
Hi, my first article submission.
First, we have to declare the API's and the constants we will be using.
'constants required by Shell_NotifyIcon API call:
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const WM_MOUSEMOVE = &H200
'all these are for the mousemouve event
Private Const WM_LBUTTONDOWN = &H201 'Button down
Private Const WM_LBUTTONUP = &H202 'Button up
Private Const WM_LBUTTONDBLCLK = &H203 'Double-click
Private Const WM_RBUTTONDOWN = &H204 'Button down
Private Const WM_RBUTTONUP = &H205 'Button up
Private Const WM_RBUTTONDBLCLK = &H206 'Double-click
Private Declare Function SetForegroundWindow Lib "user32" _
(ByVal hwnd As Long) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
'and 1 type that we need
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
Basically, I will just be explaining the use of Shell_NotifyIcon from here. The calls to SetForegroundWindow are pretty simple. Heres the code that goes into the form load code so that it will put itself into the system tray, I would suggest making the Form1.ShowInTaskBar = false.
Private Sub Form_Load()
	Me.Show
	Me.Refresh
	With nid
		.cbSize = Len(nid)
		.hwnd = Me.hwnd
		.uId = vbNull
		.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
		''''''The callback should be the mousemove event
		.uCallBackMessage = WM_MOUSEMOVE
		.hIcon = Me.Icon
		''''''Heres the tooltip in the taskbar'''''
		.szTip = "Your app name" & vbNullChar
	End With
	Shell_NotifyIcon NIM_ADD, nid
End Sub
and now remove the icon when we unload
Private Sub Form_Unload(Cancel As Integer)
	'remove the icon
	Shell_NotifyIcon NIM_DELETE, nid
End Sub
'hide the form when the menuitem is clicked
Private Sub mnuHide_Click()
	Me.Hide
End Sub
'show the form when the menuitem is clicked
Private Sub mnuShow_Click()
	Me.Show
End Sub
'unload the form when we click the quit menuitem
Private Sub mnuQuit_Click()
	Unload Me
End Sub
Thanks for a lot of great responses on my Streaming Screenshots project, but I need more globes :-) Brandon
```


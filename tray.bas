Attribute VB_Name = "Module1"
'const and declares for disable the X in the titelbar
Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long


Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
    
    Public Const SC_CLOSE = &HF060&
    Public Const MF_BYCOMMAND = &H0&



'shut down const and declare
Public Const EWX_SHUTDOWN = 1
Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dw As Long) As Long

      'user defined type required by Shell_NotifyIcon API call
      Public Type NOTIFYICONDATA
       cbSize As Long
       hwnd As Long
       uId As Long
       uFlags As Long
       uCallBackMessage As Long
       hIcon As Long
       szTip As String * 64
      End Type

      'constants required by Shell_NotifyIcon API call:
      Public Const NIM_ADD = &H0
      Public Const NIM_MODIFY = &H1
      Public Const NIM_DELETE = &H2
      Public Const NIF_MESSAGE = &H1
      Public Const NIF_ICON = &H2
      Public Const NIF_TIP = &H4
      Public Const WM_MOUSEMOVE = &H200
      Public Const WM_LBUTTONDOWN = &H201     'Button down
      Public Const WM_LBUTTONUP = &H202       'Button up
      Public Const WM_LBUTTONDBLCLK = &H203   'Double-click
      Public Const WM_RBUTTONDOWN = &H204     'Button down
      Public Const WM_RBUTTONUP = &H205       'Button up
      Public Const WM_RBUTTONDBLCLK = &H206   'Double-click

      Public Declare Function SetForegroundWindow Lib "user32" _
      (ByVal hwnd As Long) As Long
      Public Declare Function Shell_NotifyIcon Lib "SHELL32" _
      Alias "Shell_NotifyIconA" _
      (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

      Public nid As NOTIFYICONDATA




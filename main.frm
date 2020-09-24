VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Timing ShutDown"
   ClientHeight    =   7095
   ClientLeft      =   -405
   ClientTop       =   750
   ClientWidth     =   4455
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command8 
      Caption         =   "About"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   120
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Caption         =   "Shut down by minutes from now"
      Height          =   4215
      Left            =   120
      TabIndex        =   16
      Top             =   2760
      Visible         =   0   'False
      Width           =   4215
      Begin VB.CommandButton Command5 
         Caption         =   "OK"
         Height          =   495
         Left            =   1680
         TabIndex        =   22
         Top             =   3480
         Width           =   855
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   3600
         Width           =   735
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Exit"
         Height          =   375
         Left            =   3240
         TabIndex        =   20
         Top             =   3600
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2040
         MaxLength       =   4
         TabIndex        =   18
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "minutes from now."
         Height          =   255
         Left            =   2760
         TabIndex        =   19
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Shut down my computer"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1200
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Shut down by date"
      Height          =   4215
      Left            =   120
      TabIndex        =   5
      Top             =   2760
      Width           =   4215
      Begin VB.ComboBox Combo5 
         Height          =   315
         ItemData        =   "main.frx":0442
         Left            =   2160
         List            =   "main.frx":046A
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   2760
         Width           =   1095
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         ItemData        =   "main.frx":049E
         Left            =   840
         List            =   "main.frx":04EA
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   2760
         Width           =   1095
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "main.frx":0545
         Left            =   2160
         List            =   "main.frx":056D
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1800
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "main.frx":05D3
         Left            =   840
         List            =   "main.frx":0634
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Exit"
         Height          =   375
         Left            =   3240
         TabIndex        =   8
         Top             =   3600
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   3600
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "OK"
         Height          =   495
         Left            =   1680
         TabIndex        =   6
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   ":"
         Height          =   255
         Left            =   1920
         TabIndex        =   24
         Top             =   2760
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "When do you want your computer to shut down?"
         Height          =   495
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Width           =   3495
      End
      Begin VB.Label Label4 
         Caption         =   "Choose the time:"
         Height          =   255
         Left            =   840
         TabIndex        =   12
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Choose the date:"
         Height          =   255
         Left            =   840
         TabIndex        =   11
         Top             =   1440
         Width           =   1335
      End
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3480
      Top             =   0
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Shut down by minutes from now"
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   1440
      Width           =   2895
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Shut down by date"
      Height          =   195
      Left            =   720
      TabIndex        =   3
      Top             =   1080
      Value           =   -1  'True
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "_"
      Height          =   255
      Left            =   4200
      TabIndex        =   2
      ToolTipText     =   "Minimize to tray"
      Top             =   0
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3360
      Top             =   0
   End
   Begin VB.Label Label6 
      Caption         =   "Version 1.0"
      Height          =   255
      Left            =   960
      TabIndex        =   25
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   600
      TabIndex        =   23
      Top             =   2040
      UseMnemonic     =   0   'False
      Width           =   3495
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Left            =   480
      Top             =   1920
      Width           =   3735
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   1800
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   975
   End
   Begin VB.Menu MnuFile 
      Caption         =   "&File"
      Visible         =   0   'False
      Begin VB.Menu MnuRestore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu MnuExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X As String, mon As String, t As Long

Private Sub Command1_Click()
Dim AMPM As String, X As Integer
'For the first item put "01", for the second item put "02 etc.
Select Case Combo2.ListIndex
Case Is = 0
    mon = "01"
Case Is = 1
    mon = "02"
Case Is = 2
    mon = "03"
Case Is = 3
    mon = "04"
Case Is = 4
    mon = "05"
Case Is = 5
    mon = "06"
Case Is = 6
    mon = "07"
Case Is = 7
    mon = "08"
Case Is = 8
    mon = "09"
Case Is = 9
    mon = "10"
Case Is = 10
    mon = "11"
Case Is = 11
    mon = "12"
End Select
    
Frame1Disabled
X = Combo4.Text
'check wether it's AM or PM
If X <= 12 Then
    AMPM = "AM"
Else
    AMPM = "PM"
End If
Label7.Caption = "Shutting down at  " & Combo1.Text & " " & Combo2.Text & "  "
Label7.Caption = Label7.Caption & Combo4.Text & ":" & Combo5.Text & " " & AMPM
Me.Caption = "Shutting down at  " & Combo1.Text & " " & Combo2.Text & "  "
Me.Caption = Me.Caption & Combo4.Text & ":" & Combo5.Text & " " & AMPM
End Sub

Private Sub Command2_Click()
Frame1Enabled
End Sub

Private Sub Command3_Click()
Shell_NotifyIcon NIM_DELETE, nid 'delete tray icon
End
End Sub

Private Sub Command4_Click()
Me.Visible = False
End Sub

Private Sub Command5_Click()
If (Text2.Text = "") Or (IsNumeric(Text2.Text) = False) Or (Text2.Text <= "0") Then
    If MsgBox("Please enter a number between 1 to 9,999.", , "") = vbOK Then
        Text2.Text = ""
    End If
Else
    Frame2Disabled
    t = Text2.Text * 60000 '60 seconds = 1 minute
    Timer2.Interval = 1000
    Timer2.Enabled = True
End If
Label7.Caption = t / 1000 & " " & "seconds until shut down"
Me.Caption = t / 1000 & " " & "seconds until shut down"
End Sub

Private Sub Command6_Click()
Timer2.Enabled = False
Frame2Enabled
End Sub

Private Sub Command7_Click()
Shell_NotifyIcon NIM_DELETE, nid 'delete tray icon
End
End Sub

Private Sub Command8_Click()
Form3.Show 1
End Sub

Private Sub Form_Load()
'remove the X in titlebar
Dim hSysMenu As Long
hSysMenu = GetSystemMenu(hwnd, False)
RemoveMenu hSysMenu, SC_CLOSE, MF_BYCOMM
'show tray icon
   With nid                                        ''''''''''''''''''''''''''
      .cbSize = Len(nid)                           'These are the parameters'
      .hwnd = Me.hwnd                              'of the tray icon. You   '
      .uId = vbNull                                'have to put all this    '
      .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE 'before you show the tray'
      .uCallBackMessage = WM_MOUSEMOVE             'itself (look down).     '
      .hIcon = Me.Icon                             '                        '
      .szTip = "Timing ShutDown" & vbNullChar      '                        '
   End With                                        ''''''''''''''''''''''''''
Shell_NotifyIcon NIM_ADD, nid 'show tray icon
'"cancel" is disabled in the beginning
Command2.Enabled = False
Command6.Enabled = False
'show the current time
Label1.Caption = "Current time:"
Label2.Caption = Now
'show the first item in the list of every combo box
Combo1.ListIndex = 0
Combo2.ListIndex = 0
Combo4.ListIndex = 0
Combo5.ListIndex = 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
   Dim Result As Long
   Dim msg As Long
   If Me.ScaleMode = vbPixels Then
      msg = X
   Else
      msg = X / Screen.TwipsPerPixelX
   End If
      
   Select Case msg
      Case WM_LBUTTONDBLCLK
          Me.WindowState = vbNormal
          Result = SetForegroundWindow(Me.hwnd)
          Me.Visible = True
      Case WM_RBUTTONUP
         PopupMenu MnuFile
   End Select
End Sub

Private Sub MnuExit_Click()
Shell_NotifyIcon NIM_DELETE, nid 'delete tray icon
End
End Sub

Private Sub MnuRestore_Click()
Me.Visible = True
End Sub

Private Sub Option1_Click()
Frame1.Visible = True
Frame2.Visible = False
Timer2.Enabled = False
Frame2Enabled
End Sub

Private Sub Option2_Click()
Frame1.Visible = False
Frame2.Visible = True
Timer2.Interval = 0
Timer2.Enabled = True
Frame1Enabled
End Sub

Private Sub Timer1_Timer()
Label2.Caption = Now
X = Combo1.Text & "/" & mon & "/" & "00" & " " & Combo4.Text & Combo5.Text & ":00"
'if the current time is the time that the user chose
'and the OK button is pressed then shut down
If (X = Now) And (Command1.Enabled = False) Then
    ExitWindowsEx 1, 0
End If
End Sub

Private Sub Timer2_Timer()
t = t - 1000
Label7.Caption = t / 1000 & " " & "seconds until shut down"
Me.Caption = t / 1000 & " " & "seconds until shut down"
If t = 0 Then
    ExitWindowsEx 1, 0
End If
End Sub

Public Sub Frame1Disabled()
Command1.Enabled = False
Command2.Enabled = True
Combo1.Enabled = False
Combo2.Enabled = False
Combo4.Enabled = False
Combo5.Enabled = False
Label3.Enabled = False
Label4.Enabled = False
Label5.Enabled = False
Label11.Enabled = False
Label7.Visible = True
Me.Caption = "Timing ShutDown"
End Sub

Public Sub Frame1Enabled()
Command1.Enabled = True
Command2.Enabled = False
Combo1.Enabled = True
Combo2.Enabled = True
Combo4.Enabled = True
Combo5.Enabled = True
Label3.Enabled = True
Label4.Enabled = True
Label5.Enabled = True
Label11.Enabled = True
Label7.Visible = False
Me.Caption = "Timing ShutDown"
End Sub

Public Sub Frame2Disabled()
Command5.Enabled = False
Command6.Enabled = True
Label8.Enabled = False
Label9.Enabled = False
Text2.Enabled = False
Label7.Visible = True
End Sub

Public Sub Frame2Enabled()
Command5.Enabled = True
Command6.Enabled = False
Label8.Enabled = True
Label9.Enabled = True
Text2.Enabled = True
Label7.Visible = False
Me.Caption = "Timing ShutDown"
End Sub


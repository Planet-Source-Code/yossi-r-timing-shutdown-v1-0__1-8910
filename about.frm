VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3480
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   3480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   3255
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2880
      Top             =   0
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   0
      ScaleHeight     =   2295
      ScaleWidth      =   3375
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4095
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   1
         Text            =   "about.frx":0000
         Top             =   2160
         Width           =   3255
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Text1.Text = "Timing ShutDown v1.0" & vbCrLf
Text1.Text = Text1.Text & "-=Made by Yossi Ramot=-" & vbCrLf & vbCrLf
Text1.Text = Text1.Text & "This is a very useful program which is made "
Text1.Text = Text1.Text & "to help you shut down your computer when you're "
Text1.Text = Text1.Text & "not near it." & " For instance, let's say you want "
Text1.Text = Text1.Text & "to defrag your hard disk at night and you want that "
Text1.Text = Text1.Text & "after finishing, the computer will turn off. Just "
Text1.Text = Text1.Text & "set the timer for how much time you think this job "
Text1.Text = Text1.Text & "will take, minimize Timing ShutDown to the tray "
Text1.Text = Text1.Text & "if you want and... "
Text1.Text = Text1.Text & "good night!"
End Sub

Private Sub Text1_GotFocus()
Picture1.SetFocus 'make sure there will be no cursor over the text box
End Sub

Private Sub Timer1_Timer()
If Text1.Top <= -3900 Then
    Text1.Top = 2300
Else
    Text1.Top = Text1.Top - 10
End If
End Sub

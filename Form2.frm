VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   Caption         =   "Communicator   -Asim Creations"
   ClientHeight    =   3600
   ClientLeft      =   1920
   ClientTop       =   1170
   ClientWidth     =   5790
   DrawStyle       =   3  'Dash-Dot
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   5790
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000C0&
      Caption         =   "In&tro."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   105
      Picture         =   "Form2.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Introduction"
      Top             =   2625
      Width           =   855
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H000000C0&
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3885
      Picture         =   "Form2.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2625
      Width           =   855
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H000000C0&
      Caption         =   "&>>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4830
      Picture         =   "Form2.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Next"
      Top             =   2625
      Width           =   855
   End
   Begin VB.CommandButton cmdAbout 
      BackColor       =   &H000000C0&
      Caption         =   "&About"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2940
      Picture         =   "Form2.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2625
      Width           =   855
   End
   Begin VB.CommandButton cmdHelp 
      BackColor       =   &H000000C0&
      Caption         =   "&Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1995
      Picture         =   "Form2.frx":1412
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2625
      Width           =   855
   End
   Begin VB.CommandButton cmdConfig 
      BackColor       =   &H000000C0&
      Caption         =   "C&onfig."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1050
      Picture         =   "Form2.frx":1854
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Configuration of Modem."
      Top             =   2625
      Width           =   855
   End
   Begin VB.OptionButton optTerminal 
      BackColor       =   &H80000007&
      Caption         =   "Terminal"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   540
      Left            =   3885
      MaskColor       =   &H00808080&
      TabIndex        =   2
      Top             =   1785
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.OptionButton optReceive 
      BackColor       =   &H80000012&
      Caption         =   "Receive Call"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   3885
      TabIndex        =   1
      Top             =   840
      Width           =   1590
   End
   Begin VB.OptionButton optCall 
      BackColor       =   &H80000008&
      Caption         =   "To Call"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   3885
      Picture         =   "Form2.frx":1C96
      TabIndex        =   0
      Top             =   315
      Width           =   1485
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Height          =   2325
      Left            =   105
      Top             =   105
      Width           =   3585
   End
   Begin VB.OLE OLE1 
      Class           =   "PowerPoint.Show.8"
      Height          =   330
      Left            =   4935
      OleObjectBlob   =   "Form2.frx":20D8
      SourceDoc       =   "C:\MY DOCUMENTS\ADIALER.PPT"
      TabIndex        =   9
      Top             =   2625
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2115
      Left            =   210
      Picture         =   "Form2.frx":8D6F0
      Stretch         =   -1  'True
      ToolTipText     =   "Asim 96-E-83,U.E.T. LHR,Pakistan."
      Top             =   210
      Width           =   3375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FlagNext As Integer
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Const SND_ASYNC = &H1
Dim Sound As Integer
Dim msg As String
Private Sub cmdAbout_Click()
Load frmAbout
frmAbout.Show
frmMain.Visible = False
End Sub

Private Sub cmdAbout_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If cmdAbout.BackColor = &HC0& Then
cmdAbout.BackColor = vbGreen
Sound = sndPlaySound("pluck_a.wav", SND_ASYNC)

End If
If Command1.BackColor = vbGreen Then
Command1.BackColor = &HC0&
ElseIf cmdConfig.BackColor = vbGreen Then
cmdConfig.BackColor = &HC0&
ElseIf cmdHelp.BackColor = vbGreen Then
cmdHelp.BackColor = &HC0&
ElseIf cmdExit.BackColor = vbGreen Then
cmdExit.BackColor = &HC0&
ElseIf cmdNext.BackColor = vbGreen Then
cmdNext.BackColor = &HC0&
End If

End Sub

Private Sub cmdConfig_Click()
frmConfig.Visible = True
frmMain.Visible = False
End Sub

Private Sub cmdConfig_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If cmdConfig.BackColor = &HC0& Then
cmdConfig.BackColor = vbGreen
Sound = sndPlaySound("pluck_a.wav", SND_ASYNC)

End If
If Command1.BackColor = vbGreen Then
Command1.BackColor = &HC0&
ElseIf cmdHelp.BackColor = vbGreen Then
cmdHelp.BackColor = &HC0&
ElseIf cmdAbout.BackColor = vbGreen Then
cmdAbout.BackColor = &HC0&
ElseIf cmdExit.BackColor = vbGreen Then
cmdExit.BackColor = &HC0&
ElseIf cmdNext.BackColor = vbGreen Then
cmdNext.BackColor = &HC0&
End If

End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If cmdExit.BackColor = &HC0& Then
cmdExit.BackColor = vbGreen
Sound = sndPlaySound("pluck_a.wav", SND_ASYNC)

End If
If Command1.BackColor = vbGreen Then
Command1.BackColor = &HC0&
ElseIf cmdConfig.BackColor = vbGreen Then
cmdConfig.BackColor = &HC0&
ElseIf cmdHelp.BackColor = vbGreen Then
cmdHelp.BackColor = &HC0&
ElseIf cmdAbout.BackColor = vbGreen Then
cmdAbout.BackColor = &HC0&
ElseIf cmdNext.BackColor = vbGreen Then
cmdNext.BackColor = &HC0&
End If

End Sub

Private Sub cmdHelp_Click()
msg = "The very first step is to change the setting in Configuration Window " + vbCr
msg = msg + "according to your modem setting in the Control Panel." + vbCr + vbCr
msg = msg + "The second step is to select one option and then click ' >> '." + vbCr
msg = msg + vbCr + "By Clicking ' Intro ' you will be able to see the introduction." + vbCr
msg = msg + vbCr + "For more information read USER MANUAL. " + vbCr + vbCr
msg = msg + "If thing, goes wrong then contact SYED ASIM HUSSAIN ABBASI" + vbCr
msg = msg + "(96-E-83 U.E.T. Lahore,Pakistan)"


MsgBox msg, vbExclamation, "Asim Creations"
End Sub

Private Sub cmdHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If cmdHelp.BackColor = &HC0& Then
cmdHelp.BackColor = vbGreen
Sound = sndPlaySound("pluck_a.wav", SND_ASYNC)

End If
If Command1.BackColor = vbGreen Then
Command1.BackColor = &HC0&
ElseIf cmdConfig.BackColor = vbGreen Then
cmdConfig.BackColor = &HC0&
ElseIf cmdAbout.BackColor = vbGreen Then
cmdAbout.BackColor = &HC0&
ElseIf cmdExit.BackColor = vbGreen Then
cmdExit.BackColor = &HC0&
ElseIf cmdNext.BackColor = vbGreen Then
cmdNext.BackColor = &HC0&
End If

End Sub

Private Sub cmdNext_Click()
FlagNext = True
Unload Me
End Sub

Private Sub cmdNext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If cmdNext.BackColor = &HC0& Then
cmdNext.BackColor = vbGreen
Sound = sndPlaySound("pluck_a.wav", SND_ASYNC)

End If
If Command1.BackColor = vbGreen Then
Command1.BackColor = &HC0&
ElseIf cmdConfig.BackColor = vbGreen Then
cmdConfig.BackColor = &HC0&
ElseIf cmdHelp.BackColor = vbGreen Then
cmdHelp.BackColor = &HC0&
ElseIf cmdAbout.BackColor = vbGreen Then
cmdAbout.BackColor = &HC0&
ElseIf cmdExit.BackColor = vbGreen Then
cmdExit.BackColor = &HC0&
End If

End Sub

Private Sub Command1_Click()
OLE1.DoVerb
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Command1.BackColor = &HC0& Then
Command1.BackColor = vbGreen
Sound = sndPlaySound("pluck_a.wav", SND_ASYNC)

End If
If cmdConfig.BackColor = vbGreen Then
cmdConfig.BackColor = &HC0&
ElseIf cmdHelp.BackColor = vbGreen Then
cmdHelp.BackColor = &HC0&
ElseIf cmdAbout.BackColor = vbGreen Then
cmdAbout.BackColor = &HC0&
ElseIf cmdExit.BackColor = vbGreen Then
cmdExit.BackColor = &HC0&
ElseIf cmdNext.BackColor = vbGreen Then
cmdNext.BackColor = &HC0&
End If

End Sub

Private Sub Form_Load()
FlagNext = False
frmConfig.Visible = False
Load frmConfig

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Command1.BackColor = vbGreen Then
Command1.BackColor = &HC0&
ElseIf cmdConfig.BackColor = vbGreen Then
cmdConfig.BackColor = &HC0&
ElseIf cmdHelp.BackColor = vbGreen Then
cmdHelp.BackColor = &HC0&
ElseIf cmdAbout.BackColor = vbGreen Then
cmdAbout.BackColor = &HC0&
ElseIf cmdExit.BackColor = vbGreen Then
cmdExit.BackColor = &HC0&
ElseIf cmdNext.BackColor = vbGreen Then
cmdNext.BackColor = &HC0&
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If optCall.Value = True And FlagNext = True Then
Load frmCall
frmCall.Show
ElseIf optReceive.Value = True Then
Load frmReceive
frmReceive.Show
ElseIf optTerminal.Value = True Then
Load frmTerminal
frmTerminal.Show
Else
End
End If

End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Command1.BackColor = vbGreen Then
Command1.BackColor = &HC0&
ElseIf cmdConfig.BackColor = vbGreen Then
cmdConfig.BackColor = &HC0&
ElseIf cmdHelp.BackColor = vbGreen Then
cmdHelp.BackColor = &HC0&
ElseIf cmdAbout.BackColor = vbGreen Then
cmdAbout.BackColor = &HC0&
ElseIf cmdExit.BackColor = vbGreen Then
cmdExit.BackColor = &HC0&
ElseIf cmdNext.BackColor = vbGreen Then
cmdNext.BackColor = &HC0&
End If

End Sub

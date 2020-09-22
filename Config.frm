VERSION 5.00
Object = "{184C2D02-FD61-11D0-8B58-000000000000}#1.0#0"; "FHMAGICC.OCX"
Begin VB.Form frmConfig 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Communicator  -Asim Creations"
   ClientHeight    =   4770
   ClientLeft      =   1545
   ClientTop       =   1200
   ClientWidth     =   6945
   BeginProperty Font 
      Name            =   "Noble"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Config.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   6945
   Visible         =   0   'False
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3570
      TabIndex        =   29
      Top             =   3675
      Width           =   2745
      Begin VB.OptionButton optSoftNo 
         Caption         =   "No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1680
         TabIndex        =   31
         Top             =   105
         Width           =   1065
      End
      Begin VB.OptionButton optSoftYes 
         Caption         =   "Yes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   105
         TabIndex        =   30
         Top             =   105
         Value           =   -1  'True
         Width           =   960
      End
   End
   Begin fhMagicControlsB1.MagicImage MagicImage1 
      Height          =   480
      Left            =   315
      TabIndex        =   27
      Top             =   105
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Style           =   3
   End
   Begin VB.OptionButton optPulse 
      Caption         =   "Pulse Dialing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5250
      TabIndex        =   24
      Top             =   3465
      Value           =   -1  'True
      Width           =   1380
   End
   Begin VB.OptionButton optTone 
      Caption         =   "Tone Dialing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3675
      TabIndex        =   23
      Top             =   3465
      Width           =   1275
   End
   Begin VB.CommandButton Command3 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3570
      TabIndex        =   21
      Top             =   4200
      Width           =   3165
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&<<"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   210
      TabIndex        =   20
      Top             =   4200
      Width           =   3165
   End
   Begin VB.ComboBox cboHandshaking 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5145
      TabIndex        =   19
      Text            =   "comRTSXOnXOff"
      Top             =   2940
      Width           =   1590
   End
   Begin VB.TextBox txtInputBSize 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5145
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   2520
      Width           =   1590
   End
   Begin VB.TextBox txtSendBSize 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5145
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   2100
      Width           =   1590
   End
   Begin VB.TextBox txtRThreshold 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5145
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1680
      Width           =   1590
   End
   Begin VB.TextBox txtSThreshold 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5145
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   1260
      Width           =   1590
   End
   Begin VB.ComboBox cboCommPort 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5145
      TabIndex        =   9
      Text            =   "4"
      Top             =   840
      Width           =   1590
   End
   Begin VB.Frame fraConnectionPre 
      Caption         =   "Connection preferences"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   210
      TabIndex        =   1
      Top             =   735
      Width           =   3375
      Begin VB.ComboBox cboBaudRates 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1260
         TabIndex        =   25
         Text            =   "115200"
         Top             =   1995
         Width           =   1590
      End
      Begin VB.ComboBox cboStopBits 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1260
         TabIndex        =   7
         Text            =   "1"
         Top             =   1470
         Width           =   1590
      End
      Begin VB.ComboBox cboParity 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1260
         TabIndex        =   5
         Text            =   "None"
         Top             =   945
         Width           =   1590
      End
      Begin VB.ComboBox cboDataBits 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1260
         TabIndex        =   3
         Text            =   "8"
         Top             =   420
         Width           =   1590
      End
      Begin VB.Label Label12 
         Caption         =   "Baud Rates :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   26
         Top             =   1995
         Width           =   960
      End
      Begin VB.Label Label4 
         Caption         =   "Stop Bits :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   420
         TabIndex        =   6
         Top             =   1470
         Width           =   750
      End
      Begin VB.Label Label3 
         Caption         =   "Parity  :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   630
         TabIndex        =   4
         Top             =   945
         Width           =   540
      End
      Begin VB.Label Label2 
         Caption         =   "Data Bits :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   420
         TabIndex        =   2
         Top             =   420
         Width           =   750
      End
   End
   Begin VB.Label Label13 
      Caption         =   "Remote Party is using the Same software :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   315
      TabIndex        =   28
      Top             =   3780
      Width           =   3060
   End
   Begin VB.Label Label11 
      Caption         =   "The phone system at this location uses   :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   315
      TabIndex        =   22
      Top             =   3465
      Width           =   2955
   End
   Begin VB.Label Label10 
      Caption         =   "Handshaking      :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3780
      TabIndex        =   18
      Top             =   2940
      Width           =   1485
   End
   Begin VB.Label Label9 
      Caption         =   "Input Buffer Size :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3780
      TabIndex        =   16
      Top             =   2520
      Width           =   1485
   End
   Begin VB.Label Label8 
      Caption         =   "Send Buffer Size :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3780
      TabIndex        =   14
      Top             =   2100
      Width           =   1485
   End
   Begin VB.Label Label7 
      Caption         =   "RThreshold  :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3780
      TabIndex        =   11
      Top             =   1680
      Width           =   1485
   End
   Begin VB.Label Label6 
      Caption         =   "SThreshold  :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3780
      TabIndex        =   10
      Top             =   1260
      Width           =   1485
   End
   Begin VB.Label Label5 
      Caption         =   "Comm Port   :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3780
      TabIndex        =   8
      Top             =   840
      Width           =   1485
   End
   Begin VB.Label Label1 
      Caption         =   $"Config.frx":0442
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   945
      TabIndex        =   0
      Top             =   105
      Width           =   5790
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
frmConfig.Visible = False
frmMain.Visible = True

End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Form_Load()
'Setting Data Bits Combo box in Configuration Form
cboDataBits.AddItem "4", 0
cboDataBits.AddItem "5", 1
cboDataBits.AddItem "6", 2
cboDataBits.AddItem "7", 3
cboDataBits.AddItem "8", 4
'Setting Parity Combo box in Configuration Form
cboParity.AddItem "None", 0
cboParity.AddItem "Odd", 1
cboParity.AddItem "Even", 2
cboParity.AddItem "Mark", 3
cboParity.AddItem "Space", 4
'Setting Stop Bits Combo box in Confoguration Form
cboStopBits.AddItem "1", 0
cboStopBits.AddItem "1.5", 1
cboStopBits.AddItem "2", 2
cboStopBits.ListIndex = 0
'Setting Baud Rates Combo Box
cboBaudRates.AddItem "110", 0
cboBaudRates.AddItem "300", 1
cboBaudRates.AddItem "1200", 2
cboBaudRates.AddItem "2400", 3
cboBaudRates.AddItem "4800", 4
cboBaudRates.AddItem "9600", 5
cboBaudRates.AddItem "19200", 6
cboBaudRates.AddItem "38400", 7
cboBaudRates.AddItem "57600", 8
cboBaudRates.AddItem "115200", 9
'Setting of Comm Port Combo box
cboCommPort.AddItem "1", 0
cboCommPort.AddItem "2", 1
cboCommPort.AddItem "3", 2
cboCommPort.AddItem "4", 3
'Setting of Text boxes
txtSThreshold.Text = "1"
txtRThreshold.Text = "1"
txtSendBSize.Text = "Default"
txtInputBSize.Text = "Default"
'Setting of Handshaking Combo box
cboHandshaking.AddItem "comNone", 0
cboHandshaking.AddItem "comXOnXOff", 1
cboHandshaking.AddItem "comRTS", 2
cboHandshaking.AddItem "comRTSXOnXOff", 3
End Sub



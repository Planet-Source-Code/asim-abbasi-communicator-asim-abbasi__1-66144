VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form frmTerminal 
   Caption         =   "Terminal   -Asim Creations"
   ClientHeight    =   5760
   ClientLeft      =   705
   ClientTop       =   750
   ClientWidth     =   8100
   Icon            =   "frmTterminal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   8100
   Begin VB.CommandButton Command1 
      Caption         =   "S&ave"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2415
      TabIndex        =   4
      Top             =   4830
      Width           =   1380
   End
   Begin VB.CommandButton cmdDisconnect 
      Caption         =   "Disconnect"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   945
      TabIndex        =   3
      Top             =   4830
      Width           =   1380
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "&<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   105
      TabIndex        =   2
      Top             =   4830
      Width           =   750
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   5385
      Width           =   8100
      _ExtentX        =   14288
      _ExtentY        =   661
      Style           =   1
      SimpleText      =   "Not Connected."
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   7056
            MinWidth        =   7056
            Picture         =   "frmTterminal.frx":0442
            Text            =   "Asim Creations"
            TextSave        =   "Asim Creations"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtTerminal 
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   4530
      Left            =   105
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   105
      Width           =   7890
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuSaveScreen 
         Caption         =   "S&ave"
      End
      Begin VB.Menu mnuSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileTransfer 
         Caption         =   "File &Transfer..."
      End
      Begin VB.Menu mnuFileReception 
         Caption         =   "File &Reception..."
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuTHelp 
         Caption         =   "&Terminal Help"
      End
      Begin VB.Menu Sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmTerminal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dummy, Snd As Integer
Dim fromModem$
Dim dummi As String
Dim filenum As Integer
Dim msg As String

Private Sub cmdBack_Click()
Unload Me
End Sub


Private Sub cmdDisconnect_Click()
mnuDisconnect_Click
End Sub



Private Sub Command1_Click()

dummi = InputBox("Enter File Name : ", "Asim Creations", "c:\windows\desktop\chat1.txt")
If dummi <> "" And Val(dummy) = 0 Then
filenum = FreeFile
On Error GoTo ErrorHandler
Open dummi For Output As filenum
Print #filenum, txtTerminal.Text
Close #filenum
Else
MsgBox "File Name Field Rejected", vbInformation, "Asim Creations"
End If
Exit Sub

ErrorHandler:
MsgBox "Invalid File Name", vbCritical, "Asim Creations"
Exit Sub

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If frmTerminal.StatusBar1.SimpleText <> "Connected." Then
Load frmCall
frmCall.Visible = False
If frmCall.MSComm1.PortOpen = False Then
    With frmCall
        .MSComm1.RThreshold = CInt(frmConfig.txtRThreshold.Text)
        .MSComm1.SThreshold = CInt(frmConfig.txtSThreshold.Text)
        .MSComm1.InputLen = 0
        .MSComm1.InputMode = comInputModeText
        .MSComm1.Handshaking = Val(frmConfig.cboHandshaking.Text)
        .MSComm1.InBufferCount = 0
        .MSComm1.OutBufferCount = 0
        .MSComm1.CommPort = CInt(frmConfig.cboCommPort.Text)
    End With
    If frmConfig.cboParity.Text = "None" Then
    frmCall.MSComm1.Settings = "115200,N,8,1"
    End If
    If frmConfig.cboParity.Text = "Odd" Then
    frmCall.MSComm1.Settings = "115200,O,8,1"
    End If
    If frmConfig.cboParity.Text = "Even" Then
    frmCall.MSComm1.Settings = "115200,E,8,1"
    End If
    If frmConfig.cboParity.Text = "Mark" Then
    frmCall.MSComm1.Settings = "115200,M,8,1"
    End If
    If frmConfig.cboParity.Text = "Space" Then
    frmCall.MSComm1.Settings = "115200,S,8,1"
    End If
On Error GoTo PortError
frmCall.MSComm1.PortOpen = True
 frmCall.MSComm1.RTSEnable = True

    With frmCall
        .MSComm1.OutBufferCount = 0
        .MSComm1.InBufferCount = 0
    End With
    StatusBar1.SimpleText = "Connected."
End If
End If
If frmCall.MSComm1.PortOpen = True Then
frmCall.MSComm1.Output = Chr$(KeyAscii)
End If
'txtTerminal.SelText = Chr$(KeyAscii)
Exit Sub

PortError:
   MsgBox "Invalid Port Number: " + vbCr + " Change Configuration setting", vbOKOnly, "Asim Error Detectiver "
cmdDisconnect_Click
Exit Sub


End Sub

Private Sub Form_Unload(Cancel As Integer)
frmCall.FlagBack = True
Unload frmCall
Load frmMain
frmMain.Show
End Sub


Public Sub mnuConnect_Click()

End Sub

Private Sub mnuDisconnect_Click()
If frmCall.MSComm1.PortOpen = True Then
  With frmCall
        .MSComm1.Output = "ATH" + vbCr
        .MSComm1.PortOpen = False
End With
  StatusBar1.SimpleText = "Disconnected ......."
End If
txtTerminal.SetFocus
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show 1
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuFileReception_Click()
frmFileReception.Show 1
End Sub

Private Sub mnuFileTransfer_Click()
frmFileTransfer.Show 1
End Sub

Private Sub mnuNew_Click()
cmdDisconnect_Click
txtTerminal.Text = ""
txtTerminal.SetFocus
End Sub


Private Sub mnuSend_Click()
End Sub

Private Sub mnuSaveScreen_Click()
Call Command1_Click
End Sub

Private Sub mnuTHelp_Click()
msg = "Essential AT commands are as follows :- " + vbCr + vbCr
msg = msg + "AT&F  For default modem setting." + vbCr
msg = msg + "ATDTphoneNumber or ATDPphoneNumber e.g. ATDT234234." + vbCr
msg = msg + "ATA  on the arrival of 'Ring' message." + vbCr + vbCr
msg = msg + "Click the 'Disconnect' Button, when want to disconnect." + vbCr + vbCr
msg = msg + "For more AT command concern your Modem Reference Guide."

MsgBox msg, vbExclamation, "Asim Creations"

End Sub

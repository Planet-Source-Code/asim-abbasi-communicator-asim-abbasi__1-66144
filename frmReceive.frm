VERSION 5.00
Begin VB.Form frmReceive 
   Caption         =   "Communicator"
   ClientHeight    =   2055
   ClientLeft      =   2370
   ClientTop       =   1245
   ClientWidth     =   4755
   Icon            =   "frmReceive.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   4755
   Begin VB.CommandButton CmdConnect 
      Caption         =   "&Connect"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1680
      TabIndex        =   5
      Top             =   1470
      Width           =   1380
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "C&ancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3150
      TabIndex        =   3
      Top             =   1470
      Width           =   1380
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "&<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   210
      TabIndex        =   2
      Top             =   1470
      Width           =   1380
   End
   Begin VB.Image Image1 
      Height          =   540
      Left            =   420
      Picture         =   "frmReceive.frx":0442
      Stretch         =   -1  'True
      Top             =   105
      Width           =   750
   End
   Begin VB.Label lblMessage 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1050
      TabIndex        =   4
      Top             =   840
      Width           =   3585
   End
   Begin VB.Label Label2 
      Caption         =   "Message :"
      Height          =   330
      Left            =   105
      TabIndex        =   1
      Top             =   840
      Width           =   750
   End
   Begin VB.Label Label1 
      Caption         =   "Receiving A Call ......"
      BeginProperty Font 
         Name            =   "Harrington"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1575
      TabIndex        =   0
      Top             =   105
      Width           =   2640
   End
End
Attribute VB_Name = "frmReceive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dummy As Integer
Dim fromModem$
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Const SND_ASYNC = &H1
Const SND_LOOP = &H8
Dim Snd As Integer
Private Sub Command1_Click()

End Sub

Private Sub cmdBack_Click()
Unload Me
End Sub

Private Sub cmdCancel_Click()
If frmCall.MSComm1.PortOpen = True Then
  With frmCall
        .MSComm1.Output = "ATH" + vbCr
        .MSComm1.PortOpen = False
End With
  lblMessage.Caption = "Disconnected ......."
End If
End Sub

Private Sub cmdConnect_Click()

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
   

    
    With frmCall
        .MSComm1.OutBufferCount = 0
        .MSComm1.InBufferCount = 0
        .MSComm1.Output = "AT&F" + vbCr
    End With

  Do
    dummy = DoEvents()
    If frmCall.MSComm1.PortOpen = True Then
    fromModem$ = fromModem$ + frmCall.MSComm1.Input
        If InStr(fromModem$, "OK") Then
        lblMessage.Caption = "OK"
    Exit Do
        End If
    End If
  Loop
frmCall.MSComm1.InBufferCount = 0
frmCall.MSComm1.Output = "AT#CLS=8#VRN=0#VLS=6S0=0" + vbCr

fromModem$ = ""
Do
    dummy = DoEvents()
     If frmCall.MSComm1.InBufferCount >= 2 Then
     fromModem$ = fromModem$ + frmCall.MSComm1.Input
       If InStr(fromModem$, "OK") Then
       Exit Do
       End If
     End If
Loop
lblMessage.Caption = "Connected...."
frmCall.MSComm1.InBufferCount = 0
fromModem$ = ""

Do
    dummy = DoEvents()
     If frmCall.MSComm1.PortOpen = True And (frmCall.MSComm1.CommEvent = comEvRing _
     Or frmCall.MSComm1.InBufferCount >= 4) Then
     fromModem$ = fromModem$ + frmCall.MSComm1.Input
       If InStr(fromModem$, "RING") Then
         If frmReceive.WindowState = 1 Then
         frmReceive.Caption = "RING"
         End If
       lblMessage.Caption = "Ringing ...."

       Exit Do
       End If
    End If
Loop
         Snd = sndPlaySound("Ringin.wav", SND_ASYNC Or SND_LOOP)

       fromModem$ = ""
        frmCall.MSComm1.InBufferCount = 0
        frmCall.MSComm1.Output = "ATA" + vbCr

            Do
                dummy = DoEvents()
                fromModem$ = fromModem$ + frmCall.MSComm1.Input
                If InStr(fromModem$, "VCON") Then
                lblMessage.Caption = fromModem$
                Exit Do
                End If
            Loop

       lblMessage.Caption = "Pick up the Headset"
       frmReceive.Caption = "Communicator"
End If
Exit Sub

PortError:
   MsgBox "Invalid Port Number: " + vbCr + " Change Configuration setting", vbOKOnly, "Asim Error Detectiver "
      cmdBack_Click
Exit Sub

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode <> vbKeyReturn Then
Snd = sndPlaySound(0&, 0)
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If frmCall.MSComm1.PortOpen = True Then
  With frmCall
        .MSComm1.Output = "ATH" + vbCr
        .MSComm1.PortOpen = False
End With
  lblMessage.Caption = "Disconnected ......."
End If
Load frmMain
frmMain.Show
frmMain.optReceive.Value = False

End Sub

VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmCall 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Communicator"
   ClientHeight    =   4230
   ClientLeft      =   3030
   ClientTop       =   1200
   ClientWidth     =   3075
   Icon            =   "ADialer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   3075
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00008000&
      Caption         =   "&<<"
      Height          =   330
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   2835
      Width           =   1065
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   2520
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   327680
      CommPort        =   4
      DTREnable       =   -1  'True
      InputMode       =   1
   End
   Begin VB.CommandButton cmdConnect 
      BackColor       =   &H00FF8080&
      Caption         =   "&Connect"
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
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1470
      Width           =   1065
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      TabIndex        =   16
      Top             =   3795
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   767
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   5715
            MinWidth        =   5715
            Picture         =   "ADialer.frx":0442
            Text            =   "Asim Creations"
            TextSave        =   "Asim Creations"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Calligraphy"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmd0 
      BackColor       =   &H00004080&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   0
      Left            =   1365
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3045
      Width           =   1590
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H000000FF&
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
      Height          =   330
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3255
      Width           =   1065
   End
   Begin VB.CommandButton cmdAddress 
      BackColor       =   &H00008000&
      Caption         =   "&Address Book"
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
      Picture         =   "ADialer.frx":075C
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1995
      Width           =   1065
   End
   Begin VB.CommandButton cmd3 
      BackColor       =   &H00004080&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   3
      Left            =   2415
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2520
      Width           =   540
   End
   Begin VB.CommandButton cmd2 
      BackColor       =   &H00004080&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   2
      Left            =   1890
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2520
      Width           =   540
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00004080&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   1
      Left            =   1365
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2520
      Width           =   540
   End
   Begin VB.CommandButton cmd6 
      BackColor       =   &H00004080&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   6
      Left            =   2415
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1995
      Width           =   540
   End
   Begin VB.CommandButton cmd5 
      BackColor       =   &H00004080&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   5
      Left            =   1890
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1995
      Width           =   540
   End
   Begin VB.CommandButton cmd4 
      BackColor       =   &H00004080&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   4
      Left            =   1365
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1995
      Width           =   540
   End
   Begin VB.CommandButton cmd9 
      BackColor       =   &H00004080&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   9
      Left            =   2415
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1470
      Width           =   540
   End
   Begin VB.CommandButton cmd8 
      BackColor       =   &H00004080&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   8
      Left            =   1890
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1470
      Width           =   540
   End
   Begin VB.CommandButton cmd7 
      BackColor       =   &H00004080&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   7
      Left            =   1365
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1470
      Width           =   540
   End
   Begin VB.TextBox txtPhone 
      Height          =   330
      Left            =   840
      TabIndex        =   1
      Top             =   840
      Width           =   2115
   End
   Begin VB.TextBox txtName 
      Height          =   330
      Left            =   840
      TabIndex        =   0
      Top             =   420
      Width           =   2115
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   1890
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   -2147483627
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   3
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ADialer.frx":0B9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ADialer.frx":0EB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "ADialer.frx":11D2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Making a call ......."
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   315
      TabIndex        =   18
      Top             =   0
      Width           =   2640
   End
   Begin VB.Label Label2 
      Caption         =   "PHONE :"
      Height          =   225
      Left            =   105
      TabIndex        =   14
      Top             =   840
      Width           =   750
   End
   Begin VB.Label Label1 
      Caption         =   "NAME :"
      Height          =   225
      Left            =   105
      TabIndex        =   13
      Top             =   420
      Width           =   645
   End
End
Attribute VB_Name = "frmCall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************

'This form is used to dial the Phone Number. It has a very special
'graphical display of differents events happening durring dialing.
'It also has a LED simulated Musical numeric keypad and address
'book to store,retrive and delete the stored data. Moreover, it has a
'search engine too.

'******************************************************************************
Dim dummy
Dim fromModem$
Public FlagBack As Boolean
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Const SND_ASYNC = &H1         '  play asynchronously
Dim Snd As Integer
Private Sub cmd0_Click(Index As Integer)
txtPhone.Text = txtPhone.Text + "0"
'API is used for playing of sound asynchronously i.e. without
'halting any task. Computer plays the sound in background.
Snd = sndPlaySound("pluck_a.wav", SND_ASYNC)

End Sub

Private Sub cmd0_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'The following procedure is used to make the LED like display.
If cmd0(0).BackColor = &H4080& Then
cmd0(0).BackColor = vbGreen
   If cmd0(0).BackColor = vbGreen Then
   cmd1(1).BackColor = &H4080&      'This constant is the code for dark brown color
   cmd2(2).BackColor = &H4080&
   cmd3(3).BackColor = &H4080&
   cmd4(4).BackColor = &H4080&
   cmd5(5).BackColor = &H4080&
   cmd6(6).BackColor = &H4080&
   cmd7(7).BackColor = &H4080&
   cmd8(8).BackColor = &H4080&
   cmd9(9).BackColor = &H4080&
   End If
End If
End Sub

Private Sub cmd1_Click(Index As Integer)
txtPhone.Text = txtPhone.Text + "1"
'for each clicking of button the sound is played. The file is located
'in current directory.
Snd = sndPlaySound("pluck_a.wav", SND_ASYNC)

End Sub

Private Sub cmd1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'The following procedure is used to make the LED like display.
If cmd1(1).BackColor = &H4080& Then
cmd1(1).BackColor = vbGreen
   If cmd1(1).BackColor = vbGreen Then
   cmd0(0).BackColor = &H4080&
   cmd2(2).BackColor = &H4080&
   cmd3(3).BackColor = &H4080&
   cmd4(4).BackColor = &H4080&
   cmd5(5).BackColor = &H4080&
   cmd6(6).BackColor = &H4080&
   cmd7(7).BackColor = &H4080&
   cmd8(8).BackColor = &H4080&
   cmd9(9).BackColor = &H4080&
   End If
End If
End Sub

Private Sub cmd2_Click(Index As Integer)
txtPhone.Text = txtPhone.Text + "2"
   Snd = sndPlaySound("pluck_a.wav", SND_ASYNC)

End Sub

Private Sub cmd2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'The following procedure is used to make the LED like display.
If cmd2(2).BackColor = &H4080& Then
cmd2(2).BackColor = vbGreen
   If cmd2(2).BackColor = vbGreen Then
   cmd1(1).BackColor = &H4080&
   cmd0(0).BackColor = &H4080&
   cmd3(3).BackColor = &H4080&
   cmd4(4).BackColor = &H4080&
   cmd5(5).BackColor = &H4080&
   cmd6(6).BackColor = &H4080&
   cmd7(7).BackColor = &H4080&
   cmd8(8).BackColor = &H4080&
   cmd9(9).BackColor = &H4080&
   End If
End If
End Sub

Private Sub cmd3_Click(Index As Integer)
txtPhone.Text = txtPhone.Text + "3"
   Snd = sndPlaySound("pluck_a.wav", SND_ASYNC)

End Sub

Private Sub cmd3_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If cmd3(3).BackColor = &H4080& Then
cmd3(3).BackColor = vbGreen
   If cmd3(3).BackColor = vbGreen Then
   cmd1(1).BackColor = &H4080&
   cmd2(2).BackColor = &H4080&
   cmd0(0).BackColor = &H4080&
   cmd4(4).BackColor = &H4080&
   cmd5(5).BackColor = &H4080&
   cmd6(6).BackColor = &H4080&
   cmd7(7).BackColor = &H4080&
   cmd8(8).BackColor = &H4080&
   cmd9(9).BackColor = &H4080&
   End If
End If
End Sub

Private Sub cmd4_Click(Index As Integer)
txtPhone.Text = txtPhone.Text + "4"
   Snd = sndPlaySound("pluck_a.wav", SND_ASYNC)

End Sub

Private Sub cmd4_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If cmd4(4).BackColor = &H4080& Then
cmd4(4).BackColor = vbGreen
   If cmd4(4).BackColor = vbGreen Then
   cmd1(1).BackColor = &H4080&
   cmd2(2).BackColor = &H4080&
   cmd3(3).BackColor = &H4080&
   cmd0(0).BackColor = &H4080&
   cmd5(5).BackColor = &H4080&
   cmd6(6).BackColor = &H4080&
   cmd7(7).BackColor = &H4080&
   cmd8(8).BackColor = &H4080&
   cmd9(9).BackColor = &H4080&
   End If
End If
End Sub

Private Sub cmd5_Click(Index As Integer)
txtPhone.Text = txtPhone.Text + "5"
   Snd = sndPlaySound("pluck_a.wav", SND_ASYNC)

End Sub

Private Sub cmd5_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If cmd5(5).BackColor = &H4080& Then
cmd5(5).BackColor = vbGreen
   If cmd5(5).BackColor = vbGreen Then
   cmd1(1).BackColor = &H4080&
   cmd2(2).BackColor = &H4080&
   cmd3(3).BackColor = &H4080&
   cmd4(4).BackColor = &H4080&
   cmd0(0).BackColor = &H4080&
   cmd6(6).BackColor = &H4080&
   cmd7(7).BackColor = &H4080&
   cmd8(8).BackColor = &H4080&
   cmd9(9).BackColor = &H4080&
   End If
End If
End Sub

Private Sub cmd6_Click(Index As Integer)
txtPhone.Text = txtPhone.Text + "6"
   Snd = sndPlaySound("pluck_a.wav", SND_ASYNC)

End Sub

Private Sub cmd6_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If cmd6(6).BackColor = &H4080& Then
cmd6(6).BackColor = vbGreen
   If cmd6(6).BackColor = vbGreen Then
   cmd1(1).BackColor = &H4080&
   cmd2(2).BackColor = &H4080&
   cmd3(3).BackColor = &H4080&
   cmd4(4).BackColor = &H4080&
   cmd5(5).BackColor = &H4080&
   cmd0(0).BackColor = &H4080&
   cmd7(7).BackColor = &H4080&
   cmd8(8).BackColor = &H4080&
   cmd9(9).BackColor = &H4080&
   End If
End If
End Sub

Private Sub cmd7_Click(Index As Integer)
txtPhone.Text = txtPhone.Text + "7"
   Snd = sndPlaySound("pluck_a.wav", SND_ASYNC)
End Sub

Private Sub cmd7_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If cmd7(7).BackColor = &H4080& Then
cmd7(7).BackColor = vbGreen
   If cmd7(7).BackColor = vbGreen Then
   cmd1(1).BackColor = &H4080&
   cmd2(2).BackColor = &H4080&
   cmd3(3).BackColor = &H4080&
   cmd4(4).BackColor = &H4080&
   cmd5(5).BackColor = &H4080&
   cmd6(6).BackColor = &H4080&
   cmd0(0).BackColor = &H4080&
   cmd8(8).BackColor = &H4080&
   cmd9(9).BackColor = &H4080&
   End If
End If
End Sub

Private Sub cmd8_Click(Index As Integer)
txtPhone.Text = txtPhone.Text + "8"
   Snd = sndPlaySound("pluck_a.wav", SND_ASYNC)

End Sub

Private Sub cmd8_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If cmd8(8).BackColor = &H4080& Then
cmd8(8).BackColor = vbGreen
   If cmd8(8).BackColor = vbGreen Then
   cmd1(1).BackColor = &H4080&
   cmd2(2).BackColor = &H4080&
   cmd3(3).BackColor = &H4080&
   cmd4(4).BackColor = &H4080&
   cmd5(5).BackColor = &H4080&
   cmd6(6).BackColor = &H4080&
   cmd7(7).BackColor = &H4080&
   cmd0(0).BackColor = &H4080&
   cmd9(9).BackColor = &H4080&
   End If
End If
End Sub

Private Sub cmd9_Click(Index As Integer)
txtPhone.Text = txtPhone.Text + "9"
   Snd = sndPlaySound("pluck_a.wav", SND_ASYNC)

End Sub

Private Sub cmd9_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If cmd9(9).BackColor = &H4080& Then
cmd9(9).BackColor = vbGreen
   If cmd9(9).BackColor = vbGreen Then
   cmd1(1).BackColor = &H4080&
   cmd2(2).BackColor = &H4080&
   cmd3(3).BackColor = &H4080&
   cmd4(4).BackColor = &H4080&
   cmd5(5).BackColor = &H4080&
   cmd6(6).BackColor = &H4080&
   cmd7(7).BackColor = &H4080&
   cmd8(8).BackColor = &H4080&
   cmd0(0).BackColor = &H4080&
   End If
End If
End Sub



Private Sub cmdAddress_Click()
'To make frmCall invisible, when Address Book button is clicked.
'Also making frmAddressB visible or loading.
If frmAddressB.Visible = False Then
frmAddressB.Visible = True
frmCall.Visible = False
Else
Load frmAddressB
frmAddressB.Show
End If
End Sub

Private Sub cmdBack_Click()
FlagBack = True
Unload frmAddressB
Unload Me
End Sub

Private Sub cmdConnect_Click()
'Check so that the user will not enter invalid number.
If Val(txtPhone.Text) = 0 Then
MsgBox "Enter Valid Number", vbOKOnly, "Asim Creations"
txtPhone.Text = ""
Exit Sub
End If
'Check so that the user will not enter invalid name.
If Val(txtName.Text) <> 0 Then
MsgBox "Enter Valid Name", vbOKOnly, "Asim Creations"
txtName.Text = ""
Exit Sub
End If
'Check so that the name or phone field empty
If txtName.Text = "" Or txtPhone.Text = "" Then
MsgBox "Name Field Empty Or Phone Field Empty", vbInformation, "Asim Creations"
    Else
'If the caption of cmdConnect is disconnect ATH command is executed and the If loop
'terminates.
    If cmdConnect.Caption = "&Disconnect" Then
    cmdConnect.Caption = "Connect"
    StatusBar1.Panels(1).Picture = ImageList1.ListImages(1).Picture
    MSComm1.Output = "ATH" + vbCr
            Do
            dummy = DoEvents()
            If MSComm1.PortOpen = True Then
            fromModem$ = fromModem$ + MSComm1.Input
                If InStr(fromModem$, "OK") Then           'Checking whether the reply from
                Exit Do                                                'contains the string OK.
                End If
            End If
        Loop
        StatusBar1.Panels(1).Text = "Disconnected ......."
    MSComm1.PortOpen = False
'On the otherhand if the caption of cmdConnect button is Connect (Default) then the
'following procedure will execute.
        Else
        'Checking either what type of setting are used in frmConfigure
        cmdConnect.Caption = "&Disconnect"
        PhoneNo$ = txtPhone.Text
        MSComm1.CommPort = CInt(frmConfig.cboCommPort.Text)
        If frmConfig.cboParity.Text = "None" Then
        MSComm1.Settings = "115200,N,8,1"
        End If
        If frmConfig.cboParity.Text = "Odd" Then
        MSComm1.Settings = "115200,O,8,1"
        End If
        If frmConfig.cboParity.Text = "Even" Then
        MSComm1.Settings = "115200,E,8,1"
        End If
        If frmConfig.cboParity.Text = "Mark" Then
        MSComm1.Settings = "115200,M,8,1"
        End If
        If frmConfig.cboParity.Text = "Space" Then
        MSComm1.Settings = "115200,S,8,1"
        End If
        On Error GoTo PortError
        MSComm1.PortOpen = True
        MSComm1.Output = "&C1&D2S7=75" + vbCr
           MSComm1.Output = "AT" + vbCr    'Activating and default setting of
        Do                                                    'modem.
            dummy = DoEvents()
            If MSComm1.PortOpen = True Then
            fromModem$ = fromModem$ + MSComm1.Input
                If InStr(fromModem$, "OK") Then           'Checking whether the reply from
                Exit Do                                                'contains the string OK.
                End If
            End If
        Loop
        MSComm1.OutBufferCount = 0
        MSComm1.InBufferCount = 0
        MSComm1.Output = "AT&F" + vbCr    'Activating and default setting of
        Do                                                    'modem.
            dummy = DoEvents()
            If MSComm1.PortOpen = True Then
            fromModem$ = fromModem$ + MSComm1.Input
                If InStr(fromModem$, "OK") Then           'Checking whether the reply from
                Exit Do                                                'contains the string OK.
                End If
            End If
        Loop
        StatusBar1.Panels(1).Text = fromModem$          'Updating of status bar text
        'updating of status bar image
        StatusBar1.Panels(1).Picture = ImageList1.ListImages(1).Picture
        'Activating the modem in To Call mode.
        MSComm1.Output = "AT#CLS=8#VRN=0#VLS=6" + vbCr
        fromModem$ = ""
        Do                                                                'Loop is used to update the string
            dummy = DoEvents()                                  'fromModem which contains the reply
            If MSComm1.PortOpen = True Then             'of modem.
            fromModem$ = fromModem$ + MSComm1.Input
                If InStr(fromModem$, "OK") Then
                Exit Do
                End If
            End If
        Loop
       StatusBar1.Panels(1).Text = "OK"                      'Updating of status bar
            If frmConfig.optPulse.Value = True Then
            MSComm1.Output = "ATDP" + PhoneNo$ + vbCr
            Else
                MSComm1.Output = "ATDT" + PhoneNo$ + vbCr
            End If
       StatusBar1.Panels(1).Text = "Dialing =" + PhoneNo$
       'For updating of status bar image ImageList control has used.
       StatusBar1.Panels(1).Picture = ImageList1.ListImages(3).Picture
       fromModem$ = ""
       Do
            dummy = DoEvents()
            If MSComm1.PortOpen = True Then
            fromModem$ = fromModem$ + MSComm1.Input
                If InStr(fromModem$, "VCON") Then
                Exit Do
                End If
            End If
       Loop
      StatusBar1.Panels(1).Text = "Pick up HeadSet "
      StatusBar1.Panels(1).Picture = ImageList1.ListImages(2).Picture
      End If
End If
Exit Sub

PortError:
   MsgBox "Invalid Port Number: " + vbCr + " Change Configuration setting", vbOKOnly, "Asim Error Detectiver "
      cmdBack_Click
Exit Sub

End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub Form_Load()
Load frmAddressB
'Updating of properties as mentioned by user in frmConfiguration.
frmAddressB.Visible = False
MSComm1.InputLen = 0
MSComm1.RThreshold = CInt(frmConfig.txtRThreshold.Text)
MSComm1.SThreshold = CInt(frmConfig.txtSThreshold.Text)
MSComm1.InputMode = comInputModeText
MSComm1.Handshaking = Val(frmConfig.cboHandshaking.Text)
MSComm1.InBufferCount = 0
MSComm1.OutBufferCount = 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (cmd0(0).BackColor = vbGreen Or _
cmd1(1).BackColor = vbGreen Or _
cmd2(2).BackColor = vbGreen Or _
cmd3(3).BackColor = vbGreen Or _
cmd4(4).BackColor = vbGreen) Or _
(cmd5(5).BackColor = vbGreen Or _
cmd6(6).BackColor = vbGreen Or _
cmd7(7).BackColor = vbGreen Or _
cmd8(8).BackColor = vbGreen Or _
cmd9(9).BackColor = vbGreen) Then
cmd0(0).BackColor = &H4080&
cmd1(1).BackColor = &H4080&
cmd2(2).BackColor = &H4080&
cmd3(3).BackColor = &H4080&
cmd4(4).BackColor = &H4080&
cmd5(5).BackColor = &H4080&
cmd6(6).BackColor = &H4080&
cmd7(7).BackColor = &H4080&
cmd8(8).BackColor = &H4080&
cmd9(9).BackColor = &H4080&
End If
End Sub

Private Sub mnuCom_Click()

End Sub

Private Sub mnuConfig_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
If FlagBack = True Then
Load frmMain
frmMain.Show
FlagBack = False
ElseIf frmCall.MSComm1.PortOpen = True Then
  With frmCall
        .MSComm1.Output = "ATH" + vbCr
        .MSComm1.PortOpen = False
End With
Load frmMain
frmMain.Show
Else
End
End If
End Sub

Private Sub MSComm1_OnComm()
' Common error are trapped in this procedure.
If MSComm1.CommEvent = comEventBreak Then
MsgBox "A Break signal was received", vbCritical, "Asim Creations"
    'When frmterminal is visible then its status bar is updated.
    If frmTerminal.Visible = True Then
    frmTerminal.StatusBar1.SimpleText = "A Break signal was received"
    End If
End If
If MSComm1.CommEvent = comEventCTSTO Then
MsgBox "Clear  To Send Timeout", vbExclamation, "Asim Creations"
    If frmTerminal.Visible = True Then
    frmTerminal.StatusBar1.SimpleText = "Clear  To Send Timeout"
    End If
End If
If MSComm1.CommEvent = comEventDSRTO Then
MsgBox "Data Set Ready Timeout", vbExclamation, "Asim Creations"
    If frmTerminal.Visible = True Then
    frmTerminal.StatusBar1.SimpleText = "Data Set Ready Timeout"
    End If
End If
If MSComm1.CommEvent = comEventFrame Then
MsgBox "Framing Error", vbCritical, "Asim Creations"
    If frmTerminal.Visible = True Then
    frmTerminal.StatusBar1.SimpleText = "Framing Error"
    End If
End If
If MSComm1.CommEvent = comEventOverrun Then
MsgBox "Port Overrun :" + vbCr + "Use Handshaking in Config", vbCritical, "Asim Creations"
    If frmTerminal.Visible = True Then
    frmTerminal.StatusBar1.SimpleText = "Port Overrun"
    End If
End If
'This error is generated when the connection to remote party drops.
If MSComm1.CommEvent = comEventCDTO Then
MsgBox "Carrier Detect Timeout", vbCritical, "Asim Creations"
    'if frmterminal is visible then Status bar text is updated
    If frmTerminal.Visible = True Then
    frmTerminal.StatusBar1.SimpleText = "Carrier Detect Timeout"
    End If
End If
If MSComm1.CommEvent = comEventDCB Then
MsgBox "Unexpected error retrieving Device Control Block for the port", vbCritical, "Asim Creations"
End If
'****************************************************************************************
'The following if loop is used when using Terminal. The following statement displays whatever
'in the receive or input buffer in txtTerminal.
If frmTerminal.Visible = True And MSComm1.PortOpen = True _
And frmFileTransfer.Visible = False _
And frmFileReception.Visible = False Then
 frmTerminal.txtTerminal.SelText = MSComm1.Input
End If

'*****************************************************************************************
End Sub


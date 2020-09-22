VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form frmFileReception 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "File Reception    -Asim Creations"
   ClientHeight    =   2595
   ClientLeft      =   2385
   ClientTop       =   1530
   ClientWidth     =   4755
   Icon            =   "frmFileReception.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdConnect 
      Caption         =   "C&onnect"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1680
      TabIndex        =   7
      Top             =   1680
      Width           =   1380
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   6
      Top             =   2265
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   582
      Style           =   1
      SimpleText      =   "Asim Creations"
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   9596
            MinWidth        =   9596
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3255
      TabIndex        =   5
      Top             =   1680
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
      TabIndex        =   4
      Top             =   1680
      Width           =   1380
   End
   Begin VB.TextBox txtDestDir 
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "c:\windows\desktop\"
      Top             =   1155
      Width           =   4530
   End
   Begin VB.TextBox txtFileArriving 
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   420
      Width           =   4530
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Destination Directory :"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      TabIndex        =   2
      Top             =   840
      Width           =   2220
   End
   Begin VB.Label lblFileArriving 
      BackStyle       =   0  'Transparent
      Caption         =   "File Arriving :"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   1380
   End
End
Attribute VB_Name = "frmFileReception"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dummy As Integer
Dim Modem$, File$
Dim AlphaNums As Integer
Dim Flage As Boolean
Dim filenum As Integer
Private Sub cmdBrowse_Click()

End Sub

Private Sub cmdBack_Click()
Flage = False
Unload Me
End Sub

Private Sub cmdCancel_Click()
Flage = False
StatusBar1.SimpleText = "Disconnected..."
End Sub

Private Sub cmdConnect_Click()
Flage = True
StatusBar1.SimpleText = "Connected.."
    Do
        dummy = DoEvents()
        If Flage = False Then Exit Do
            If (frmCall.MSComm1.PortOpen = True _
            And frmCall.MSComm1.CDHolding = True) Then
            Modem$ = Modem$ + frmCall.MSComm1.Input
            Modem$ = "CTFACS?imran.txt?AC32"
                If InStr(Modem$, "AC32") Then           'Checking whether the reply from
                AlphaNums = InStr(Modem$, "?AC32") - (InStr(Modem$, "CTFACS?") + 7)
                File$ = Mid(Modem$, InStr(Modem$, "CTFACS?") + 7, AlphaNums)
                txtFileArriving.Text = Trim(txtDestDir.Text + File$)
                frmCall.MSComm1.Output = "ACSRECEIVED"
                filenum = FreeFile
                Open txtFileArriving.Text For Binary As filenum
                frmCall.MSComm1.EOFEnable = True
                frmCall.MSComm1.InBufferCount = 0
                Do While frmCall.MSComm1.CommEvent <> comEvEOF
                Put #filenum, , frmCall.MSComm1.Input
                Loop
                Close filenum
                Exit Do
                frmCall.MSComm1.EOFEnable = False
                End If
         End If
   Loop

End Sub

Private Sub Form_Unload(Cancel As Integer)
Flage = False
End Sub

VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Begin VB.Form frmFileTransfer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Transfer    -Asim creations"
   ClientHeight    =   4860
   ClientLeft      =   2505
   ClientTop       =   780
   ClientWidth     =   4650
   Icon            =   "frmFileTransfer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   4650
   Begin RichTextLib.RichTextBox txtRich 
      Height          =   2010
      Left            =   105
      TabIndex        =   7
      Top             =   2415
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   3545
      _Version        =   327680
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmFileTransfer.frx":0442
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   6
      Top             =   4530
      Width           =   4650
      _ExtentX        =   8202
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   330
      Left            =   105
      TabIndex        =   5
      Top             =   1575
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   582
      _Version        =   327682
      Appearance      =   1
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
      Left            =   2415
      TabIndex        =   4
      ToolTipText     =   "Back"
      Top             =   945
      Width           =   1065
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "B&rowse"
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
      Left            =   1260
      TabIndex        =   3
      Top             =   945
      Width           =   1065
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
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
      Left            =   105
      TabIndex        =   2
      Top             =   945
      Width           =   1065
   End
   Begin VB.TextBox txtFilePath 
      Height          =   330
      Left            =   105
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   420
      Width           =   4425
   End
   Begin VB.Label Label1 
      Caption         =   "View :"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   105
      TabIndex        =   8
      Top             =   2100
      Width           =   1275
   End
   Begin VB.Image Image1 
      Height          =   645
      Left            =   3570
      Picture         =   "frmFileTransfer.frx":050B
      Stretch         =   -1  'True
      Top             =   840
      Width           =   960
   End
   Begin VB.Label lblSendFile 
      BackStyle       =   0  'Transparent
      Caption         =   "Sending File :"
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
      Width           =   1590
   End
End
Attribute VB_Name = "frmFileTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ptPosition As Double
Dim DataByte As String
Dim dummy As Integer
Dim FileLength
Dim FileName As String
Dim filenum As Integer
Dim fromModem$
Dim flag1 As Boolean
Private Sub cmdBack_Click()
flag1 = False
Unload Me
End Sub

Private Sub cmdBrowse_Click()
txtRich.Text = " "

frmBrowser.Show 1
End Sub

Private Sub cmdSend_Click()
ptPosition = 0
FileName = "CTFACS?" + frmBrowser.filFiles.FileName + "?AC32"
If txtFilePath.Text = "" Then
MsgBox "Enter the valid File Name, using Browser !", vbInformation, "Asim Creations"
Exit Sub
End If
If frmCall.MSComm1.PortOpen = False Or _
frmCall.MSComm1.CDHolding = False Then
MsgBox "There is no Carrier available", vbInformation, "Asim Creations"
Exit Sub
End If
frmCall.MSComm1.Output = FileName
filenum = FreeFile
cmdSend.Enabled = False
 If frmConfig.optSoftYes.Value = True Then
 flag1 = True
            Do
            dummy = DoEvents()
            If frmCall.MSComm1.PortOpen = True Then
            fromModem$ = fromModem$ + frmCall.MSComm1.Input
                If InStr(fromModem$, "ACSRECE") Or flag1 = False Then
                Exit Do
                End If
            End If
            Loop
End If
On Error GoTo IncorrectFile
Open Trim(txtFilePath.Text) For Binary As filenum

FileLength = LOF(filenum)
ProgressBar1.Max = FileLength
If StatusBar1.SimpleText <> "Sending..." Then
StatusBar1.SimpleText = "Sending..."
End If

Do While Not EOF(filenum)
    ptPosition = ptPosition + 1
If ProgressBar1.Value >= FileLength Then
ProgressBar1.Value = FileLength
Else
    ProgressBar1.Value = ptPosition
    End If
    DataByte = String(1, " ")
    Get #filenum, ptPosition, DataByte
txtRich.SelText = DataByte
frmCall.MSComm1.Output = DataByte
Loop
Close filenum
StatusBar1.SimpleText = "Sent Successfully..."
ProgressBar1.Value = 0
cmdSend.Enabled = True
Exit Sub
IncorrectFile:
MsgBox "Invalid File Name", vbCritical, "Asim Creations"
Exit Sub

End Sub

Private Sub Form_Load()
ProgressBar1.Value = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
flag1 = False
End Sub


VERSION 5.00
Begin VB.Form frmBrowser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Browser   -Asim Creations"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   5715
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
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
      TabIndex        =   11
      Top             =   3570
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "C&ancel"
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
      Left            =   4200
      TabIndex        =   10
      Top             =   3570
      Width           =   1275
   End
   Begin VB.ComboBox cboFileType 
      Height          =   315
      Left            =   105
      TabIndex        =   5
      Top             =   3150
      Width           =   2325
   End
   Begin VB.TextBox txtFileName 
      Height          =   330
      Left            =   105
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   420
      Width           =   2325
   End
   Begin VB.DirListBox dirDirectory 
      Height          =   1890
      Left            =   2625
      TabIndex        =   2
      Top             =   840
      Width           =   2850
   End
   Begin VB.DriveListBox drvDrive 
      Height          =   315
      Left            =   2625
      TabIndex        =   1
      Top             =   3150
      Width           =   2850
   End
   Begin VB.FileListBox filFiles 
      Height          =   1845
      Left            =   105
      TabIndex        =   0
      Top             =   840
      Width           =   2325
   End
   Begin VB.Label lblDirName 
      Height          =   225
      Left            =   2625
      TabIndex        =   9
      Top             =   420
      Width           =   3060
   End
   Begin VB.Label lblDirectories 
      Caption         =   "Directories :"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2625
      TabIndex        =   8
      Top             =   105
      Width           =   1590
   End
   Begin VB.Label Label2 
      Caption         =   "Drive :"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2835
      TabIndex        =   7
      Top             =   2835
      Width           =   1380
   End
   Begin VB.Label lblFileType 
      Caption         =   "File Type :"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   105
      TabIndex        =   6
      Top             =   2835
      Width           =   1275
   End
   Begin VB.Label lblFileName 
      Caption         =   "File Name :"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   105
      TabIndex        =   4
      Top             =   105
      Width           =   1485
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PathandName, Path As String
Dim Response As Integer
Private Sub cboFileType_Click()
Select Case cboFileType.ListIndex
Case 0
filFiles.Pattern = "*.*"
Case 1
filFiles.Pattern = "*.txt"
Case 2
filFiles.Pattern = "*.Doc"
Case 3
filFiles.Pattern = "*.jpg"
Case 4
filFiles.Pattern = "*.bmp"
Case 5
filFiles.Pattern = "*.ACs"
End Select
txtFileName.text = ""
End Sub

Private Sub cmdCancel_Click()
frmBrowser.Visible = False
End Sub

Private Sub cmdOK_Click()
If txtFileName.text = "" Then
MsgBox "You must first select a file !", vbInformation, "Asim Creations"
Exit Sub
End If
If Right(filFiles.Path, 1) <> "\" Then
Path = filFiles.Path + "\"
Else
Path = filFiles.Path
End If
PathandName = Path + filFiles.FileName
If frmFileTransfer.Visible = True Then
frmFileTransfer.txtFilePath = PathandName
End If
'Response = MsgBox("Send the following file ?" + vbCr + PathandName _
', vbOKCancel, "Asim Creations")
'If Response = vbOK Then
'    If frmCall.MSComm1.PortOpen = True And _
'    frmCall.MSComm1.CDHolding = True Then
'    frmCall.MSComm1.Output = PathandName
'    frmTerminal.StatusBar1.SimpleText = "Sending " + PathandName
'    Else
'    MsgBox "There is no Carrier available", vbInformation, "Asim Creations"
'    End If
'End If
frmFileTransfer.StatusBar1.SimpleText = "Ready to send..."

cmdCancel_Click
End Sub

Private Sub dirDirectory_Change()
filFiles.Path = dirDirectory.Path
lblDirName.Caption = dirDirectory.Path
txtFileName.text = ""

End Sub

Private Sub drvDrive_Change()
On Error GoTo DriveError
txtFileName.text = ""
dirDirectory.Path = drvDrive.Drive
Exit Sub

DriveError:
MsgBox "Drive Error ! ", vbExclamation, "Asim Creations"
drvDrive.Drive = dirDirectory.Path
txtFileName.text = ""
Exit Sub

End Sub

Private Sub filFiles_Click()
txtFileName.text = filFiles.FileName

End Sub

Private Sub Form_Load()
cboFileType.AddItem "All files    (*.*)"
cboFileType.AddItem "Text files  (*.txt)"
cboFileType.AddItem "Doc files   (*.Doc)"
cboFileType.AddItem "Jpg files    (*.jpg)"
cboFileType.AddItem "Bmp files  (*.bmp)"
cboFileType.AddItem "ACs files    (*.ACs)"
cboFileType.ListIndex = 0
lblDirName.Caption = dirDirectory.Path
txtFileName.text = ""
End Sub

Private Sub ShellFolderViewOC1_SelectionChanged()

End Sub

Private Sub Medview1_HotspotClicked(ByVal HotspotType As Long, ByVal HotspotData As Long, Cancel As Boolean)

End Sub


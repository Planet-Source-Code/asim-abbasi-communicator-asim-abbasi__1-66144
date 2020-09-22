VERSION 5.00
Begin VB.Form frmToCall 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "To Call   -Asim Creations"
   ClientHeight    =   1440
   ClientLeft      =   2580
   ClientTop       =   1890
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   4680
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
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
      Left            =   2520
      TabIndex        =   3
      Top             =   840
      Width           =   1485
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "C&onnect"
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
      Left            =   525
      TabIndex        =   2
      Top             =   840
      Width           =   1485
   End
   Begin VB.TextBox txtPhone 
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1155
      TabIndex        =   0
      Top             =   315
      Width           =   3165
   End
   Begin VB.Label Label1 
      Caption         =   "Phone No :"
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
      TabIndex        =   1
      Top             =   315
      Width           =   1065
   End
End
Attribute VB_Name = "frmToCall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fromModem$
Dim dummy As Integer
Private Sub cmdCancel_Click()
If frmCall.MSComm1.PortOpen = True Then
  With frmCall
        .MSComm1.Output = "ATH" + vbCr
        .MSComm1.PortOpen = False
End With
  frmTerminal.StatusBar1.SimpleText = "Disconnected ......."
End If
txtPhone.SetFocus
End Sub

Private Sub cmdConnect_Click()
If txtPhone.text = "" Then
MsgBox "Phone Number field empty !", vbExclamation, "Asim Creations"
txtPhone.SetFocus
Exit Sub
End If

If Val(txtPhone.text) = 0 Then
MsgBox "Enter Valid Number", vbOKOnly, "Asim Creations"
txtPhone.text = ""
txtPhone.SetFocus
Exit Sub
End If
frmTerminal.mnuConnect_Click
 If frmCall.MSComm1.PortOpen = True Then
     With frmCall
        .MSComm1.OutBufferCount = 0
        .MSComm1.InBufferCount = 0
        .MSComm1.Output = "AT" + vbCr
    End With

  Do
    dummy = DoEvents()
    If frmCall.MSComm1.PortOpen = True Then
    fromModem$ = fromModem$ + frmCall.MSComm1.Input
        If InStr(fromModem$, "OK") Then
    Exit Do
        End If
    End If
  Loop
  
   fromModem$ = ""
frmCall.MSComm1.InBufferCount = 0
If frmConfig.optPulse.Value = True Then
frmCall.MSComm1.Output = "ATDP" + txtPhone.text + vbCr
Else
frmCall.MSComm1.Output = "ATDT" + txtPhone.text + vbCr
End If
   Do
                dummy = DoEvents()
                If frmCall.MSComm1.PortOpen = True Then
                fromModem$ = fromModem$ + frmCall.MSComm1.Input
                If InStr(fromModem$, "VCON") Then
                frmTerminal.StatusBar1.SimpleText = "Start Communicating Text..."
                Exit Do
                End If
               End If
   Loop
End If
frmToCall.Show
frmToCall.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
If frmCall.MSComm1.PortOpen = True Then
  With frmCall
        .MSComm1.Output = "ATH" + vbCr
        .MSComm1.PortOpen = False
End With
  frmTerminal.StatusBar1.SimpleText = "Disconnected......."
End If
End Sub

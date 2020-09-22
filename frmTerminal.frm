VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H80000012&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About :  ""Communicator"""
   ClientHeight    =   4545
   ClientLeft      =   2070
   ClientTop       =   780
   ClientWidth     =   5670
   Icon            =   "frmTerminal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   5670
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "This software is dedicated to my Father."
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   540
      Left            =   420
      TabIndex        =   5
      Top             =   2520
      Width           =   4845
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmTerminal.frx":0442
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   420
      TabIndex        =   4
      Top             =   3255
      Width           =   4845
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Student of B.Sc. Engg (Electronics) 96-E-83, U.E.T. Lahore, Pakistan. E-Mail: abfdani@pol.com.pk"
      ForeColor       =   &H0000FFFF&
      Height          =   645
      Left            =   1995
      TabIndex        =   3
      Top             =   1575
      Width           =   2640
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "SYED ASIM HUSSAIN ABBASI"
      BeginProperty Font 
         Name            =   "System"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   945
      TabIndex        =   2
      Top             =   1155
      Width           =   3690
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Created by"
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   2310
      TabIndex        =   1
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "COMMUNICATOR"
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   645
      Left            =   315
      TabIndex        =   0
      Top             =   105
      Width           =   5055
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Unload(Cancel As Integer)
If frmTerminal.Visible = False Then
frmMain.Visible = True
End If
End Sub


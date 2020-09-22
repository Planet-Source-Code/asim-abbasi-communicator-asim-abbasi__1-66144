VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form frmAddressB 
   Caption         =   "Address Book  -Asim Creations"
   ClientHeight    =   4065
   ClientLeft      =   2190
   ClientTop       =   1320
   ClientWidth     =   4860
   Icon            =   "frmAddressB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   4860
   Begin VB.CommandButton cmdSave 
      Caption         =   "Sa&ve"
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
      Left            =   1050
      TabIndex        =   13
      Top             =   3255
      Width           =   855
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   3690
      Width           =   4860
      _ExtentX        =   8573
      _ExtentY        =   661
      Style           =   1
      SimpleText      =   "Asim Creations"
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Geometr231 Hv BT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "&<<"
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
      Left            =   3885
      TabIndex        =   11
      Top             =   3255
      Width           =   855
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "D&elete"
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
      Left            =   2940
      TabIndex        =   10
      Top             =   3255
      Width           =   855
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
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
      Left            =   1995
      TabIndex        =   9
      Top             =   3255
      Width           =   855
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
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
      TabIndex        =   8
      Top             =   3255
      Width           =   855
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "&Previous"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   105
      Picture         =   "frmAddressB.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1680
      Width           =   1065
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Nex&t"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   105
      Picture         =   "frmAddressB.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2415
      Width           =   1065
   End
   Begin VB.TextBox txtNotes 
      Height          =   1800
      Left            =   1365
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1260
      Width           =   3375
   End
   Begin VB.TextBox txtPhone 
      Height          =   330
      Left            =   1365
      TabIndex        =   2
      Top             =   735
      Width           =   3375
   End
   Begin VB.TextBox txtName 
      Height          =   345
      Left            =   1365
      TabIndex        =   0
      Top             =   210
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Name :"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   105
      TabIndex        =   7
      Top             =   210
      Width           =   1170
   End
   Begin VB.Label Label3 
      Caption         =   "Notes :"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   105
      TabIndex        =   3
      Top             =   1260
      Width           =   1170
   End
   Begin VB.Label Label2 
      Caption         =   "Phone :"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   105
      TabIndex        =   1
      Top             =   735
      Width           =   1170
   End
End
Attribute VB_Name = "frmAddressB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************

'This form is a Address Book. It is used to store permanently
'the name, phone numbers, notes fields. This form appears when the
'user will click the Address Book button on frmCall.

'****************************************************************************

Option Explicit
Dim Person As PersonInfo  'This variable is defined in the Module1
Dim FileNo As Integer
Dim RecordLen As Long
Dim CurrentRecord As Long
Dim LastRecord As Long
Private Sub cmdBack_Click()
'When the user will click "<<" button, if the
'Name field and the Phone field contain text then it will be shifted to frmCall
'as name and phone number.
If txtName.text <> "" Or txtPhone.text <> "" Then
frmCall.txtName.text = Trim(txtName.text)
frmCall.txtPhone.text = Trim(txtPhone.text)
SaveCurrentRecord
End If
frmCall.Visible = True            'Moving back to frmCall
frmCall.txtPhone.SetFocus
Close #FileNo                       'Opened file must be closed before unloading
Unload Me
End Sub

Private Sub cmdDelete_Click()
'This procedure will be activated when user click Delete button on
'frmAddressB. This procedure has the ability to delete the current visible
'record.
Dim TmpFileNo As Integer   'file no of temporary opened file
Dim TmpRecord As Long
Dim RecNo As Long
'Message box used for the confirmation before deletion
If MsgBox("Delete this record ? ", vbYesNo, "Asim Creations") = vbNo Then
txtName.SetFocus
Exit Sub
End If
'Making sure that TmpPhone.ACs is not in the current directory. If it
'is there then it will be deleted.
If Dir("TmpPhone.ACs") = "TmpPhone.ACs" Then
Kill "TmpPhone.ACs"
End If
TmpFileNo = FreeFile    'alloting the free file no
TmpRecord = 1
RecNo = 1
Open "TmpPhone.ACs" For Random As TmpFileNo Len = RecordLen
Do While TmpRecord < LastRecord + 1
    If TmpRecord <> CurrentRecord Then      'All the records are shifted
    Get #FileNo, TmpRecord, Person          'to TmpFileNo except currently visible recored.
    Put #TmpFileNo, RecNo, Person
    RecNo = RecNo + 1
    End If
    TmpRecord = TmpRecord + 1
Loop
Close #FileNo                                           'Closing of Phone.ACs
Kill "Phone.ACs"                                       'file is closed, then deleted
Close #TmpFileNo                                     'The file in which the record is shifted except current recored, is
Name "TmpPhone.ACs" As "Phone.ACs"    'now closed and renamed.
FileNo = FreeFile
Open "Phone.ACs" For Random As FileNo Len = RecordLen
LastRecord = LastRecord - 1
If LastRecord = 0 Then LastRecord = 1
If CurrentRecord > LastRecord Then
CurrentRecord = LastRecord
End If
ShowCurrentRecord
txtName.SetFocus
End Sub

Private Sub cmdNew_Click()
CurrentRecord = LastRecord
ShowCurrentRecord

txtName.text = ""
txtPhone.text = ""
txtNotes.text = ""
If frmAddressB.Visible = True Then
txtName.SetFocus
End If
End Sub

Private Sub cmdNext_Click()
'Message box appears when the user keeps on clicking and the
'last record appears.
If CurrentRecord < LastRecord Then
CurrentRecord = CurrentRecord + 1
Else
MsgBox "Last Record !", vbExclamation, "Asim Creations"
End If
ShowCurrentRecord
txtName.SetFocus

End Sub

Private Sub cmdPrevious_Click()
'Message box appears when the user keeps on clicking and the
'first record appears.
If CurrentRecord > 1 Then
CurrentRecord = CurrentRecord - 1
Else
MsgBox "First Record !", vbExclamation, "Asim Creations"
End If
ShowCurrentRecord
txtName.SetFocus
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdSave_Click()
SaveCurrentRecord
End Sub

Private Sub cmdSearch_Click()
'This procedure is used to search the record by name.
Dim NameToSearch As String
Dim Found As Boolean
Dim tpRecNo As Long
NameToSearch = InputBox("Search for :", "Asim Creations")
If NameToSearch = "" Then       'Check if the user donot type any thing then
txtName.SetFocus                   'quit the procedure.
Exit Sub
End If
NameToSearch = UCase(NameToSearch)    'Converting all letters into upper case
Found = False                                           'Flag is used for checking either name, found or not.
For tpRecNo = 1 To LastRecord
 Get #FileNo, tpRecNo, Person
 If NameToSearch = UCase(Trim(Person.Name)) Then
 Found = True
 Exit For
 End If
Next
 If Found = True Then
 CurrentRecord = tpRecNo
 ShowCurrentRecord
 Else
 MsgBox "Name :" + NameToSearch + " not fount !", vbOKOnly, "Asim Creations"
 End If
 txtName.SetFocus
 End Sub

Private Sub Form_Load()
RecordLen = Len(Person)     'Finding the length of record
FileNo = FreeFile
Open "Phone.ACs" For Random As FileNo Len = RecordLen
CurrentRecord = 1
LastRecord = FileLen("phone.ACs") / RecordLen
If LastRecord = 0 Then
LastRecord = 1
End If
ShowCurrentRecord
End Sub

Private Sub ShowCurrentRecord()
Get #FileNo, CurrentRecord, Person
txtName.text = Person.Name
txtPhone.text = Person.Phone
txtNotes.text = Person.Notes
StatusBar1.SimpleText = "Record " + Str(CurrentRecord) + "/" + Str(LastRecord)
End Sub

Private Sub SaveCurrentRecord()
If CurrentRecord = LastRecord And (txtName.text <> "" And txtPhone.text <> "") Then
Person.Name = txtName.text
Person.Phone = txtPhone.text
Person.Notes = txtNotes.text
LastRecord = LastRecord + 1
CurrentRecord = CurrentRecord + 1
Put #FileNo, CurrentRecord, Person
StatusBar1.SimpleText = "Record " + Str(CurrentRecord) + "/" + Str(LastRecord)
End If
cmdNew_Click
End Sub


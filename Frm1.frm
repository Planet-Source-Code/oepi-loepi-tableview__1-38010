VERSION 5.00
Begin VB.Form Frm1 
   Caption         =   "Select database"
   ClientHeight    =   4485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   ScaleHeight     =   4485
   ScaleWidth      =   5610
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   4080
      Width           =   1815
   End
   Begin VB.DirListBox Dir1 
      Height          =   2565
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   5415
   End
   Begin VB.FileListBox File1 
      Height          =   870
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   5415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Select"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   4080
      Width           =   2055
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "Frm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    MDIForm1.Datab = Dir1.Path & "\" & File1.filename
    MDIForm1.path1 = Dir1.Path
    Unload Me
    MDIForm1.herladen_Click
End Sub
Private Sub Command2_Click()
    Unload Me
End Sub
Private Sub Drive1_Change()
    On Error Resume Next
    Dir1.Path = Drive1.Drive    ' When drive changes, set directory path.
End Sub
Private Sub Dir1_Change()
    On Error Resume Next
    File1.Path = Dir1.Path  ' When directory changes, set file path.
    If File1.ListCount > 0 Then
        Command1.Enabled = True
    End If
End Sub
Private Sub File1_dblClick()
    Command1_Click
End Sub
Private Sub Form_Load()
    File1.Pattern = "*.mdb" 'set pattern
    Dir1.Path = MDIForm1.path1
    Command1.Enabled = False
    If File1.ListCount Then
        Command1.Enabled = True
    End If
End Sub

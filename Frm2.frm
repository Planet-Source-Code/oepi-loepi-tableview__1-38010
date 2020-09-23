VERSION 5.00
Begin VB.Form Frm2 
   Caption         =   "Form2"
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   600
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3150
   ScaleWidth      =   600
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   0
      MultiSelect     =   1  'Simple
      TabIndex        =   0
      ToolTipText     =   "Click on the field to see possible relationships between other tables"
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "Frm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub List1_dblClick()

Dim tmp, tmp2 As String
Dim mypos, i, t As Integer

tmp = List1.Text 'get tablename from list
mypos = False


'get rid of the index marker
mypos = InStr(1, UCase(tmp), UCase(" ****"))
If mypos Then
    tmp = Left(tmp, mypos - 1)
End If


'get rid of the selections found earlier
For i = 1 To Forms.Count - 1
    For t = 0 To Forms(i).List1.ListCount - 1
        Forms(i).List1.ListIndex = t
            Forms(i).List1.Selected(t) = False
    Next t
Next i

'find text from selected table in all other tables and
'mark them
For i = 1 To Forms.Count - 1
    For t = 0 To Forms(i).List1.ListCount - 1
        Forms(i).List1.ListIndex = t
        tmp2 = Forms(i).List1.Text
        If InStr(1, UCase(tmp2), UCase(tmp)) Then
            Forms(i).List1.Selected(t) = True
        End If
    Next t
Next i

End Sub

VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Tableview"
   ClientHeight    =   5925
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu Scherm 
      Caption         =   "Window"
      WindowList      =   -1  'True
   End
   Begin VB.Menu herladen 
      Caption         =   "Reload"
   End
   Begin VB.Menu Database 
      Caption         =   "Database"
   End
   Begin VB.Menu About 
      Caption         =   "About"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xpos
Dim ypos
Public Datab
Public path1
Public Sub filenew(cap As String, ind As String)
    'make a new form based on frm2
    Dim frmnew As New Frm2
    frmnew.Height = 1200
    frmnew.Width = 1500
    frmnew.Show
    frmnew.Caption = cap
    frmnew.Width = (Len(frmnew.Caption) * 100) + 200
    frmnew.List1.Width = frmnew.Width - 200
    frmnew.Left = xpos
    frmnew.Top = ypos
End Sub
Private Sub GetTableList()

    On Error GoTo FTLErr
  
    Dim gdbCurrentDB As Database
    Dim list As ListBox
    Dim i As Integer
    Dim sTmp As String
    Dim tbl As TableDef
    Dim qdf As QueryDef
    Dim fld As Field
    Dim idx As Index
    
    i = 0
    Set gdbCurrentDB = OpenDatabase(Datab)
    
    'add the tabledefs
    For Each tbl In gdbCurrentDB.TableDefs
            sTmp = tbl.Name
            If (gdbCurrentDB.TableDefs(sTmp).Attributes And dbSystemObject) = 0 Then
                i = i + 1
                'make a new form based on tablename
                filenew sTmp, sTmp
                'get the fields in current table
                For Each fld In tbl.Fields
                    sTmp = fld.Name
                    'for each field, see if it is indexed
                    'if so mark with asterixes
                    For Each idx In tbl.Indexes
                            If "+" & sTmp = idx.Fields Then
                                sTmp = sTmp & " ****"
                            End If
                    Next idx
                    'list into active list
                    ActiveForm.List1.AddItem sTmp
                    If (Len(sTmp) * 100) + 200 > ActiveForm.Width Then
                        ActiveForm.List1.Width = Len(sTmp) * 100 + 50
                        ActiveForm.Width = ActiveForm.List1.Width + 200
                    End If
                    
                    'modify the list so all items are visible
                    ActiveForm.List1.Height = (tbl.Fields.Count * 200) + 200
                    ActiveForm.Height = ActiveForm.List1.Height + 600
                Next
                
                'if many tables, than divide over screen
                xpos = xpos + ActiveForm.Width + 200
                If xpos > MDIForm1.Width - ActiveForm.Width - 300 Then
                    ypos = ypos + 3000
                    xpos = 0
                End If
            End If
        Next
    gdbCurrentDB.Close
    Exit Sub
  
FTLErr:

End Sub

Private Sub About_Click()
Frmabout.Show vbModal
End Sub

Private Sub Database_Click()
'get a new database from user input
    Frm1.Show vbModal
End Sub

Public Sub herladen_Click()
'start built again
Dim i As Integer
For i = 1 To Forms.Count - 1
    Forms(i).Hide
Next i
xpos = 0
ypos = 0
GetTableList
End Sub

Private Sub mdiForm_Load()
On Local Error GoTo FLErr
   'fill the table list
   Datab = App.Path & "\nwind.mdb"
   GetTableList
Exit Sub
FLErr:

End Sub


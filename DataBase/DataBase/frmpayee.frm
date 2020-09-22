VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmpayee 
   Caption         =   "Payee Details"
   ClientHeight    =   4230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8430
   Icon            =   "frmpayee.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   8430
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame commandframe 
      Height          =   1695
      Left            =   6960
      TabIndex        =   9
      Top             =   0
      Width           =   1335
      Begin VB.CommandButton cmdcancel 
         Caption         =   "&Cancel"
         Height          =   325
         Left            =   180
         TabIndex        =   13
         Top             =   920
         Width           =   1000
      End
      Begin VB.CommandButton cmdsave 
         Caption         =   "&Save New"
         Enabled         =   0   'False
         Height          =   325
         Left            =   180
         TabIndex        =   12
         Top             =   200
         Width           =   1000
      End
      Begin VB.CommandButton cmdclose 
         Caption         =   "&Close"
         Height          =   325
         Left            =   180
         TabIndex        =   10
         Top             =   1280
         Width           =   1000
      End
      Begin VB.CommandButton cmdupdate 
         Caption         =   "&Update"
         Enabled         =   0   'False
         Height          =   325
         Left            =   180
         TabIndex        =   11
         Top             =   560
         Width           =   1000
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data Entry For Payee's"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   6615
      Begin VB.TextBox txtnameinara 
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   2520
         TabIndex        =   2
         Top             =   1080
         Width           =   2775
      End
      Begin VB.TextBox txtcode 
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   0
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtnameineng 
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   2520
         TabIndex        =   1
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Code ßæÏ "
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Name in English ÇáÇÓã ÈÇáÇäÌáíÒí"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   2280
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Name in Arabic ÇáÇÓã ÈÇáÚÑÈí"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   2055
      End
   End
   Begin MSComctlLib.ListView payeelist 
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   873
      SortKey         =   1
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arabic Transparent"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Payee Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Payee Name in English"
         Object.Width           =   5468
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Payee Name in Arabic"
         Object.Width           =   5468
      EndProperty
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "View / Edit Payee Details"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   1995
   End
   Begin VB.Menu pop1 
      Caption         =   "rightmouse"
      Visible         =   0   'False
      Begin VB.Menu editdata 
         Caption         =   "Edit Data"
      End
      Begin VB.Menu refresh 
         Caption         =   "Refresh"
      End
      Begin VB.Menu dash 
         Caption         =   "-"
      End
      Begin VB.Menu delete 
         Caption         =   "Delete"
      End
   End
End
Attribute VB_Name = "frmpayee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public savecon As Boolean
Public oldcode As String
Dim recpayee As New ADODB.Recordset
Dim CON1 As New ADODB.Connection
Dim myclass As New HabitatClass
Dim sqltable As Boolean
Dim xtable As String

Private Sub cmdcancel_Click()
If Trim(txtcode.Text) = "" Then
    For Each Control In Me
        If TypeOf Control Is TextBox Then
            Control.Text = ""
        End If
    Next
    savecon = True
    cmdclose.Enabled = True
    cmdsave.Enabled = False
    cmdUpdate.Enabled = False
    txtcode.SetFocus
    Exit Sub
End If

If MsgBox("Are You Sure You Want to Cancel", vbQuestion + vbYesNo, "Cancel") = vbYes Then
For Each Control In Me
    If TypeOf Control Is TextBox Then
        Control.Text = ""
    End If
Next
savecon = True
cmdclose.Enabled = True
cmdsave.Enabled = False
cmdUpdate.Enabled = False
txtcode.SetFocus
End If
End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub Command2_Click()
End Sub

Private Sub cmdUpdate_Click()

If recpayee.BOF = False Then
    recpayee.MoveFirst
    Do While recpayee.EOF = False
        If UCase(Trim(recpayee!payeecode)) = UCase(Trim(txtcode.Text)) And UCase(Trim(oldcode)) <> UCase(Trim(txtcode.Text)) Then
            MsgBox "Please Check Your Payee code", vbInformation, "Code Error"
            savecon = False
            txtcode.SetFocus
            Exit Sub
        End If
        recpayee.MoveNext
    Loop
End If

recpayee.MoveFirst
While recpayee.EOF = False
    If Trim(oldcode) = Trim(recpayee!payeecode) Then
        recpayee!payeecode = UCase(Trim(txtcode.Text))
        recpayee!payeenameara = Trim(txtnameinara.Text)
        recpayee!payeenameeng = Trim(txtnameineng.Text)
        recpayee.Update
        MsgBox "Your Data Updated Successfully", vbInformation, "Saved"
        okkalas = 1
        savecon = True 'this is identifier from listview
        cmdclose.Enabled = True
        cmdsave.Enabled = False
        cmdUpdate.Enabled = False
        txtcode.Text = ""
        txtnameinara.Text = ""
        txtnameineng.Text = ""
        txtcode.SetFocus
        Call addlist
        Exit Sub
    End If
    recpayee.MoveNext
Wend
If okkalas <> 1 Then
    MsgBox "Please Refresh Your Data and Try Again", vbInformation, "Data Error"
    Exit Sub
End If
End Sub

Private Sub cmdsave_Click()
'code for save button
'to serch for the repeating code
If recpayee.BOF = False Then
    recpayee.MoveFirst
    While recpayee.EOF = False
        If UCase(Trim(recpayee!payeecode)) = UCase(Trim(txtcode.Text)) Then
            MsgBox "Please Check Your Payee Code", vbInformation, "Data Repeating"
            txtcode.SetFocus
            Exit Sub
        End If
        recpayee.MoveNext
    Wend
End If
        
recpayee.AddNew
recpayee!payeecode = UCase(Trim(txtcode.Text))
recpayee!payeenameara = Trim(txtnameinara.Text)
recpayee!payeenameeng = Trim(txtnameineng.Text)
recpayee.Update
MsgBox "Your Data Saved Successfully", vbInformation, "Data Saved"
cmdsave.Enabled = False
cmdUpdate.Enabled = False
'cmdnew.Enabled = True
cmdclose.Enabled = True
For Each Control In Me
    If TypeOf Control Is TextBox Then
        Control.Text = ""
    End If
Next
Call addlist
txtcode.SetFocus
End Sub


Private Sub delete_Click()
If MsgBox("Are You Sure You Want to Delete This Item", vbQuestion + vbYesNo, "Conformation Deletion") = vbYes Then
    deleteitem = Trim(payeelist.SelectedItem.Text)
    recpayee.MoveFirst
        Do While recpayee.EOF = False
            If deleteitem = Trim(recpayee!payeecode) Then
                   recpayee.Delete
                Exit Do
            End If
            recpayee.MoveNext
        Loop
    txtcode.Text = ""
    txtnameinara.Text = ""
    txtnameineng.Text = ""
    savecon = True
    Call addlist
    End If
End Sub

Private Sub editdata_Click()
oldcode = payeelist.SelectedItem.Text
'MsgBox oldcode

savecon = False
txtcode.SetFocus
For Each Control In Me
    If TypeOf Control Is TextBox Then
        Control.Text = ""
    End If
Next
txtcode.Text = payeelist.SelectedItem.Text
txtnameinara.Text = payeelist.SelectedItem.ListSubItems(2).Text
txtnameineng.Text = payeelist.SelectedItem.ListSubItems(1).Text
End Sub

Private Sub Form_Activate()
txtcode.SetFocus
savecon = True

End Sub

Private Sub Form_Load()
xtable = "Select * from payee"
sqltable = True
myclass.GetTables recpayee, CON1, xtable, constring, sqltable
Call addlist

End Sub

Private Sub Form_Resize()
On Error Resume Next
payeelist.Height = Me.Height - 2600
payeelist.Width = Me.Width - 360
'cmdclose.Top = Me.Height - 850
'cmdclose.Left = Me.Width - 1450
commandframe.Left = Me.Width - 1600
Frame1.Width = commandframe.Left - 300
payeelist.ColumnHeaders(2).Width = (payeelist.Width - 1470) / 2
payeelist.ColumnHeaders(3).Width = (payeelist.Width - 1470) / 2
txtnameinara.Width = Me.Width - 5775
txtnameineng.Width = Me.Width - 5775
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Frame1.Enabled = True Then
If Trim(txtcode.Text) <> "" Or Trim(txtnameinara.Text) <> "" Or Trim(txtnameineng.Text) <> "" Then
mess = MsgBox("Do you want to discard your entries?", vbYesNo + vbQuestion, "PLease confirm")
If mess = vbYes Then
    Unload Me
  Else
    Cancel = -1
End If
End If
End If
End Sub

Private Sub payeelist_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
payeelist.SortKey = ColumnHeader.Index - 1
payeelist.Sorted = True
End Sub

Private Sub payeelist_DblClick()
If payeelist.ListItems.Count > 0 Then
oldcode = payeelist.SelectedItem.Text

'MsgBox oldcode
For Each Control In Me
    If TypeOf Control Is TextBox Then
        Control.Text = ""
    End If
Next
txtcode.Text = payeelist.SelectedItem.Text
txtnameineng.Text = payeelist.SelectedItem.ListSubItems(1).Text
txtnameinara.Text = payeelist.SelectedItem.ListSubItems(2).Text
savecon = False
txtcode.SetFocus
End If
End Sub

Private Sub payeelist_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    payeelist_DblClick
End If
End Sub

Private Sub payeelist_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = 93 Then
    
    PopupMenu pop1, vbAlignRight, (payeelist.SelectedItem.Left + 800), (payeelist.SelectedItem.Top + payeelist.Top + 250)
    
End If

End Sub

Private Sub payeelist_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = vbRightButton And payeelist.ListItems.Count > 0 Then
    PopupMenu pop1
End If
End Sub

Private Sub refresh_Click()
Call addlist
savecon = True
End Sub

Private Sub txtcode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtnameineng.SetFocus
End If
If KeyAscii = 27 Then
    For Each Control In Me
    If TypeOf Control Is TextBox Then
        Control.Text = ""
    End If
Next
savecon = True
cmdclose.Enabled = True
cmdsave.Enabled = False
cmdUpdate.Enabled = False
txtcode.SetFocus

End If
End Sub

Private Sub txtnameinara_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
For Each Control In Me
    If TypeOf Control Is TextBox Then
        If Control.Text = "" Then
            If Control.Name <> "txtnameinara" Then
                Control.SetFocus
                Exit Sub
            End If
            
        End If
    End If
Next

    If savecon = True Then
    cmdUpdate.Enabled = False
    cmdsave.Enabled = True
    cmdsave.SetFocus
    Else
    cmdUpdate.Enabled = True
    cmdsave.Enabled = False
    cmdUpdate.SetFocus
    End If

End If

If KeyAscii = 27 Then
    For Each Control In Me
    If TypeOf Control Is TextBox Then
        Control.Text = ""
    End If
Next
savecon = True
cmdclose.Enabled = True
cmdsave.Enabled = False
cmdUpdate.Enabled = False
txtcode.SetFocus

End If

End Sub

Private Sub txtnameineng_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtnameinara.SetFocus
End If
If KeyAscii = 27 Then
    For Each Control In Me
    If TypeOf Control Is TextBox Then
        Control.Text = ""
    End If
Next
savecon = True
cmdclose.Enabled = True
cmdsave.Enabled = False
cmdUpdate.Enabled = False
txtcode.SetFocus

End If
End Sub

Public Sub addlist()
'this is for add the list from table
payeelist.ListItems.Clear
payeelist.Sorted = False
i = 1
If recpayee.BOF = False Then
    recpayee.MoveFirst
    While recpayee.EOF = False
        payeelist.ListItems.Add , , recpayee!payeecode
        payeelist.ListItems(i).ListSubItems.Add , , recpayee!payeenameeng
        payeelist.ListItems(i).ListSubItems.Add , , recpayee!payeenameara
        recpayee.MoveNext
        i = i + 1
    Wend
End If

End Sub

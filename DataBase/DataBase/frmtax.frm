VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmtax 
   Caption         =   "Tax Details"
   ClientHeight    =   4965
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8445
   Icon            =   "frmtax.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   8445
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame detailframe 
      Caption         =   "Data Entry For Tax Details"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   6615
      Begin VB.TextBox txtrate 
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2520
         MaxLength       =   5
         TabIndex        =   3
         Top             =   1440
         Width           =   735
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
         MaxLength       =   5
         TabIndex        =   0
         Top             =   360
         Width           =   735
      End
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Rate ÓÚÑ "
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Name in Arabic ÇáÇÓã ÇáÚÑÈí "
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   2025
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Name in English ÇáÇÓã ÈÇáÇäÌáíÒí"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   2280
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Code ßæÏ"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   690
      End
   End
   Begin VB.Frame commandframe 
      Height          =   2055
      Left            =   6960
      TabIndex        =   4
      Top             =   0
      Width           =   1335
      Begin VB.CommandButton cmdprint 
         Caption         =   "&Print"
         Height          =   325
         Left            =   180
         TabIndex        =   16
         Top             =   1280
         Width           =   1000
      End
      Begin VB.CommandButton cmdupdate 
         Caption         =   "&Update"
         Enabled         =   0   'False
         Height          =   325
         Left            =   180
         TabIndex        =   8
         Top             =   560
         Width           =   1000
      End
      Begin VB.CommandButton cmdclose 
         Caption         =   "&Close"
         Height          =   325
         Left            =   180
         TabIndex        =   7
         Top             =   1630
         Width           =   1000
      End
      Begin VB.CommandButton cmdsave 
         Caption         =   "&Save New"
         Enabled         =   0   'False
         Height          =   325
         Left            =   180
         TabIndex        =   6
         Top             =   200
         Width           =   1000
      End
      Begin VB.CommandButton cmdcancel 
         Caption         =   "&Cancel"
         Height          =   325
         Left            =   180
         TabIndex        =   5
         Top             =   920
         Width           =   1000
      End
   End
   Begin MSComctlLib.ListView taxlist 
      Height          =   2415
      Left            =   120
      TabIndex        =   13
      Top             =   2400
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   4260
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
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Tax Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Tax Name in English"
         Object.Width           =   4057
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Tax Name in Arabic"
         Object.Width           =   4057
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Rate"
         Object.Width           =   1411
      EndProperty
   End
   Begin VB.Label Label4 
      Caption         =   "View / Edit Tax Details"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Menu rightmenu 
      Caption         =   "rightmenu"
      Visible         =   0   'False
      Begin VB.Menu edit 
         Caption         =   "Edit"
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
Attribute VB_Name = "frmtax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public savecon As Boolean
Public oldcode As String
Dim rectax As New ADODB.Recordset
Dim CON1 As New ADODB.Connection
Dim myclass As New HabitatClass
Dim sqltable As Boolean
Dim xtable As String
Dim xdecimal As Integer
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

Private Sub cmdPrint_Click()
reporttax.Show 1
End Sub

Private Sub cmdsave_Click()
'code for save button
'to serch for the repeating code
If rectax.BOF = False Then
    rectax.MoveFirst
    While rectax.EOF = False
        If UCase(Trim(rectax!taxcode)) = UCase(Trim(txtcode.Text)) Then
            MsgBox "Please Check Your Tax Code", vbInformation, "Data Repeating"
            txtcode.SetFocus
            Exit Sub
        End If
        rectax.MoveNext
    Wend
End If
   
If Val(Trim(txtRate.Text)) <= 0 Or Val(Trim(txtRate.Text)) > 100 Then
    MsgBox "Please Conform Your Tax Rate", vbInformation, "Error Rate"
    txtRate.SetFocus
    Exit Sub
End If

rectax.AddNew
rectax!taxcode = UCase(Trim(txtcode.Text))
rectax!taxnameara = Trim(txtnameinara.Text)
rectax!taxnameeng = Trim(txtnameineng.Text)
rectax!taxrate = Trim(txtRate.Text)
rectax.Update
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
txtRate.Text = ""
Call addlist
txtcode.SetFocus

End Sub

Private Sub cmdUpdate_Click()

If rectax.BOF = False Then
    rectax.MoveFirst
    Do While rectax.EOF = False
        If UCase(Trim(rectax!taxcode)) = UCase(Trim(txtcode.Text)) And UCase(Trim(oldcode)) <> UCase(Trim(txtcode.Text)) Then
            MsgBox "Please Check Your Tax code", vbInformation, "Code Error"
            savecon = False
            txtcode.SetFocus
            Exit Sub
        End If
        rectax.MoveNext
    Loop
End If

rectax.MoveFirst
While rectax.EOF = False
    If Trim(oldcode) = Trim(rectax!taxcode) Then
        rectax!taxcode = UCase(Trim(txtcode.Text))
        rectax!taxnameara = Trim(txtnameinara.Text)
        rectax!taxnameeng = Trim(txtnameineng.Text)
        If Val(Trim(txtRate.Text)) = 0 Then
            MsgBox "Please check Your Tax Rate", vbInformation, "Invalid Percentage"
            txtRate.SetFocus
            Exit Sub
        Else
        rectax!taxrate = Trim(txtRate.Text)
        End If
        rectax.Update
        MsgBox "Your Data Updated Successfully", vbInformation, "Saved"
        okkalas = 1
        savecon = True 'this is identifier from listview
        cmdclose.Enabled = True
        cmdsave.Enabled = False
        cmdprint.Enabled = True
        cmdUpdate.Enabled = False
        txtcode.Text = ""
        txtnameinara.Text = ""
        txtnameineng.Text = ""
        txtRate.Text = ""
        txtcode.SetFocus
        Call addlist
        Exit Sub
    End If
    rectax.MoveNext
Wend
If okkalas <> 1 Then
    MsgBox "Please Refresh Your Data and Try Again", vbInformation, "Data Error"
    Exit Sub
End If

End Sub

Private Sub delete_Click()
If MsgBox("Are You Sure You Want to Delete This Item", vbQuestion + vbYesNo, "Conformation Deletion") = vbYes Then
    deleteitem = Trim(taxlist.SelectedItem.Text)
    rectax.MoveFirst
        Do While rectax.EOF = False
            If deleteitem = Trim(rectax!taxcode) Then
                   rectax.Delete
                Exit Do
            End If
            rectax.MoveNext
        Loop
    txtcode.Text = ""
    txtnameinara.Text = ""
    txtnameineng.Text = ""
    txtRate.Text = ""
    savecon = True
    Call addlist
    End If

End Sub

Private Sub edit_Click()
If taxlist.ListItems.Count > 0 Then
oldcode = taxlist.SelectedItem.Text

'MsgBox oldcode
For Each Control In Me
    If TypeOf Control Is TextBox Then
        Control.Text = ""
    End If
Next
txtcode.Text = taxlist.SelectedItem.Text
txtnameineng.Text = taxlist.SelectedItem.ListSubItems(1).Text
txtnameinara.Text = taxlist.SelectedItem.ListSubItems(2).Text
txtRate.Text = taxlist.SelectedItem.ListSubItems(3).Text
savecon = False
txtcode.SetFocus
End If

End Sub

Private Sub Form_Activate()
txtcode.SetFocus
End Sub

Private Sub Form_Load()
xtable = "Select * from tax"
sqltable = True
myclass.GetTables rectax, CON1, xtable, constring, sqltable
Call addlist

savecon = True
End Sub

Private Sub Form_Resize()
On Error Resume Next
taxlist.Height = Me.Height - 3000
taxlist.Width = Me.Width - 360
commandframe.Left = Me.Width - 1600
detailframe.Width = commandframe.Left - 300
taxlist.ColumnHeaders(2).Width = (taxlist.Width - 2270) / 2
taxlist.ColumnHeaders(3).Width = (taxlist.Width - 2270) / 2
taxlist.ColumnHeaders(4).Width = 800
txtnameinara.Width = Me.Width - 5775
txtnameineng.Width = Me.Width - 5775

End Sub

Private Sub Form_Unload(Cancel As Integer)
If detailframe.Enabled = True Then
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

Private Sub refresh_Click()
Call addlist
savecon = True

End Sub

Private Sub taxlist_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
taxlist.SortKey = ColumnHeader.Index - 1
taxlist.Sorted = True
End Sub

Private Sub taxlist_DblClick()
If taxlist.ListItems.Count > 0 Then
oldcode = taxlist.SelectedItem.Text

'MsgBox oldcode
For Each Control In Me
    If TypeOf Control Is TextBox Then
        Control.Text = ""
    End If
Next
txtcode.Text = taxlist.SelectedItem.Text
txtnameineng.Text = taxlist.SelectedItem.ListSubItems(1).Text
txtnameinara.Text = taxlist.SelectedItem.ListSubItems(2).Text
txtRate.Text = taxlist.SelectedItem.ListSubItems(3).Text
savecon = False
txtcode.SetFocus
End If

End Sub

Private Sub taxlist_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    taxlist_DblClick
End If

End Sub

Private Sub taxlist_KeyUp(KeyCode As Integer, Shift As Integer)
'MsgBox taxlist.SelectedItem.Left & "  " & taxlist.Left
 If KeyCode = 93 Then
    PopupMenu rightmenu, vbAlignRight, (taxlist.SelectedItem.Left + 800), (taxlist.SelectedItem.Top + taxlist.Top + 250)
End If

End Sub

Private Sub taxlist_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = vbRightButton Then
    PopupMenu rightmenu
End If
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
cmdprint.Enabled = True
cmdUpdate.Enabled = False
txtcode.SetFocus

End If

End Sub

Private Sub txtnameinara_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtRate.SetFocus
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
cmdprint.Enabled = True
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
cmdprint.Enabled = True
cmdUpdate.Enabled = False
txtcode.SetFocus

End If

End Sub


Private Sub txtrate_GotFocus()
xdecimal = 0
End Sub

Private Sub txtrate_KeyPress(KeyAscii As Integer)
'start number testing
If KeyAscii = 46 Then
   xdecimal = xdecimal + 1
 End If
If xdecimal = 2 Then
  Beep
  xdecimal = 1
 txtRate.SetFocus
 SendKeys "{Left}+{End}"
 SendKeys "{Delete}"
End If
If KeyAscii = 13 Or KeyAscii = 46 Or KeyAscii >= 48 And KeyAscii <= 57 Then
 c = KeyAscii
 
 
 Else
   If KeyAscii = 46 And xdecimal > 1 Then
       xdecimal = 0
   End If
 If txtRate.Text <> " " Then
        xdecimal = 0

  SendKeys "{End}+{Home}"
  SendKeys "{Delete}"
  txtRate.SetFocus

 End If
End If
'end number testing
If KeyAscii = 13 Then
For Each Control In Me
    If TypeOf Control Is TextBox Then
        If Control.Text = "" Then
            If Control.Name <> "txtnameinara" Then
                Control.SetFocus
                Exit Sub
            End If
        End If
        
        If Control.Name = "txtrate" Then
                If Val(Trim(Control.Text)) <= 0 Or Val(Trim(Control.Text)) > 100 Then
                    MsgBox "Please check Your Tax rate", vbInformation, "Inalid Percentage"
                    txtRate.SetFocus
                    Exit Sub
                End If
            End If
    End If
Next
    If savecon = True Then
    cmdUpdate.Enabled = False
    cmdsave.Enabled = True
    cmdprint.Enabled = True
    cmdsave.SetFocus
    Else
    cmdUpdate.Enabled = True
    cmdsave.Enabled = False
    cmdprint.Enabled = True
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

Public Sub addlist()
'this is for add the list from table
taxlist.ListItems.Clear
taxlist.Sorted = False
i = 1
If rectax.BOF = False Then
    rectax.MoveFirst
    While rectax.EOF = False
        taxlist.ListItems.Add , , rectax!taxcode
        taxlist.ListItems(i).ListSubItems.Add , , rectax!taxnameeng
        taxlist.ListItems(i).ListSubItems.Add , , IIf(IsNull(rectax!taxnameara) = True, " ", Trim(rectax!taxnameara))
        taxlist.ListItems(i).ListSubItems.Add , , rectax!taxrate
        rectax.MoveNext
        i = i + 1
    Wend
End If

End Sub

Private Sub txtrate_LostFocus()
xdecimal = 0
End Sub

VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPayCat 
   Caption         =   "Payment Category"
   ClientHeight    =   4230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   ScaleHeight     =   4230
   ScaleWidth      =   6675
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   7
      ToolTipText     =   "Click to Save"
      Top             =   0
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   720
      Width           =   3255
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   1200
      Width           =   3255
   End
   Begin VB.CommandButton CMDEDIT 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   5400
      TabIndex        =   3
      Top             =   405
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   5400
      TabIndex        =   2
      Top             =   810
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   5400
      TabIndex        =   1
      Top             =   1215
      Width           =   1095
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   1680
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   4471
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label1 
      Caption         =   "Payment Code"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Name"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Arab"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "frmPayCat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RstStat As ADODB.Recordset

Dim MItem As ListItem



Private Sub CMDEDIT_Click()
Set RstStat = New ADODB.Recordset

Dim CON1 As ADODB.Connection
Set CON1 = New ADODB.Connection


conStr1 = "Provider=MSDASQL;DSN=Finance;UID=sa; PWD=;"
CON1.Open conStr1


    RstStat.Open "Select * from newPAymentfor", CON1, adOpenDynamic, adLockOptimistic

If Me.ListView1.ListItems.Count = 0 Then
    'Me.edit.Enabled = False
    Exit Sub
  End If
varindex = Me.ListView1.SelectedItem.Index
varitem = Me.ListView1.SelectedItem.Text
varsubitem = Me.ListView1.SelectedItem.SubItems(1)

xvar = Trim(varitem)


If RstStat.EOF = False Then
RstStat.MoveFirst
End If

  If Me.cmdedit.caption = Trim("&Edit") Then
      
      While RstStat.EOF = False
        If Trim(RstStat!Code) = Trim(xvar) Then
        Me.Text1.Text = IIf(IsNull(Trim(RstStat!Code)), "", Trim(RstStat!Code))
        Me.Text2.Text = IIf(IsNull(Trim(RstStat!Name)), "", Trim(RstStat!Name))
        Me.Text3.Text = IIf(IsNull(Trim(RstStat!arab)), "", Trim(RstStat!arab))
        Me.cmdsave.Enabled = False
        Me.Command4.Enabled = False
        Me.cmdedit.caption = "Update"
        Me.Command5.caption = "&Cancel"
        frmMenu.sclear.Enabled = True
        frmMenu.sEdit.Enabled = False

        End If
     RstStat.MoveNext
        Wend
ElseIf Me.cmdedit.caption = Trim("Update") Then
If RstStat.EOF = False Then
RstStat.MoveFirst
End If

      While RstStat.EOF = False
        If Trim(RstStat!Code) = Trim(xvar) Then

        RstStat!Code = Text1.Text
        RstStat!Name = Text2.Text
        RstStat!arab = Text3.Text
        Me.cmdsave.Enabled = True
        Me.Command4.Enabled = True
        Me.cmdedit.caption = "&Edit"
        frmMenu.sEdit.Enabled = True
        Me.Command5.caption = "E&xit"
      End If
   RstStat.MoveNext
   Wend
   
End If 'edit

End Sub

Private Sub cmdsave_Click()
If Text1 = "" Or Text2 = "" Then
MsgBox "Please select Status Code and Name ", vbInformation, "Payroll"
Exit Sub
End If

v = MsgBox("Are You sure you want to Save the records", vbYesNoCancel, "Payroll")
If v = 7 Then
Exit Sub
End If
RstStat.AddNew
            RstStat!Code = Text1.Text
            RstStat!Name = Text2.Text
            RstStat!arab = Text3.Text

RstStat.Update

             MsgBox "records Of: " & Trim(RstStat!Code) & " " & _
         Trim(RstStat!Name) & "Has Been Added Successfully  ÊãÊ ÇáÇÖÇÝÉ ÈäÌÇÍ ", vbInformation, "Habitat"

End Sub

Private Sub Command4_Click()
Set RstStat = New ADODB.Recordset
Dim CON1 As ADODB.Connection
Set CON1 = New ADODB.Connection
conStr1 = "Provider=MSDASQL;DSN=Finance;UID=sa; PWD=;"
CON1.Open conStr1


    RstStat.Open "Select * from newPAymentfor", CON1, adOpenDynamic, adLockOptimistic

If Me.ListView1.ListItems.Count = 0 Then
    frmMenu.sDel.Enabled = False
    Exit Sub
  End If
varindex = Me.ListView1.SelectedItem.Index
varitem = Me.ListView1.SelectedItem.Text
'varsubit0 = me.
varsubitem = Me.ListView1.SelectedItem.SubItems(1)
 xmsg = MsgBox("Are you Sure Deleting Status?    " & varsubitem & "", vbQuestion + vbYesNo, " Delete...")

  If xmsg = vbYes Then
     'me.ListView1.ListItems.Remove varindex
     xvar = Trim(varitem)
   If RstStat.BOF Then
   Exit Sub
   End If
If RstStat.EOF = False Then
RstStat.MoveFirst
End If

      While RstStat.EOF = False
        If Trim(RstStat!Code) = Trim(xvar) Then
            RstStat.Delete
            RstStat.Update
            MsgBox "Records of '" & varsubitem & "' Deleted Successfully ", vbInformation, "Confirmation"
        End If
     RstStat.MoveNext
        Wend
     
End If


End Sub

Private Sub Command5_Click()
If cmdedit.caption = "&Edit" Then
Unload Me
Else '...cancel
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""

cmdedit.caption = "&Edit"
Command4.Enabled = True
cmdsave.Enabled = True
Command5.caption = "E&xit"
frmMenu.sEdit.Enabled = True

End If
End Sub

Private Sub Form_Activate()
Set RstStat = New ADODB.Recordset
'Combo1.Visible = False

Dim CON1 As ADODB.Connection
Set CON1 = New ADODB.Connection

Set xcol = Me.ListView1.ColumnHeaders.Add(, , "CODE", 1000)
Set xcol = Me.ListView1.ColumnHeaders.Add(, , "Name", 2650)
Set xcol = Me.ListView1.ColumnHeaders.Add(, , "Arab", 2500)

'conString = "Provider=MSDASQL;DSN=finance;UID=; PWD=;"
'con1.Open conString

conStr1 = "Provider=MSDASQL;DSN=Finance;UID=sa; PWD=;"
CON1.Open conStr1


    RstStat.Open "Select * from newPAymentfor", CON1, adOpenDynamic, adLockOptimistic
If RstStat.EOF = False Then
RstStat.MoveFirst
End If

  While RstStat.EOF = False
     Set MItem = Me.ListView1.ListItems.Add(, , Format(RstStat!Code))
     MItem.SubItems(1) = Format(RstStat!Name)
     MItem.SubItems(2) = Format(RstStat!arab)

     RstStat.MoveNext
     Wend
frmMenu.sclear.Enabled = False
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = 2 Then
  PopupMenu frmMyMenu2.Status
  End If
End Sub



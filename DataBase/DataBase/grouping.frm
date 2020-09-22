VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Grouping 
   Caption         =   "Accounts Grouping"
   ClientHeight    =   6840
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   8550
   Icon            =   "grouping.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   8550
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "grouping.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "grouping.frx":0894
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "grouping.frx":0CE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "grouping.frx":1138
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "grouping.frx":158A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   5106
      _Version        =   393217
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3495
      Left            =   0
      TabIndex        =   1
      Top             =   3360
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   6165
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Add Group Cat"
      TabPicture(0)   =   "grouping.frx":19DC
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label7"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label13"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Combo5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Combo6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Combo11"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Command2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Add Group Name"
      TabPicture(1)   =   "grouping.frx":19F8
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label9"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label10"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label11"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label12"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Command3"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Combo7"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Combo8"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Combo9"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Combo10"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Add  Member"
      TabPicture(2)   =   "grouping.frx":1A14
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label2"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label3"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label4"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label5"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label14"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Text1"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Combo1"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Combo2"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Combo3"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Combo4"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "Combo12"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "Command1"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).ControlCount=   13
      Begin VB.CommandButton Command1 
         Caption         =   "Ok"
         Height          =   350
         Left            =   4800
         TabIndex        =   9
         Top             =   3000
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Ok"
         Enabled         =   0   'False
         Height          =   350
         Left            =   -70560
         TabIndex        =   7
         Top             =   3000
         Width           =   1095
      End
      Begin VB.ComboBox Combo12 
         Height          =   315
         Left            =   1680
         TabIndex        =   31
         Top             =   600
         Width           =   2655
      End
      Begin VB.ComboBox Combo11 
         Height          =   315
         Left            =   -73080
         Style           =   1  'Simple Combo
         TabIndex        =   29
         Top             =   1440
         Width           =   2535
      End
      Begin VB.ComboBox Combo10 
         Height          =   315
         Left            =   -73320
         Style           =   1  'Simple Combo
         TabIndex        =   26
         Top             =   1680
         Width           =   2895
      End
      Begin VB.ComboBox Combo9 
         Height          =   315
         Left            =   -73320
         Style           =   1  'Simple Combo
         TabIndex        =   24
         Top             =   1320
         Width           =   2895
      End
      Begin VB.ComboBox Combo8 
         Height          =   315
         Left            =   -73320
         TabIndex        =   22
         Top             =   960
         Width           =   2175
      End
      Begin VB.ComboBox Combo7 
         Height          =   315
         Left            =   -73320
         Locked          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   20
         Top             =   600
         Width           =   1095
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   1680
         TabIndex        =   13
         Top             =   2040
         Width           =   3855
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   1680
         TabIndex        =   12
         Top             =   1680
         Width           =   3855
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1680
         TabIndex        =   11
         Top             =   1320
         Width           =   2055
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1680
         TabIndex        =   10
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         Height          =   525
         Left            =   1680
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   8
         Text            =   "grouping.frx":1A30
         Top             =   2400
         Width           =   3855
      End
      Begin VB.ComboBox Combo6 
         Height          =   315
         Left            =   -73080
         Style           =   1  'Simple Combo
         TabIndex        =   5
         Top             =   1080
         Width           =   2535
      End
      Begin VB.ComboBox Combo5 
         Height          =   315
         Left            =   -73080
         Locked          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   3
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Ok"
         Enabled         =   0   'False
         Height          =   350
         Left            =   -70440
         TabIndex        =   27
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "Group Category"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label13 
         Caption         =   "Group Cat Name Arab"
         Height          =   255
         Left            =   -74760
         TabIndex        =   28
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label12 
         Caption         =   "Group NameArab"
         Height          =   375
         Left            =   -74760
         TabIndex        =   25
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "Group NameEng"
         Height          =   255
         Left            =   -74760
         TabIndex        =   23
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Group Category"
         Height          =   375
         Left            =   -74760
         TabIndex        =   21
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "ID No"
         Height          =   375
         Left            =   -74760
         TabIndex        =   19
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Classification"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "AccountNamArab"
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "AccountNameEng"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Account Code"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Group Names"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Group Cat Name Eng"
         Height          =   375
         Left            =   -74760
         TabIndex        =   4
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "ID No"
         Height          =   255
         Left            =   -74760
         TabIndex        =   2
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.Label Label8 
      Caption         =   "Click right mouse button to select options"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   80
      Width           =   4455
   End
   Begin VB.Menu xmenu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu xAddGRpName 
         Caption         =   "Add item.. "
      End
      Begin VB.Menu bar1 
         Caption         =   "-"
      End
      Begin VB.Menu xDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu xrefresh 
         Caption         =   "Refresh"
      End
      Begin VB.Menu xFind 
         Caption         =   "Find..."
      End
   End
End
Attribute VB_Name = "Grouping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xClass As New HabitatClass
Dim CON1 As New ADODB.Connection
Dim AcctCode As New ADODB.Recordset
Dim acctnames As ADODB.Recordset
Dim xtable As String
Dim sqltable As Boolean
Dim acctNo As String
Dim nodex As Node

Private Sub Combo1_GotFocus()
Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
  Tmp = SendMessage(Combo1.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub Combo12_Click()
Dim rsGrpName As New ADODB.Recordset
groupNameCode = Left(Trim(Me.Combo12), 2)
rsGrpName.Open "Select * from GroupName where GroupCatCode=" & "'" & groupNameCode & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
Me.Combo1.clear
Do Until rsGrpName.EOF = True
    Me.Combo1.AddItem rsGrpName!Idno & "-" & rsGrpName!GroupNameEng
    rsGrpName.MoveNext
Loop
rsGrpName.close

End Sub

Private Sub Combo2_Click()
xAccount = Me.Combo2
Dim RstBA As ADODB.Recordset
Set RstBA = New ADODB.Recordset
xtable = "FinanceMaster"
xKey = "select * from " & xtable & " where " & _
       " AccountCode = " & "'" & xAccount & "'"
         
RstBA.Open xKey, constring, adOpenDynamic, adLockOptimistic, adCmdText
If RstBA.EOF = False Then
  Me.Combo3 = RstBA!accountnameeng
  Me.Combo4 = RTrim(RstBA!accountnamearab)
End If
RstBA.close

Dim catName As String
Dim Prevcap As String
Dim acctNo As String
acctNo = Trim(Me.Combo2)
Prevcap = Trim(Me.caption)
Call DisplayCats(Prevcap, acctNo, catName)
Me.Text1.Text = catName

End Sub

Private Sub Combo2_GotFocus()
Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
  Tmp = SendMessage(Combo2.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub Combo2_LostFocus()
Call Combo2_Click
End Sub

Private Sub Combo3_Click()
xAccount = Trim(Me.Combo3)
Dim RstBA As ADODB.Recordset
Set RstBA = New ADODB.Recordset
xtable = "FinanceMaster"
xKey = "select * from " & xtable & " where " & _
       " AccountNameEng = " & "'" & xAccount & "'"
         
RstBA.Open xKey, constring, adOpenDynamic, adLockOptimistic, adCmdText
If RstBA.EOF = False Then
  Me.Combo2 = RstBA!AccountCode
  Me.Combo4 = RTrim(RstBA!accountnamearab)
End If
RstBA.close

Dim catName As String
Dim Prevcap As String
Dim acctNo As String
acctNo = Trim(Me.Combo3)
Prevcap = Trim(Me.caption)
Call DisplayCatsName(Prevcap, acctNo, catName)
Me.Text1.Text = catName
End Sub

Private Sub Combo3_LostFocus()
Call Combo3_Click
End Sub

Private Sub Combo4_Click()

acctnames.MoveFirst
While acctnames.EOF = False
    If UCase(Trim(Me.Combo4)) = UCase(Trim(acctnames!accountnamearab)) Then
       Me.Combo2 = acctnames!AccountCode
       Me.Combo3 = acctnames!accountnameeng
    End If
    acctnames.MoveNext
Wend
Dim catName As String
Dim Prevcap As String
Dim acctNo As String
acctNo = Trim(Me.Combo3)
Prevcap = Trim(Me.caption)
Call DisplayCatsName(Prevcap, acctNo, catName)
Me.Text1.Text = catName
End Sub

Private Sub Combo6_Change()
If Trim(Me.Combo6) = "" Then
    Me.Command2.Enabled = False
   Else
    Me.Command2.Enabled = True
End If
End Sub

Private Sub Combo9_Change()
If Trim(Me.Combo9) = "" Then
    Me.Command3.Enabled = False
   Else
    Me.Command3.Enabled = True
End If
End Sub

Private Sub Command1_Click()
If Me.Combo1 = "" Then
    mess = MsgBox("Please select Group Name")
    Me.Combo1.SetFocus
    Exit Sub
End If

If Me.Combo12 = "" Then
    mess = MsgBox("Please select Group Category")
    Me.Combo12.SetFocus
    Exit Sub
End If

If Me.Combo2 = "" Then
    mess = MsgBox("Please select Group Code")
    Me.Combo2.SetFocus
    Exit Sub
End If

Dim rsgrp As New ADODB.Recordset
If Me.Command1.Enabled = False Then
    Exit Sub
End If
If Trim(Me.Combo1) = "" Or Trim(Me.Combo2) = "" Then
  mess = MsgBox("Please don't leave with blank", vbExclamation + vbOKOnly, "Message")
  Exit Sub
 Else
  mess = MsgBox("Save entries now?", vbQuestion + vbYesNo, "Please confirm")
End If
If mess = vbYes Then
    rsgrp.Open "select count(idno) as cTOtal from  Groupmember", constring, adOpenKeyset, adLockPessimistic, adCmdText
    cTOtal = rsgrp!cTOtal + 1
    rsgrp.close
    rsgrp.Open "Groupmember", constring, adOpenKeyset, adLockPessimistic, adCmdTable
  
    With rsgrp
        .addnew
        
        !Idno = cTOtal
        !GroupCatCode = Left(Me.Combo12, 2)
        !groupNameCode = Left(Me.Combo1, 4)
        !memberAcctCode = Trim(Me.Combo2)
        !memberNameEng = Trim(Me.Combo3)
        !memberNameArab = Trim(Me.Combo4)
        !Classification = Trim(Me.Text1.Text)
        .Update
        .close
        
    End With
End If
For Each Control In Me
    If TypeOf Control Is ComboBox Then
        Control.Text = ""
    End If
Next
Me.Text1.Text = ""
DisplayGroup
Exit Sub
End Sub

Private Sub Command1_LostFocus()
Call Command1_Click
End Sub

Private Sub Command2_Click()
Dim rsgrp As New ADODB.Recordset
If Me.Command2.Enabled = False Then
    Exit Sub
End If
mess = MsgBox("Save entries now?", vbQuestion + vbYesNo, "Please confirm")
If mess = vbYes Then
    rsgrp.Open "Groupcat", constring, adOpenKeyset, adLockPessimistic, adCmdTable
    With rsgrp
        .addnew
        !Code = Trim(Me.Combo5)
        !NameEng = Trim(Me.Combo6)
        !NameArab = Trim(Me.Combo11)
        .Update
        .close
    End With
End If
Me.Combo5 = ""
Me.Combo6 = ""
DisplayGroup
Dim rsGrpList As New ADODB.Recordset
rsGrpList.Open "GroupCat", constring, adOpenKeyset, adLockPessimistic, adCmdTable
i = 0
While rsGrpList.EOF = False
     i = i + 1
    Me.Combo1.AddItem rsGrpList!Code & "-" & rsGrpList!NameEng
    rsGrpList.MoveNext
Wend
cTOtal = i + 1
If cTOtal > 9 And cTOtal < 100 Then
        xCode = "" & cTOtal
      ElseIf cTOtal < 10 Then
        xCode = "0" & cTOtal
End If
Me.Combo5 = xCode
rsGrpList.close
End Sub

Private Sub Command3_Click()
If Me.Combo8 = "" Then
   mess = MsgBox("Please select Group Category")
   Me.Combo8.SetFocus
    Exit Sub
End If
Dim rsGrpName As New ADODB.Recordset
If Me.Command3.Enabled = False Then
    Exit Sub
End If
mess = MsgBox("Save entries now?", vbQuestion + vbYesNo, "Please confirm")
If mess = vbYes Then
    rsGrpName.Open "GroupName", constring, adOpenKeyset, adLockPessimistic, adCmdTable
    With rsGrpName
        .addnew
        !Idno = Trim(Me.Combo7)
        !GroupCatCode = Left(Trim(Me.Combo8), 2)
        !GroupNameEng = Trim(Me.Combo9)
        !GroupNameArab = Trim(Me.Combo10)
        .Update
        .close
    End With
End If
Me.Combo7 = ""
Me.Combo8 = ""
Me.Combo9 = ""
Me.Combo10 = ""
DisplayGroup

End Sub

Private Sub Form_Load()
 Dim xClass As New HabitatClass
 Set CON1 = New ADODB.Connection
 Set acctnames = New ADODB.Recordset
 xtable = "Select * from FinanceMaster order by AccountCode"
 sqltable = True
 xClass.GetTables acctnames, CON1, xtable, constring, sqltable
 While acctnames.EOF = False
  If acctnames!Active <> 0 Then
   Me.Combo2.AddItem acctnames!AccountCode
   Me.Combo3.AddItem acctnames!accountnameeng
   Me.Combo4.AddItem acctnames!accountnamearab
  End If
  acctnames.MoveNext
  Wend
 
Dim rsGrpList As New ADODB.Recordset

'rsGrpList.Close



DisplayGroup


Me.TreeView1.Height = 6495
End Sub
Sub DisplayGroup()
Dim rsgrpCat As New ADODB.Recordset
Dim rsGrpMember As New ADODB.Recordset
Dim rsGrpName As New ADODB.Recordset
Me.TreeView1.Nodes.clear
rsgrpCat.Open "select * from GroupCat order by code", constring, adOpenDynamic, adLockOptimistic, adCmdText
i = 0
Me.Combo8.clear
Me.Combo12.clear
While rsgrpCat.EOF = False
     i = i + 1
    Me.Combo8.AddItem rsgrpCat!Code & "-" & rsgrpCat!NameEng
    Me.Combo12.AddItem rsgrpCat!Code & "-" & rsgrpCat!NameEng
    rsgrpCat.MoveNext
Wend
cTOtal = i + 1
If cTOtal > 9 And cTOtal < 100 Then
        xCode = "" & cTOtal
     ElseIf cTOtal < 10 Then
        xCode = "0" & cTOtal
End If
Me.Combo5 = xCode
'=========================================
'group Name
'Dim rsGrpName As New ADODB.Recordset
rsGrpName.Open "select * from GroupName order by idno", constring, adOpenKeyset, adLockPessimistic, adCmdText
i = 0
Me.Combo1.clear
While rsGrpName.EOF = False
     i = i + 1
    Me.Combo1.AddItem rsGrpName!Idno & "-" & rsGrpName!GroupNameEng & " \" & rsGrpName!GroupNameArab
    rsGrpName.MoveNext
Wend
cTOtal = i + 1
If cTOtal > 9 And cTOtal < 100 Then
        xCode = "00" & cTOtal
     ElseIf cTOtal > 99 Then
        xCode = "0" & cTOtal
      ElseIf cTOtal < 10 Then
        xCode = "000" & cTOtal
End If
Me.Combo7 = xCode
rsGrpName.close



'list it in treeview
On Error Resume Next
rsgrpCat.MoveFirst
On Error GoTo 0
Set nodex = Me.TreeView1.Nodes.Add(, , "a", "Groups", 2)
Do Until rsgrpCat.EOF = True
    xCode = Trim(rsgrpCat!Code)
    Set nodex = Me.TreeView1.Nodes.Add("a", tvwChild, "b" & rsgrpCat!Code, rsgrpCat!Code & "-" & rsgrpCat!NameEng, 1)
    xKey = "b" & rsgrpCat!Code
    
    
    rsGrpName.Open "select * from GroupNAMe where GroupCatcode=" & "'" & xCode & "'" & "order by idno", constring, adOpenDynamic, adLockOptimistic, adCmdText
    i = 0
    Do Until rsGrpName.EOF = True
          i = i + 1
          xcode1 = Trim(rsGrpName!Idno)
          Set nodex = Me.TreeView1.Nodes.Add(xKey, tvwChild, "c" & LTrim(i) & rsGrpName!Idno & rsGrpName!GroupCatCode, rsGrpName!Idno & "-" & rsGrpName!GroupNameEng & "\" & rsGrpName!GroupNameArab, 4)
          xKey1 = "c" & LTrim(i) & rsGrpName!Idno & rsGrpName!GroupCatCode
          
          
          rsGrpMember.Open "select * from Groupmember where GroupCatcode=" & "'" & xCode & "'" & "and GroupNameCode =" & "'" & xcode1 & "'" & "order by idno", constring, adOpenDynamic, adLockOptimistic, adCmdText
          i = 0
          Do Until rsGrpMember.EOF = True
             i = i + 1
             Set nodex = Me.TreeView1.Nodes.Add(xKey1, tvwChild, "d" & LTrim(i) & rsGrpMember!Idno, rsGrpMember!memberAcctCode & "-" & rsGrpMember!memberNameEng & "\" & rsGrpMember!memberNameArab, 5)
             rsGrpMember.MoveNext
          Loop
          rsGrpMember.close
      
      
      rsGrpName.MoveNext
     Loop
     rsGrpName.close
      
  rsgrpCat.MoveNext
Loop
End Sub

Private Sub Form_Resize()
On Error Resume Next
Me.Height = 7245
Me.TreeView1.Width = Me.Width - 120
Me.SSTab1.Width = Me.Width - 120
Me.Combo1.Width = Me.SSTab1.Width - 2300
Me.Combo2.Width = Me.SSTab1.Width - 2800
Me.Combo3.Width = Me.SSTab1.Width - 1800
Me.Combo4.Width = Me.SSTab1.Width - 1800
Me.Combo5.Width = Me.SSTab1.Width - 2700
Me.Combo6.Width = Me.SSTab1.Width - 2300
Me.Text1.Width = Me.SSTab1.Width - 1800
Me.Command1.Left = Me.SSTab1.Width - 1500
Me.Command2.Left = Me.SSTab1.Width - 1500
Me.Command3.Left = Me.SSTab1.Width - 1500
End Sub

Private Sub TreeView1_Click()
Me.TreeView1.Height = 6495
Dim rsClass As New ADODB.Recordset
If Len(acctNo) = 12 Then
 rsClass.Open "select * from groupmember where memberacctCode=" & "'" & acctNo & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
 If rsClass.EOF = False Then
   Me.TreeView1.ToolTipText = rsClass!Classification
   Else
   Me.TreeView1.ToolTipText = ""
  End If
 rsClass.close
 Else
 Me.TreeView1.ToolTipText = ""
End If
End Sub

Private Sub TreeView1_Collapse(ByVal Node As MSComctlLib.Node)
If Node.Index = 1 Then
 Node.Image = 2
End If
End Sub

Private Sub TreeView1_Expand(ByVal Node As MSComctlLib.Node)
If Node.Index = 1 Then
 Node.Image = 3
End If
End Sub

Private Sub TreeView1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = 2 Then
    PopupMenu xmenu
End If
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
On Error Resume Next
TakeAN = InStr(1, Node.Text, "-", vbTextCompare)
acctNo = Left(Node.Text, TakeAN - 1)

End Sub

Private Sub xAddGrpMember_Click()
Me.TreeView1.Height = 2895
End Sub

Private Sub xAddGRpName_Click()
Me.TreeView1.Height = 2895
If Trim(Me.SSTab1.caption) <> "Add New Group" Then
 Me.SSTab1.SetFocus
 SendKeys "{Left}"
End If
End Sub

Private Sub xFind_Click()
FindGrpMember.Show
End Sub

Private Sub xREfresh_Click()
DisplayGroup
End Sub

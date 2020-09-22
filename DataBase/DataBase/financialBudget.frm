VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form financialBudget 
   Caption         =   "Financial Budget"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   8565
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8400
      _ExtentX        =   14817
      _ExtentY        =   6800
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Data Entry"
      TabPicture(0)   =   "financialBudget.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(5)=   "Combo1"
      Tab(0).Control(6)=   "Combo2"
      Tab(0).Control(7)=   "Combo3"
      Tab(0).Control(8)=   "Combo5"
      Tab(0).Control(9)=   "Command1"
      Tab(0).Control(10)=   "Text1"
      Tab(0).Control(11)=   "Timer1"
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "List"
      TabPicture(1)   =   "financialBudget.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "ImageList2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "ListView1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   -74280
         Top             =   3000
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   -73200
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   1680
         Width           =   6375
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3435
         Left            =   60
         TabIndex        =   11
         Top             =   360
         Width           =   8280
         _ExtentX        =   14605
         _ExtentY        =   6059
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList2"
         SmallIcons      =   "ImageList2"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Account Code"
            Object.Width           =   2734
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Account Name Eng"
            Object.Width           =   4586
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Account Name Arab"
            Object.Width           =   4586
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Budget Amount"
            Object.Width           =   2558
         EndProperty
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   -69240
         TabIndex        =   10
         Top             =   3360
         Width           =   1095
      End
      Begin VB.ComboBox Combo5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73200
         Style           =   1  'Simple Combo
         TabIndex        =   9
         Top             =   2160
         Width           =   2175
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   -73200
         TabIndex        =   6
         Top             =   1320
         Width           =   3255
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   -73200
         TabIndex        =   4
         Top             =   960
         Width           =   3255
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   -73200
         TabIndex        =   2
         Top             =   600
         Width           =   2055
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   0
         Top             =   0
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
               Picture         =   "financialBudget.frx":0038
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "financialBudget.frx":048A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "financialBudget.frx":08DC
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "financialBudget.frx":0D2E
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "financialBudget.frx":1048
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label Label5 
         Caption         =   "Budget Amount"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74760
         TabIndex        =   8
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Classifications"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   7
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "AccountName Arab"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   5
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "AccountName Eng"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   3
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Account Code"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   1
         Top             =   600
         Width           =   1335
      End
   End
End
Attribute VB_Name = "financialBudget"
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
Dim MItem As ListItem
Dim acctNo As String

Private Sub Combo1_Click()
xAccount = Me.Combo1
Dim RstBA As ADODB.Recordset
Set RstBA = New ADODB.Recordset
xtable = "FinanceMaster"
xKey = "select * from " & xtable & " where " & _
       " AccountCode = " & "'" & xAccount & "'"
         
RstBA.Open xKey, constring, adOpenDynamic, adLockOptimistic, adCmdText
If RstBA.EOF = False Then
  
  Me.Combo2 = RstBA!accountnameeng
  Me.Combo3 = RTrim(RstBA!accountnamearab)
  On Error Resume Next
  Me.Combo5 = Format(RstBA!Budget, "###,###,###.#0")
  On Error GoTo 0
 Else
  msg = MsgBox("Account not found", vbExclamation + vbOKOnly, "Message")
  Me.Combo1.SetFocus
  Exit Sub
End If
RstBA.Close

Dim catName As String
Dim Prevcap As String
Dim acctNo As String
acctNo = Trim(Me.Combo1)
Prevcap = Trim(Me.caption)
Call DisplayCats(Prevcap, acctNo, catName)
Me.Text1.Text = catName
End Sub

Private Sub Combo1_GotFocus()
Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
  Tmp = SendMessage(Combo1.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.Combo5.SetFocus
End If
End Sub

Private Sub Combo1_LostFocus()
Call Combo1_Click
End Sub

Private Sub Combo2_Click()
xAccount = Trim(Me.Combo2)
Dim RstBA As ADODB.Recordset
Set RstBA = New ADODB.Recordset
xtable = "FinanceMaster"
xKey = "select * from " & xtable & " where " & _
       " AccountNameEng = " & "'" & xAccount & "'"
         
RstBA.Open xKey, constring, adOpenDynamic, adLockOptimistic, adCmdText
If RstBA.EOF = False Then
  Me.Combo1 = RstBA!AccountCode
  Me.Combo3 = RTrim(RstBA!accountnamearab)
  Me.Combo5 = Format(RstBA!Budget, "###,###,###.#0")
End If
RstBA.Close
Dim catName As String
Dim Prevcap As String
acctNo = Trim(Me.Combo2)
Prevcap = Trim(Me.caption)
Call DisplayCatsName(Prevcap, acctNo, catName)
Me.Text1.Text = catName
End Sub

Private Sub Combo3_Click()

Dim catName As String
Dim Prevcap As String
Dim acctNo As String
acctNo = Trim(Me.Combo1)
Prevcap = Trim(Me.caption)
Call DisplayCatsName(Prevcap, acctNo, catName)
Me.Text1.Text = catName
acctnames.MoveFirst
While acctnames.EOF = False
    If UCase(Trim(Me.Combo3)) = UCase(Trim(acctnames!accountnamearab)) Then
       Me.Combo1 = acctnames!AccountCode
       Me.Combo2 = acctnames!accountnameeng
       Me.Combo5 = Format(acctnames!Budget, "###,###,###.#0")
    End If
    acctnames.MoveNext
Wend

acctNo = Trim(Me.Combo2)
Prevcap = Trim(Me.caption)
Call DisplayCatsName(Prevcap, acctNo, catName)
Me.Text1.Text = catName


End Sub

Private Sub Combo5_Change()
If Me.Combo1 <> "" Then
    Me.Command1.Enabled = True
End If
End Sub

Private Sub Combo5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Me.Command1.SetFocus
End If
End Sub

Private Sub Command1_Click()
Dim Budget As Currency
Budget = Me.Combo5
mess = MsgBox("Save it?", vbOKCancel + vbQuestion, "Plesae confirm")
If mess = vbOK Then
    Dim rstFM As New ADODB.Recordset
    rstFM.Open "Update FInancemaster set budget =" & Budget & "where accountCode =" & "'" & Trim(Me.Combo1) & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
    Me.Combo5 = ""
    Me.Combo1.SetFocus
    
End If
End Sub

Private Sub Form_Activate()
Me.Combo1.SetFocus
SendKeys "{Down}"
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
   Set MItem = Me.ListView1.ListItems.Add(, , acctnames!AccountCode, , IIf(IsNull(acctnames!Budget) = True Or (acctnames!Budget) = 0, 4, 1))
   MItem.SubItems(1) = acctnames!accountnameeng
   MItem.SubItems(2) = acctnames!accountnamearab
   MItem.SubItems(3) = IIf(IsNull(acctnames!Budget) = True, 0, Format(acctnames!Budget, "###,###,###.#0"))
   Me.Combo1.AddItem acctnames!AccountCode
   Me.Combo2.AddItem acctnames!accountnameeng
   Me.Combo3.AddItem acctnames!accountnamearab
  End If
  acctnames.MoveNext
  Wend
End Sub

Private Sub Form_Resize()
On Error Resume Next
Me.ListView1.Width = Me.Width - 240
Me.ListView1.Height = Me.Height - 870
Me.SSTab1.Height = Me.Height - 450

Me.SSTab1.Width = Me.Width - 120
Me.Command1.Top = Me.SSTab1.Height - 500
Me.Command1.Left = Me.SSTab1.Width - 1500
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Me.ListView1.SortKey = ColumnHeader.Index - 1
Me.ListView1.Sorted = True
End Sub

Private Sub Timer1_Timer()
If Me.Combo1 = "" Or Me.Combo5 = "" Then
    Me.Command1.Enabled = False
   Else
   Me.Command1.Enabled = True
End If
   
End Sub

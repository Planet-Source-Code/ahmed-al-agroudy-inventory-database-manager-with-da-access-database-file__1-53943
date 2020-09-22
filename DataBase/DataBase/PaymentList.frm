VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form PaymentList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payment List"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   7815
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
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
      Left            =   6480
      TabIndex        =   1
      Top             =   6120
      Width           =   1095
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   7800
      _ExtentX        =   13758
      _ExtentY        =   9975
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
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   18
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Payable #"
         Object.Width           =   1676
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Cr Account Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Account Name"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Amount"
         Object.Width           =   1852
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Click an item you want pick."
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
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "PaymentList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim acctnames As ADODB.Recordset
Dim con As ADODB.Connection
Dim xtable As String
Dim sqltable As Boolean
Dim MItem As ListItem

Private Sub Command1_Click()
BankTransaction1.Combo16.SetFocus

 
Unload Me
End Sub

Private Sub Form_Activate()
Me.Command1.SetFocus
End Sub

Private Sub Form_Load()
 
 Dim xClass As New HabitatClass
 Set con = New ADODB.Connection
 Set acctnames = New ADODB.Recordset
 xtable = "Select * from payablesetup where confirmedmark='1' and Left(paymode,4)='Bank' order by serialno"
 sqltable = True
 If BankTransaction1.Check2.Enabled = False And BankTransaction1.Check2.Value = 0 Then
     xClass.GetTables acctnames, con, xtable, constring, sqltable
     While acctnames.EOF = False
         Set MItem = Me.ListView1.ListItems.Add(, , "")
         MItem.SubItems(1) = acctnames!SerialNo
         MItem.SubItems(2) = acctnames!AccNo
         MItem.SubItems(3) = acctnames!AccName
         MItem.SubItems(4) = Format(acctnames!amount, "###,###,###.#0")
         acctnames.MoveNext
    Wend
 Else
     Me.caption = "Cash Deposit List"
     xtable = "Select * from vouchers where left(payopt,3)='006' and accountnumber is not null and left(paymode,2)='01'and ItsDeposit<>'Yes'  order by receiptdate"
     xClass.GetTables acctnames, con, xtable, constring, sqltable
     Me.ListView1.ColumnHeaders(2).Text = "Receipt#"
     Me.ListView1.ColumnHeaders(3).Text = "AccountCode"
     Me.ListView1.ColumnHeaders(4).Text = "AccountName"
     Me.ListView1.ColumnHeaders(5).Text = "Amount"
     While acctnames.EOF = False
         Set MItem = Me.ListView1.ListItems.Add(, , "")
         
         MItem.SubItems(1) = acctnames!receiptno
         MItem.SubItems(2) = acctnames!accountnumber
         MItem.SubItems(3) = acctnames!accountname
         MItem.SubItems(4) = Format(acctnames!creditamount, "###,###,###.#0")
         acctnames.MoveNext
    Wend
 End If
 acctnames.Close
 
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Me.ListView1.SortKey = ColumnHeader.Index - 1
Me.ListView1.Sorted = True

End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
If Left(Me.caption, 4) <> "Cash" Then
    BankTransaction1.Combo16 = Me.ListView1.SelectedItem.SubItems(4)
    BankTransaction1.Text7 = Me.ListView1.SelectedItem.SubItems(1)
    BankTransaction1.Text8 = Me.ListView1.SelectedItem.SubItems(2)
    BankTransaction1.Text9 = Me.ListView1.SelectedItem.SubItems(3)
  Else
    BankTransaction1.Combo16 = Me.ListView1.SelectedItem.SubItems(4)
    BankTransaction1.Combo21 = Me.ListView1.SelectedItem.SubItems(1)
    BankTransaction1.Text8 = Me.ListView1.SelectedItem.SubItems(2)
    BankTransaction1.Text9 = Me.ListView1.SelectedItem.SubItems(3)
End If
End Sub

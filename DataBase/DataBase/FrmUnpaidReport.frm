VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FrmUnpaidReport 
   Caption         =   "UNPAID VOUCHER REPORT"
   ClientHeight    =   2520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5865
   Icon            =   "FrmUnpaidReport.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2520
   ScaleWidth      =   5865
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   2520
      TabIndex        =   12
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   0
      TabIndex        =   9
      Top             =   840
      Width           =   5775
      Begin VB.ComboBox cmbPayee 
         Height          =   315
         Left            =   1200
         TabIndex        =   10
         Top             =   240
         Width           =   3855
      End
      Begin VB.Label Label2 
         Caption         =   "By Payee"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Print All"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Done"
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   2040
      Width           =   1000
   End
   Begin VB.CheckBox Check1 
      Caption         =   "áÊÍæíá Çáí ÇáÚÑÈí "
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   5775
      Begin MSMask.MaskEdBox mskDueDate 
         Height          =   330
         Left            =   2400
         TabIndex        =   4
         Top             =   240
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   582
         _Version        =   393216
         MousePointer    =   1
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MskTo 
         Height          =   330
         Left            =   4080
         TabIndex        =   6
         Top             =   240
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   582
         _Version        =   393216
         MousePointer    =   1
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblTo 
         Caption         =   "To"
         Height          =   255
         Left            =   3480
         TabIndex        =   8
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblFrom 
         Caption         =   "From"
         Height          =   255
         Left            =   1800
         TabIndex        =   7
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "By DateDue"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "FrmUnpaidReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()

If Check1.Value = 1 Then

cmbPayee.RightToLeft = True
Check1.caption = "Set to Arabic"

Me.caption = "ÊÞÑíÑ ÇÐä ÕÑÝ ÛíÑ ãÏÝæÚ "
Me.Command1.caption = "Êã "
Label1.Left = 4080
Label2.Left = 4080
Label1.caption = "ÊÇÑíÎ ÇáÇÓÊÍÞÇÞ "
Label2.caption = "ÈæÇÓØÉ ÇáÏÝÚ "
Label1.Alignment = vbRightJustify
Label2.Alignment = vbRightJustify
mskDueDate.Left = 1200
MskTo.Left = 3000
cmbPayee.Left = 120
Me.RightToLeft = True
lblFrom.Left = 240
lblTo.Left = 2280






Else
cmbPayee.RightToLeft = False
Check1.caption = "áÊÍæíá Çáí ÇáÚÑÈí "
Me.RightToLeft = False
Me.caption = "UNPAID VOUCHER REPORT"
Me.Command1.caption = "&Done"
Label1.Left = 120
Label2.Left = 120
cmbPayee.Left = 1200
Label1.Alignment = vbLeftJustify
Label2.Alignment = vbLeftJustify
Label1.caption = "By DueDate"
Label2.caption = "By Payee"

lblFrom.Left = 1320
lblTo.Left = 3250
mskDueDate.Left = 2200
MskTo.Left = 4080



End If



If Me.Check1.Value = 1 Then
cmbPayee.Clear
Dim recpayee2 As New ADODB.Recordset
recpayee2.Open "select * from financemaster where substring(accountcode,1,3) = '131' order by accountnamearab", constring, adOpenKeyset, adLockOptimistic
While recpayee2.EOF = False

'If Trim(recpayee2!accountnamearab) = "" Then
   ' anu = recpayee2!accountnameeng
'Else
    anu = recpayee2!accountnamearab
'End If
cmbPayee.AddItem anu
recpayee2.MoveNext
Wend
recpayee2.Close

Else
cmbPayee.Clear

Dim recpayee22 As New ADODB.Recordset
recpayee22.Open "select * from financemaster where substring(accountcode,1,3) = '131' order by accountnameeng", constring, adOpenKeyset, adLockOptimistic
While recpayee22.EOF = False
    anu = recpayee22!accountnameeng
cmbPayee.AddItem anu
recpayee22.MoveNext
Wend
recpayee22.Close

End If



End Sub

Private Sub Command1_Click()

'Dim VarList
'If Check1.Value = 0 And cmbPayee.Text = "" And mskDueDate.Text = "__/__/____" Then
'MsgBox "Nothing is Selected", vbInformation, "Select the Date or Payee"
'Exit Sub
'End If
'
'If Check1.Value = 1 And cmbPayee.Text = "" And mskDueDate.Text = "__/__/____" Then
'MsgBox "íÈ ÓíÇÈäÓ ÓÈÊÓÈ ÔÊÈÔÊÇÈÊä ÓÔÇÈÓí", vbInformation, "íÓäÈ ÈÓÇÈ"
'Exit Sub
'End If
        

        
       ' VarList = FrmUnpaidReport.mskDueDate.Text
       ' DataEnvironment1.rsUnpaidByDueDate.Close
       ' DataEnvironment1.UnpaidByDueDate VarList
        
         ' ProcedByDate UnpaidByDateDue.Sections(1).Controls("label3")
          
         On Error Resume Next
         ProcedPrepBy UnpaidByDateDue.Sections(2).Controls("lblPrepby")
         UnpaidByDateDue.Show 1
         Unload Me
'        VarList = FrmUnpaidReport.cmbPayee.Text
'        On Error Resume Next
'        DataEnvironment1.rsUnPaidByPayee.Close
'        DataEnvironment1.UnpaidByPayee VarList
'
'          ProcedByDate2 UnpaidByPayee.Sections(1).Controls("label3")
'          'ProcedPrepBy UnpaidByPayee.Sections(1).Controls("lblPrepby")
'
'        UnpaidByPayee.Show 1
'        Unload Me
'

End Sub
Private Sub ProcedByDate(a As RptLabel)
a.caption = mskDueDate.Text
End Sub
Private Sub ProcedByDate2(X As RptLabel)
X.caption = cmbPayee.Text
End Sub

Private Sub ProcedPrepBy(b As RptLabel)
b.caption = cLogUser
End Sub

Private Sub Command2_Click()
        On Error Resume Next
       ' DataEnvironment1.rsUnpaidByDueDate.Close
       ' DataEnvironment1.UnpaidByDueDate VarList
        
         ' ProcedByDate UnpaidByDateDue.Sections(1).Controls("label3")
          ProcedPrepBy paidByDateDue.Sections(2).Controls("lblPrepby")
        
        paidByDateDue.Show 1
        Unload Me

End Sub

Private Sub Form_Load()
 
Dim recpayee As New ADODB.Recordset
recpayee.Open "select * from financemaster where substring(accountcode,1,3) = '131'", constring, adOpenKeyset, adLockOptimistic

While recpayee.EOF = False

If Me.Check1.Value = 1 Then

If Trim(recpayee!accountnamearab) = "" Then
    anu = recpayee!accountnameeng
Else
    anu = recpayee!accountnamearab
End If

Else


'If Trim(recpayee!accountnamearab) = "" Then
    anu = recpayee!accountnameeng
'Else
  '  anu = recpayee!accountnamearab
'End If

End If

    cmbPayee.AddItem anu
    recpayee.MoveNext

Wend
recpayee.Close

End Sub

Private Sub SSTab1_DblClick()

End Sub

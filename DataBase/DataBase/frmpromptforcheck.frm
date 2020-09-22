VERSION 5.00
Begin VB.Form frmpromptforcheck 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Check  ÇÎÊÇÑÇáÔíß "
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5430
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   5430
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1400
      Left            =   0
      TabIndex        =   0
      Top             =   -50
      Width           =   5415
      Begin VB.CommandButton cmdcancel 
         Caption         =   "&Cancel"
         Height          =   350
         Left            =   2760
         TabIndex        =   4
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "&Ok"
         Height          =   350
         Left            =   1440
         TabIndex        =   3
         Top             =   960
         Width           =   975
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   600
         Width           =   2175
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Check Number ÑÞã ÇáÔíß"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Bank Name ÇÓã ÇáÈäß"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmpromptforcheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim recbank As New ADODB.Recordset
Dim recche As New ADODB.Recordset
Dim con3 As New ADODB.Connection
Public dollarbank As String


Private Sub cmdcancel_Click()
Unload Me
'frmpaymentvou.comsetmode.ListIndex = 5
frmpaymentvou.txtchecknumber.Enabled = False
frmpaymentvou.txtcheckdate.Enabled = False
frmcheckdeposit.Show 1
End Sub

Private Sub cmdok_Click()
If Combo1.Text <> "" And Combo2.Text <> "" Then
frmpaymentvou.txtchecknumber.Text = Combo2.Text
frmpaymentvou.txtchecknumber.Enabled = False
frmpaymentvou.bankname = Trim(Combo1.Text)
If InStr(1, Combo1.Text, "$") > 0 Then
    dollarbank = 1
Else
    dollarbank = 0
End If
Unload Me
frmpaymentvou.txtcheckdate.SetFocus
Else
    MsgBox "Plase choose the Bank Name and Check Number", vbInformation, "Choose"
    Exit Sub
End If
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    frmpaymentvou.comsetmode.ListIndex = 5
    frmpaymentvou.txtchecknumber.Enabled = False
    frmpaymentvou.txtcheckdate.Enabled = False
    frmcheckdeposit.Show 1
    Unload Me
    Exit Sub
End If

Select Case KeyCode
Case vbKeyPageDown
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(Combo1.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Select

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Trim(Combo1.Text) <> "" Then
    con3.Open "dsN=fINANCE;UID=SA;PWD=;"
    recche.Open "Select * from checkregister where valuedate is Null and cancel <> '1' and bank= " & "'" & Trim(Combo1.Text) & "'" & " order by checknumber", con3, adOpenKeyset, adLockOptimistic, adCmdText
    hello = Combo1.Text
    mohamed = "HSBCD     HSBC BANK-HELIOPOLIS ($ ACCT.)"
    If hello = mohamed Then
        raja = 1
    End If
    If recche.BOF = True Then
        MsgBox "You Have No Any Assign check Plese Assign check Befor Make the check Payments", vbInformation, "Empty Checks"
        recche.Close
        con3.Close
        Combo2.Enabled = False
        Exit Sub
    Else
        recche.MoveFirst
        Combo2.Clear
        While recche.EOF = False
            Combo2.AddItem recche!checknumber
            recche.MoveNext
        Wend
        Combo2.Enabled = True
        Combo2.SetFocus
        recche.Close
        con3.Close
    End If
End If
End Sub

Private Sub Combo2_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyPageDown
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(Combo2.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Select

End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Trim(Combo2.Text) <> "" Then
    cmdOK.SetFocus
End If
End Sub

Private Sub Form_Activate()
On Error Resume Next
    Combo1.SetFocus
    Combo2.Enabled = False
On Error GoTo 0
End Sub

Private Sub Form_Load()
con3.Open "dsN=fINANCE;UID=SA;PWD=;"
recbank.Open "Select * from financemaster where substring(accountcode,1,7) = '1110201'", con3, adOpenKeyset, adLockOptimistic
If recbank.BOF = True Then
    MsgBox "Please add the Bank Details", vbInformation, "Empty Bank Account"
    Unload Me
    Exit Sub
Else
    recbank.MoveFirst
    While recbank.EOF = False
        Combo1.AddItem Trim(recbank!accountnameeng) ' & "\" & Trim(recbank!accountnamearab)
        recbank.MoveNext
    Wend
End If
con3.Close
'this is the table for registerd checks checkregister
End Sub


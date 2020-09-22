VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form PostingJournal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Please wait..."
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7785
   Icon            =   "PostingJournal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   7785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check1 
      Caption         =   "Post &Continuously"
      Height          =   255
      Left            =   5880
      TabIndex        =   7
      Top             =   840
      Width           =   1695
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   480
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PostingJournal.frx":0442
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Done Êã "
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   6600
      TabIndex        =   5
      Top             =   3960
      Width           =   1095
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2655
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   4683
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "TransDate"
         Object.Width           =   2734
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "No of Trans"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Total Debit Amt"
         Object.Width           =   2717
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Total Credit Amt"
         Object.Width           =   2717
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Status"
         Object.Width           =   2911
      EndProperty
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   "this Textbox use to control what type of journal to post"
      Top             =   3480
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.PictureBox Picture1 
      Height          =   350
      Left            =   120
      ScaleHeight     =   285
      ScaleWidth      =   7515
      TabIndex        =   1
      Top             =   360
      Width           =   7575
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   330
         Left            =   -10
         TabIndex        =   2
         Top             =   -15
         Width           =   7550
         _ExtentX        =   13335
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Progress:ÇáÈÑäÇãÌ"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   5775
   End
End
Attribute VB_Name = "PostingJournal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsJRN As New ADODB.Recordset
Dim DayNo As Integer
Dim trandate As Date
Dim MItem As ListItem
Dim i As Long
Dim cTotrec As Long
Dim Taken As Long
Dim cVal As Long
Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Activate()
OpenJOurnals
End Sub
Sub OpenJOurnals()
Dim rsDate1 As Date
DayNo = DayNo + 1
trDate = Me.ListView1.ListItems.Item(DayNo).Text
trDate1 = Mid(trDate, 4, 2) & "/" & Left(trDate, 2) & "/" & Mid(trDate, 7, 4)
Dim strFindMe As String
Dim itmFound As ListItem   ' FoundItem variable.
intSelectedOption = lvwText
strFindMe = trDate
Set itmFound = Me.ListView1.Finditem(strFindMe, intSelectedOption, , lvwPartial)
If itmFound Is Nothing Then  ' If no match, inform user and exit.
 Else
  itmFound.EnsureVisible
  itmFound.Selected = True   ' Select the ListItem.
End If


If Trim(Me.Text1.Text) = "IVY" Then
    trDate1 = Mid(trDate, 4, 2) & "/" & Left(trDate, 2) & "/" & Mid(trDate, 7, 4)
    rsJRN.Open "Select * from InventoryJournal where transdate = " & "'" & trDate1 & "'" & "and remarks is null", constring, adOpenKeyset, adLockPessimistic, adCmdText
    cTotrec = Int(rsJRN.RecordCount / 100)
    If cTotrec = 0 Then
     cTotrec = 1
    End If
    i = 0
    cVal = 0
    PostingJournal.caption = "Inventory Journal Posting"
    PostIVY (cTotrec)

ElseIf Trim(Me.Text1.Text) = "GEN" Then
    rsJRN.Open "Select * from GenJOurnalTrans where transdate = " & "'" & trDate1 & "'" & "and remarks is null", constring, adOpenKeyset, adLockPessimistic, adCmdText
    cTotrec = Int(rsJRN.RecordCount / 100)
    If cTotrec = 0 Then
     cTotrec = 1
    End If
    i = 0
    cVal = 0
    PostingJournal.caption = "General Journal Posting"
    PostGEN (cTotrec)

ElseIf Trim(Me.Text1.Text) = "AST" Then
    rsJRN.Open "Select * from AssetJOurnal where Transdate= " & "'" & trDate1 & "'" & "and  remarks is null", constring _
                   , adOpenKeyset, adLockPessimistic, adCmdText
    cTotrec = Int(rsJRN.RecordCount / 100)
    If cTotrec = 0 Then
     cTotrec = 1
    End If
    i = 0
    cVal = 0
    PostingJournal.caption = "Assets Journal Posting"
    PostAST (cTotrec)

ElseIf Trim(Me.Text1.Text) = "BNK" Then
    rsJRN.Open "Select * from BankJOurnal where transdate = " & "'" & trDate1 & "'" & "and remarks is null", constring, adOpenKeyset, adLockPessimistic, adCmdText
    cTotrec = Int(rsJRN.RecordCount / 100)
    If cTotrec = 0 Then
     cTotrec = 1
    End If
    i = 0
    cVal = 0
    PostingJournal.caption = "Bank Journal Posting"
    PostBNK (cTotrec)


ElseIf Trim(Me.Text1.Text) = "SAL" Then
    rsJRN.Open "Select * from SalesJOurnal where Transdate= " & "'" & trDate1 & "'" & "and remarks is null", constring, _
                        adOpenKeyset, adLockPessimistic, adCmdText
    cTotrec = Int(rsJRN.RecordCount / 100)
    If cTotrec = 0 Then
     cTotrec = 1
    End If
    i = 0
    cVal = 0
    PostingJournal.caption = "Local Sales Journal Posting"
    PostSAL

ElseIf Trim(Me.Text1.Text) = "CSR" Then ' for anushath receipt
   'trDate1 = Mid(trDate, 4, 2) & "/" & Left(trDate, 2) & "/" & Mid(trDate, 7, 4)
    rsJRN.Open "Select * from CASHJournal where transdate = " & "'" & trDate1 & "'" & "and remarks is null and tranType = 'R' ", constring, adOpenKeyset, adLockPessimistic, adCmdText
    cTotrec = Int(rsJRN.RecordCount / 100)
    If cTotrec = 0 Then
     cTotrec = 1
    End If
    i = 0
    cVal = 0
    PostingJournal.caption = "Cash Receipt Journal Posting"
    PostCSR (cTotrec)

ElseIf Trim(Me.Text1.Text) = "CSP" Then ' for anushath Payment
   'trDate1 = Mid(trDate, 4, 2) & "/" & Left(trDate, 2) & "/" & Mid(trDate, 7, 4)
    rsJRN.Open "Select * from CASHJournal where transdate = " & "'" & trDate1 & "'" & "and remarks is null and tranType = 'P' ", constring, adOpenKeyset, adLockPessimistic, adCmdText
    cTotrec = Int(rsJRN.RecordCount / 100)
    If cTotrec = 0 Then
     cTotrec = 1
    End If
    i = 0
    cVal = 0
    PostingJournal.caption = "Cash Payment Journal Posting"
    PostCSP (cTotrec)

ElseIf Trim(Me.Text1.Text) = "PYB" Then

    rsJRN.Open "Select * from Payjournal where confirmeddate = " & "'" & trDate1 & "'" & "and remarks is null", constring, adOpenKeyset, adLockPessimistic, adCmdText
    cTotrec = Int(rsJRN.RecordCount / 100)
    If cTotrec = 0 Then
     cTotrec = 1
    End If
    i = 0
    cVal = 0
    PostingJournal.caption = "Payable Journal Posting"
    Call PostPyb(cTotrec)

ElseIf Trim(Me.Text1.Text) = "PTC" Then

    rsJRN.Open "Select * from PettyJournal where confirmeddate = " & "'" & trDate1 & "'" & "and Postmark is null", constring, adOpenKeyset, adLockPessimistic, adCmdText
    cTotrec = Int(rsJRN.RecordCount / 100)
    If cTotrec = 0 Then
     cTotrec = 1
    End If
    i = 0
    cVal = 0
    PostingJournal.caption = "Petty CAsh Journal Posting"
    Call PostPetty(cTotrec)



ElseIf Trim(Me.Text1.Text) = "SRL" Then ' for anushath credit note
    rsJRN.Open "Select * from Creditnote where transdate = " & "'" & trDate1 & "'" & "and remarks is null", constring, adOpenKeyset, adLockPessimistic, adCmdText
    cTotrec = Int(rsJRN.RecordCount / 100)
    If cTotrec = 0 Then
     cTotrec = 1
    End If
    i = 0
    cVal = 0
    PostingJournal.caption = "Credit Note Journal Posting"
    PostSRL (cTotrec)

ElseIf Trim(Me.Text1.Text) = "SPA" Then ' for anushath credit note
    rsJRN.Open "Select * from debitnote where transdate = " & "'" & trDate1 & "'" & "and remarks is null", constring, adOpenKeyset, adLockPessimistic, adCmdText
    cTotrec = Int(rsJRN.RecordCount / 100)
    If cTotrec = 0 Then
     cTotrec = 1
    End If
    i = 0
    cVal = 0
    PostingJournal.caption = "Credit Note Journal Posting"
    PostSPA (cTotrec)

End If


End Sub
Sub PostIVY(cTotrec As Long)
        Dim rstGLMaster As New ADODB.Recordset
        Dim FinanceMAster As New ADODB.Recordset
        rstGLMaster.Open "GLMaster", constring, adOpenKeyset, adLockPessimistic, adCmdTable
        rstGLMaster.MoveLast
        If rstGLMaster.EOF = True Then
            cBeginBal = 0
          Else
          cBeginBal = rstGLMaster!Balance
        End If
        i = 0
        Me.Label1.caption = "Status: Posting " & Trim(Me.ListView1.ListItems.Item(DayNo)) & " Transactions..."
        Do Until rsJRN.EOF = True
              i = i + 1
              If i = cTotrec Then
                cVal = cVal + 1
                i = 0
                If cVal <= 100 Then
                    Me.ProgressBar1.Value = cVal
                    Me.ListView1.ListItems.Item(DayNo).SubItems(4) = cVal & "% Process"
                End If
              End If
              Ac = rsJRN!accountnumber
              Dramt = rsJRN!DebitAmount
              CrAmt = rsJRN!creditamount

              'FinanceMAster.Open "FInanceMASter", conString, adOpenKeyset, adLockPessimistic, adCmdTable
              'FinanceMAster.Close
              FinanceMAster.Open "SElect * From FInanceMASter where AccountCode = " & "'" & Ac & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
              With FinanceMAster
               If Val(Dramt) <> 0 Then
                 If Trim(!LastTransdate) <> Trim(Format(Date, "dd/mm/yyyy")) Then
                   !BeginBal = !EndingBal
                   !Debit = Dramt
                   Else
                   !Debit = !Debit + Dramt
                  End If
                  !EndingBal = !BeginBal + Dramt
                  !TotalDebit = !TotalDebit + Dramt
                  !LastTransdate = Format(Date, "dd/mm/yyyy")
                  !LastTransType = "Debit"
                  !TotalTrans = !TotalTrans + 1
                  .Update
                  .close
                End If
               If Val(CrAmt) <> 0 Then
                 If Trim(!LastTransdate) <> Trim(Format(Date, "dd/mm/yyyy")) Then
                  !BeginBal = !EndingBal
                   !Credit = CrAmt
                   Else
                   !Credit = !Credit + CrAmt
                  End If
                  !EndingBal = !BeginBal + CrAmt
                  !TotalCredit = !TotalCredit + CrAmt
                  !LastTransdate = Format(Date, "dd/mm/yyyy")
                  !LastTransType = "Credit"
                  !TotalTrans = !TotalTrans + 1
                  .Update
                  .close
                End If
              End With


            With rstGLMaster
                .addnew
                !JOurnalNo = rsJRN!SerialNo
                !AccountCode = rsJRN!accountnumber
                !accountname = rsJRN!accountname
                !PostDate = Date & " " & Time
                !recorddate = rsJRN!TRansDate
                !Particulars = rsJRN!Description & "/ Posted by : " & cLogUser
                !DebitAmount = Dramt
                !creditamount = CrAmt
                If Val(Dramt) <> 0 Then
                    !Balance = cBeginBal + Dramt
                 Else
                    !Balance = cBeginBal - CrAmt
                End If
                cBeginBal = !Balance
                .Update
             End With
            Dramt = 0
            CrAmt = 0
            rsJRN!remarks = "Posted Last " & Date & " " & Time
            rsJRN!Status = "Posted"
            rsJRN.Update
           rsJRN.MoveNext
           DoEvents
        Loop
        rsJRN.close
        'update the JOurnalCode
        Dim rstJOurcode As New ADODB.Recordset
        rstJOurcode.Open "Select * from JOurnalCode order by Code", constring, adOpenKeyset, adLockPessimistic, adCmdText
        rstJOurcode.Move 4, 1
        rstJOurcode!lastpostingdate = Format(Date, "mm/dd/yyyy")
        rstJOurcode.Update
        rstJOurcode.close
        POstNextTrans
End Sub
Sub POstNextTrans()
    If cVal < 100 Then
        Me.ProgressBar1.Value = 100
    End If
    On Error Resume Next
        Me.ListView1.ListItems.Item(DayNo).Checked = True
        Me.ListView1.ListItems.Item(DayNo).SubItems(4) = "Posted!!"
        On Error GoTo 0
        If (DayNo + 1) <= Me.ListView1.ListItems.Count Then
          If Me.Check1.Value = 0 Then
           mess = MsgBox(Me.ListView1.ListItems.Item(DayNo) & " Transaction is Finished! ÇáÚãáíÇÊ ÇáäåÇÆíÉ  " & vbCrLf _
                      & "Do you want to continue for the " & Me.ListView1.ListItems.Item(DayNo + 1) & " transactions?", vbQuestion + vbYesNo, "Please confirm ãä ÝÖáß ÇáÊÇßíÏ")
           Else
           mess = vbYes
          End If
        Else
          Me.Enabled = True
          Me.Command1.Enabled = True
          mess = MsgBox("All transactions are completely posted! ", vbInformation + vbOKOnly, "Message")
          DayNo = 0
          Me.Label1.caption = "Status: Succesfully posted! áÕÞ ÈäÌÇÍ"
          Exit Sub
        End If
        If mess = vbYes Then
            OpenJOurnals
           Else
           DayNo = 0
         Unload Me
        End If
End Sub

Sub PostAST(cTotrec As Long)
        Dim rstGLMaster As New ADODB.Recordset
        Dim FinanceMAster As New ADODB.Recordset
        rstGLMaster.Open "GLMaster", constring, adOpenKeyset, adLockPessimistic, adCmdTable
        rstGLMaster.MoveLast
        If rstGLMaster.EOF = True Then
            cBeginBal = 0
          Else
          cBeginBal = rstGLMaster!Balance
        End If
        i = 0
        Me.Label1.caption = "Status: Posting " & Trim(Me.ListView1.ListItems.Item(DayNo)) & " Transactions..."
        Do Until rsJRN.EOF = True
              i = i + 1
              If i = cTotrec Then
                cVal = cVal + 1
                i = 0
                If cVal <= 100 Then
                    Me.ProgressBar1.Value = cVal
                    Me.ListView1.ListItems.Item(DayNo).SubItems(4) = cVal & "% Process"
                End If
              End If
              Ac = rsJRN!accountnumber
              Dramt = rsJRN!DebitAmount
              CrAmt = rsJRN!creditamount

              'FinanceMAster.Open "FInanceMASter", conString, adOpenKeyset, adLockPessimistic, adCmdTable
              'FinanceMAster.Close
              FinanceMAster.Open "SElect * From FInanceMASter where AccountCode = " & "'" & Ac & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
              With FinanceMAster
               If Val(Dramt) <> 0 Then
                 If Trim(!LastTransdate) <> Trim(Format(Date, "dd/mm/yyyy")) Then
                   !BeginBal = !EndingBal
                   !Debit = Dramt
                   Else
                   !Debit = !Debit + Dramt
                  End If
                  !EndingBal = !BeginBal + Dramt
                  !TotalDebit = !TotalDebit + Dramt
                  !LastTransdate = Format(Date, "dd/mm/yyyy")
                  !LastTransType = "Debit"
                  !TotalTrans = !TotalTrans + 1
                  .Update
                  .close
                End If
               If Val(CrAmt) <> 0 Then
                 If Trim(!LastTransdate) <> Trim(Format(Date, "dd/mm/yyyy")) Then
                  !BeginBal = !EndingBal
                   !Credit = CrAmt
                   Else
                   !Credit = !Credit + CrAmt
                  End If
                  !EndingBal = !BeginBal + CrAmt
                  !TotalCredit = !TotalCredit + CrAmt
                  !LastTransdate = Format(Date, "dd/mm/yyyy")
                  !LastTransType = "Credit"
                  !TotalTrans = !TotalTrans + 1
                  .Update
                  .close
                End If
              End With


            With rstGLMaster
                .addnew
                !JOurnalNo = rsJRN!SerialNo
                !AccountCode = rsJRN!accountnumber
                !accountname = rsJRN!accountname
                !PostDate = Date & " " & Time
                !recorddate = rsJRN!TRansDate
                !Particulars = rsJRN!Description & "/ Posted by : " & cLogUser
                !DebitAmount = Dramt
                !creditamount = CrAmt
                If Val(Dramt) <> 0 Then
                    !Balance = cBeginBal + Dramt
                 Else
                    !Balance = cBeginBal - CrAmt
                End If
                cBeginBal = !Balance
                .Update
             End With
            Dramt = 0
            CrAmt = 0
            rsJRN!remarks = "Posted Last " & Date & " " & Time
            rsJRN.Update
           rsJRN.MoveNext
           DoEvents
        Loop
        rsJRN.close
        POstNextTrans
End Sub
Sub PostSAL()
            Dim rstGLMaster As New ADODB.Recordset
            Dim FinanceMAster As New ADODB.Recordset
            Dim rstJOurcode As New ADODB.Recordset
            rstJOurcode.Open "Select * from JOurnalCode where Code='SAL'", constring, adOpenKeyset, adLockPessimistic, adCmdText
            xCode = rstJOurcode!JOurnalName
            trandate = rstJOurcode!lastpostingdate
            rstJOurcode.close
            rstGLMaster.Open "GLMaster", constring, adOpenKeyset, adLockPessimistic, adCmdTable
            Me.Label1.caption = "Status: Posting " & Trim(Me.ListView1.ListItems.Item(DayNo)) & " Transactions..."
            Do Until rsJRN.EOF = True
                    i = i + 1

                    If i = cTotrec Then
                      cVal = cVal + 1
                      i = 0
                      If cVal <= 100 Then
                          Me.ProgressBar1.Value = cVal
                          On Error Resume Next
                          Me.ListView1.ListItems.Item(DayNo).SubItems(4) = cVal & "% Process"
                          On Error GoTo 0
                      End If
                    End If


                 Ac = Trim(rsJRN!accountnumber)
                 If Ac <> "" Then
                  TR = IIf(IsNull(rsJRN!tradercvble) = True, 0, rsJRN!tradercvble)
                  If Trim(Ac) <> "" Then 'Or IsNull(AC) = False Then
                     On Error Resume Next
                     FinanceMAster.close
                     FinanceMAster.Open "SElect * From FInanceMASter where AccountCode = " & "'" & Ac & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
                     an = FinanceMAster!accountnameeng
                     On Error GoTo 0
                    Else
                     'post the TR to Acct# 111041102001
                      Ac = "111041102001" 'for blank Acctno will post to Miscellaneous Acct
                      On Error Resume Next
                      FinanceMAster.close
                      FinanceMAster.Open "SElect * From FInanceMASter where AccountCode = " & "'" & Ac & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
                      an = FinanceMAster!accountnameeng
                      On Error GoTo 0
                   End If


                  If Val(TR) <> 0 Then
                     With FinanceMAster
                       If Trim(!LastTransdate) <> Trim(Format(Date, "mm/dd/yyyy")) Then
                        !BeginBal = !EndingBal
                        !Debit = TR
                         Else
                         !Debit = !Debit + TR
                       End If
                       !EndingBal = !BeginBal + !Debit  '!Debit - !Credit
                       !TotalDebit = !TotalDebit + TR
                       !LastTransdate = Format(Date, "mm/dd/yyyy")
                       !LastTransType = "Debit"
                       !TotalTrans = !TotalTrans + 1
                       .Update
                     End With
                     FinanceMAster.close
                     With rstGLMaster
                        If .RecordCount <> 0 Then
                            .MoveFirst
                            .MoveLast
                            cBeginBal = !Balance
                         ElseIf .EOF = True And .RecordCount = 0 Then
                            cBeginBal = 0
                            ElseIf .EOF <> True And .RecordCount <> 0 Then
                           '.MoveLast
                          cBeginBal = !Balance
                        End If
                          .addnew
                          !JOurnalNo = rsJRN!SerialNo
                          !AccountCode = Ac
                          !accountname = an
                          !PostDate = Date & " " & Time
                          !recorddate = rsJRN!TRansDate
                          !Particulars = rsJRN!Description & " " & rsJRN!invoicedate & "/ Posted by : " & cLogUser
                          !DebitAmount = TR
                          !creditamount = 0
                          !Balance = cBeginBal + TR
                          .Update
                     End With
                   End If


                  '-----------------------------------------
                  'post the TDA to Acct# 23201010000
                  TDA = IIf(IsNull(rsJRN!tradeDiscamt) = True, 0, rsJRN!tradeDiscamt)
                  If Val(TDA) <> 0 Then
                    If Trim(rsJRN!profitcenter) = "ZF02" Then
                      Ac = "232010101000"
                     ElseIf Trim(rsJRN!profitcenter) = "ZWO6" Then
                      Ac = "232010301000"
                     ElseIf Trim(rsJRN!profitcenter) = "ZMA1" Then
                      Ac = "232010201000"
                     ElseIf Trim(rsJRN!profitcenter) = "ZUP3" Then
                      Ac = "232010401000"
                    End If
                    FinanceMAster.Open "SElect * From FInanceMASter where AccountCode = " & "'" & Ac & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
                    an = FinanceMAster!accountnameeng
                     With FinanceMAster
                       If Trim(!LastTransdate) <> Trim(Format(Date, "mm/dd/yyyy")) Then
                          !BeginBal = !EndingBal
                          !Debit = TDA
                       Else
                          !Debit = !Debit + TDA
                       End If
                       !EndingBal = !BeginBal + !Debit
                       !TotalDebit = !TotalDebit + TDA
                       !LastTransdate = Format(Date, "mm/dd/yyyy")
                       !LastTransType = "Debit"
                       !TotalTrans = !TotalTrans + 1
                       .Update
                       .close
                      End With
                      With rstGLMaster
                        .Requery
                        If .EOF = True And .RecordCount <> 0 Then
                            .MoveFirst
                            .MoveLast
                            cBeginBal = !Balance
                           ElseIf .EOF = True And .RecordCount = 0 Then
                            cBeginBal = 0
                            Else
                           '.MoveNext
                           .MoveLast
                          cBeginBal = !Balance
                        End If
                          .addnew
                          !JOurnalNo = rsJRN!SerialNo
                          !AccountCode = Ac
                          !accountname = an
                          !PostDate = Date & " " & Time
                          !recorddate = rsJRN!TRansDate
                          !Particulars = rsJRN!Description & " " & rsJRN!invoicedate & "/ Posted by : " & cLogUser
                          !DebitAmount = TDA
                          !creditamount = 0
                          !Balance = cBeginBal + TDA
                          .Update
                     End With
                    End If


                    '-------------------------------------------
                   'post the GS to Acct#  211010100000
                  GS = IIf(IsNull(rsJRN!GrossSales) = True, 0, rsJRN!GrossSales)
                  If Val(GS) <> 0 Then
                     If Trim(rsJRN!profitcenter) = "ZFO2" Then
                      Ac = "211010100000"
                     ElseIf Trim(rsJRN!profitcenter) = "ZWO6" Then
                      Ac = "211030100000"
                     ElseIf Trim(rsJRN!profitcenter) = "ZMA1" Then
                      Ac = "211020100000"
                     ElseIf Trim(rsJRN!profitcenter) = "ZUP3" Then
                      Ac = "211040100000"
                     End If
                     FinanceMAster.Open "SElect * From FInanceMASter where AccountCode = " & "'" & Ac & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
                     an = FinanceMAster!accountnameeng
                     With FinanceMAster
                       If Trim(!LastTransdate) <> Trim(Format(Date, "mm/dd/yyyy")) Then
                          !BeginBal = !EndingBal
                          !Credit = GS
                        Else
                          !Credit = !Credit + GS
                        End If
                       !EndingBal = !BeginBal + !Credit
                       !TotalCredit = !TotalCredit + GS
                       !LastTransdate = Format(Date, "mm/dd/yyyy")
                       !LastTransType = "Credit"
                       !TotalTrans = !TotalTrans + 1
                       .Update
                       '.Close
                      End With
                      FinanceMAster.close
                      With rstGLMaster
                         .Requery
                        If .EOF = True And .RecordCount <> 0 Then
                            .MoveFirst
                            .MoveLast
                            cBeginBal = !Balance
                           ElseIf .EOF = True And .RecordCount = 0 Then
                            cBeginBal = 0
                            Else
                           .MoveLast
                          cBeginBal = !Balance
                        End If
                          .addnew
                          !JOurnalNo = rsJRN!SerialNo
                          !AccountCode = Ac
                          !accountname = an
                          !PostDate = Date & " " & Time
                          !recorddate = rsJRN!TRansDate
                          !Particulars = rsJRN!Description & " " & rsJRN!invoicedate & "/ Posted by : " & cLogUser
                          !DebitAmount = 0
                          !creditamount = GS
                          !Balance = cBeginBal + GS
                          .Update
                     End With
                   End If

                  '-------------------------------------------
                  'post the VAT to Acct# 13204010000
                  vat = IIf(IsNull(rsJRN!vat) = True, 0, rsJRN!vat)
                  If Val(vat) <> 0 Then
                     Ac = "131040203001"
                     FinanceMAster.Open "SElect * From FInanceMASter where AccountCode = " & "'" & Ac & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
                     an = FinanceMAster!accountnameeng
                     With FinanceMAster
                       If Trim(!LastTransdate) <> Trim(Format(Date, "mm/dd/yyyy")) Then
                          !BeginBal = !EndingBal
                          !Credit = vat
                       Else
                          !Credit = !Credit + vat
                       End If
                       !EndingBal = !BeginBal + !Credit
                       !TotalCredit = !TotalCredit + vat
                       !LastTransdate = Format(Date, "mm/dd/yyyy")
                       !LastTransType = "Credit"
                       !TotalTrans = !TotalTrans + 1
                       .Update
                       '.Close
                     End With
                     FinanceMAster.close
                     With rstGLMaster
                        .Requery
                        If .EOF = True And .RecordCount <> 0 Then
                            .MoveFirst
                            .MoveLast
                            cBeginBal = !Balance
                           ElseIf .EOF = True And .RecordCount = 0 Then
                            cBeginBal = 0
                           Else
                           .MoveLast
                          cBeginBal = !Balance
                        End If
                          .addnew
                          !JOurnalNo = rsJRN!SerialNo
                          !AccountCode = Ac
                          !accountname = an
                          !PostDate = Date & " " & Time
                          !recorddate = rsJRN!TRansDate
                          !Particulars = rsJRN!Description & " " & rsJRN!invoicedate & "/ Posted by : " & cLogUser
                          !creditamount = vat
                          !DebitAmount = 0
                          !Balance = cBeginBal + vat
                          .Update
                     End With
                  End If



                  '-----------------------------------
                  'post the SUR Tax to Acct# 13204020000
                  STA = IIf(IsNull(rsJRN!SurTaxAmt) = True, 0, rsJRN!SurTaxAmt)
                  If Val(STA) <> 0 Then
                     Ac = "131040202001"
                     FinanceMAster.Open "SElect * From FInanceMASter where AccountCode = " & "'" & Ac & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
                     an = FinanceMAster!accountnameeng
                     With FinanceMAster
                       If Trim(!LastTransdate) <> Trim(Format(Date, "mm/dd/yyyy")) Then
                          !BeginBal = !EndingBal
                          !Credit = STA
                       Else
                          !Credit = !Credit + STA
                       End If
                       !EndingBal = !BeginBal + !Credit
                       !TotalCredit = !TotalCredit + STA
                       !LastTransdate = Format(Date, "mm/dd/yyyy")
                       !LastTransType = "Credit"
                       !TotalTrans = !TotalTrans + 1
                       .Update
                       .close
                     End With
                     With rstGLMaster
                        .Requery
                        If .EOF = True And .RecordCount <> 0 Then
                            .MoveFirst
                            .MoveLast
                            cBeginBal = !Balance
                           ElseIf .EOF = True And .RecordCount = 0 Then
                            cBeginBal = 0
                           Else
                           .MoveLast
                          cBeginBal = !Balance
                        End If
                          .addnew
                          !JOurnalNo = rsJRN!SerialNo
                          !AccountCode = Ac
                          !accountname = an
                         !PostDate = Date & " " & Time
                          !recorddate = rsJRN!TRansDate
                          !Particulars = rsJRN!Description & " " & rsJRN!invoicedate & " / Posted by :" & cLogUser
                          !DebitAmount = 0
                          !creditamount = STA
                          !Balance = cBeginBal + STA
                          .Update
                     End With
                  End If


                  '--------------------------------
                  'post the TRansport Charges
                  TC = IIf(rsJRN!transpoCharge = 0, 0, rsJRN!transpoCharge)
                  If (TC) <> 0 Then
                     If Trim(rsJRN!profitcenter) = "ZFO2" Then
                      Ac = "216010101000"
                      ElseIf Trim(rsJRN!profitcenter) = "ZWO6" Then
                      Ac = "216010301000"
                      ElseIf Trim(rsJRN!profitcenter) = "ZMA1" Then
                      Ac = "216010201000"
                      ElseIf Trim(rsJRN!profitcenter) = "ZUP3" Then
                      Ac = "216010401000"
                     End If
                     FinanceMAster.Open "SElect * From FInanceMASter where AccountCode = " & "'" & Ac & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
                     an = FinanceMAster!accountnameeng
                     With FinanceMAster
                       If Trim(!LastTransdate) <> Trim(Format(Date, "mm/dd/yyyy")) Then
                          !BeginBal = !EndingBal
                          !Credit = TC
                       Else
                          !Credit = !Credit + TC
                       End If
                       !EndingBal = !BeginBal + !Credit
                       !TotalCredit = !TotalCredit + TC
                       !LastTransdate = Format(Date, "mm/dd/yyyy")
                       !LastTransType = "Credit"
                       !TotalTrans = !TotalTrans + 1
                       .Update
                       .close
                     End With
                     With rstGLMaster
                        If .EOF = True And .RecordCount <> 0 Then
                            .MoveFirst
                            .MoveLast
                            cBeginBal = !Balance
                           ElseIf .EOF = True And .RecordCount = 0 Then
                            cBeginBal = 0
                            Else
                           .MoveLast
                          cBeginBal = !Balance
                        End If
                          .addnew
                          !JOurnalNo = rsJRN!SerialNo
                          !AccountCode = Ac
                          !accountname = an
                          !PostDate = Date & " " & Time
                          !recorddate = rsJRN!TRansDate
                          !Particulars = rsJRN!Description & " " & rsJRN!invoicedate & "/ Posted by : " & cLogUser
                          !DebitAmount = 0
                          !creditamount = TC
                          !Balance = cBeginBal + TC
                          .Update
                     End With
                  End If
                 'End If


                 rsJRN!remarks = "Posted Last " & Date & " " & Time
                rsJRN.Update
              End If
              rsJRN.MoveNext
             DoEvents
            Loop
           rsJRN.close
           POstNextTrans
End Sub

Sub PostGEN(cTotrec As Long)
        Dim rstGLMaster As New ADODB.Recordset
        Dim FinanceMAster As New ADODB.Recordset
        rstGLMaster.Open "GLMaster", constring, adOpenKeyset, adLockPessimistic, adCmdTable
        If rstGLMaster.EOF = False Then
          rstGLMaster.MoveLast
        End If
        If rstGLMaster.EOF = True Then
            cBeginBal = 0
          Else
          cBeginBal = rstGLMaster!Balance
        End If
        i = 0
        Me.Label1.caption = "Status: Posting " & Trim(Me.ListView1.ListItems.Item(DayNo)) & " Transactions..."
        Do Until rsJRN.EOF = True
              i = i + 1
              If i = cTotrec Then
                cVal = cVal + 1
                i = 0
                If cVal <= 100 Then
                    Me.ProgressBar1.Value = cVal
                    Me.ListView1.ListItems.Item(DayNo).SubItems(4) = cVal & "% Process"
                End If
              End If
              Ac = rsJRN!accountnumber
              Dramt = rsJRN!DebitAmount
              CrAmt = rsJRN!creditamount

              'FinanceMAster.Open "FInanceMASter", conString, adOpenKeyset, adLockPessimistic, adCmdTable
              'FinanceMAster.Close
              FinanceMAster.Open "SElect * From FInanceMASter where AccountCode = " & "'" & Ac & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
              With FinanceMAster
               If Val(Dramt) <> 0 Then
                 If Trim(!LastTransdate) <> Trim(Format(Date, "dd/mm/yyyy")) Then
                   !BeginBal = !EndingBal
                   !Debit = Dramt
                   Else
                   !Debit = !Debit + Dramt
                  End If
                  !EndingBal = !BeginBal + Dramt
                  !TotalDebit = !TotalDebit + Dramt
                  !LastTransdate = Format(Date, "mm/dd/yyyy")
                  !LastTransType = "Debit"
                  !TotalTrans = !TotalTrans + 1
                  .Update
                  .close
                End If
               If Val(CrAmt) <> 0 Then
                 If Trim(!LastTransdate) <> Trim(Format(Date, "dd/mm/yyyy")) Then
                  !BeginBal = !EndingBal
                   !Credit = CrAmt
                   Else
                   !Credit = !Credit + CrAmt
                  End If
                  !EndingBal = !BeginBal + CrAmt
                  !TotalCredit = !TotalCredit + CrAmt
                  !LastTransdate = Format(Date, "mm/dd/yyyy")
                  !LastTransType = "Credit"
                  !TotalTrans = !TotalTrans + 1
                  .Update
                  .close
                End If
              End With


            With rstGLMaster
                .addnew
                !JOurnalNo = rsJRN!SerialNo
                !AccountCode = rsJRN!accountnumber
                !accountname = rsJRN!accountname
                !PostDate = Date & " " & Time
                !recorddate = rsJRN!TRansDate
                !Particulars = rsJRN!Description & "/ Posted by : " & cLogUser & "(Processed by " & rsJRN!Prepby & "-" & rsJRN!TRansDate & ")"
                !DebitAmount = Dramt
                !creditamount = CrAmt
                If Val(Dramt) <> 0 Then
                    !Balance = cBeginBal + Dramt
                 Else
                    !Balance = cBeginBal - CrAmt
                End If
                cBeginBal = !Balance
                .Update
             End With
            Dramt = 0
            CrAmt = 0
            rsJRN!remarks = "Posted Last " & Date & " " & Time
            rsJRN!Status = "Posted"
            rsJRN.Update
           rsJRN.MoveNext
           DoEvents
        Loop
        rsJRN.close
        POstNextTrans
End Sub
Sub PostCSR(cTotrec As Long)
        Dim rstGLMaster As New ADODB.Recordset
        Dim FinanceMAster As New ADODB.Recordset
        rstGLMaster.Open "GLMaster", constring, adOpenKeyset, adLockPessimistic, adCmdTable
        If rstGLMaster.EOF = True Then
            cBeginBal = 0
          Else
          cBeginBal = rstGLMaster!Balance
        End If
        i = 0
        Me.Label1.caption = "Status: Posting " & Trim(Me.ListView1.ListItems.Item(DayNo)) & " Transactions..."
        Do Until rsJRN.EOF = True
              i = i + 1
              If i = cTotrec Then
                cVal = cVal + 1
                i = 0
                If cVal <= 100 Then
                    Me.ProgressBar1.Value = cVal
                    Me.ListView1.ListItems.Item(DayNo).SubItems(4) = cVal & "% Process"
                End If
              End If
              Ac = rsJRN!accountnumber
              Dramt = rsJRN!DebitAmount
              CrAmt = rsJRN!creditamount

              'FinanceMAster.Open "FInanceMASter", conString, adOpenKeyset, adLockPessimistic, adCmdTable
              On Error Resume Next
              FinanceMAster.close
              On Error GoTo 0
              
              FinanceMAster.Open "SElect * From FInanceMASter where AccountCode = " & "'" & Ac & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
              'MsgBox FinanceMAster.RecordCount
              With FinanceMAster
              If FinanceMAster.BOF = False Then
               If Val(Dramt) <> 0 Then
                 If Trim(!LastTransdate) <> Trim(Format(Date, "dd/mm/yyyy")) Then
                   !BeginBal = !EndingBal
                   !Debit = Dramt
                   Else
                   !Debit = !Debit + Dramt
                  End If
                  !EndingBal = !BeginBal + Dramt
                  !TotalDebit = !TotalDebit + Dramt
                  !LastTransdate = Format(Date, "mm/dd/yyyy")
                  !LastTransType = "Debit"
                  !TotalTrans = !TotalTrans + 1
                  .Update
                  .close
                End If
               If Val(CrAmt) <> 0 Then
                 If Trim(!LastTransdate) <> Trim(Format(Date, "dd/mm/yyyy")) Then
                  !BeginBal = !EndingBal
                   !Credit = CrAmt
                   Else
                   !Credit = !Credit + CrAmt
                  End If
                  !EndingBal = !BeginBal + CrAmt
                  !TotalCredit = !TotalCredit + CrAmt
                  !LastTransdate = Format(Date, "mm/dd/yyyy")
                  !LastTransType = "Credit"
                  !TotalTrans = !TotalTrans + 1
                  .Update
                  .close
                End If
             End If
             End With

              Dim recupdatevou As New ADODB.Recordset ' this is for update the voucher table
              Dim con2 As New ADODB.Connection
              con2.Open "Dsn=Finance;Uid=Sa;Pwd=;"
              recupdatevou.Open "Select * from vouchers where post = 'no' and journalnumber = " & "'" & Trim(rsJRN!SerialNo) & "'", con2, adOpenKeyset, adLockOptimistic
                If recupdatevou.BOF = False Then
                    While recupdatevou.EOF = False
                        recupdatevou!Post = "Yes"
                        recupdatevou.Update
                        recupdatevou.MoveNext
                    Wend
                End If
                recupdatevou.close
                con2.close      'end voucher
            With rstGLMaster
                .addnew
                !JOurnalNo = rsJRN!SerialNo
                !AccountCode = rsJRN!accountnumber
                !accountname = rsJRN!accountname
                !PostDate = Date & " " & Time
                !recorddate = rsJRN!TRansDate
                !Particulars = rsJRN!Particulars
                !DebitAmount = Dramt
                !creditamount = CrAmt
                If Val(Dramt) <> 0 Then
                    !Balance = cBeginBal + Dramt
                 Else
                    !Balance = cBeginBal - CrAmt
                End If
                cBeginBal = !Balance
                .Update
             End With
            Dramt = 0
            CrAmt = 0
            rsJRN!remarks = "Posted Last " & Date & " " & Time
            rsJRN!Status = "Posted"
            rsJRN.Update
           rsJRN.MoveNext
           DoEvents
        Loop
        rsJRN.close
        POstNextTrans
End Sub
Sub PostCSP(cTotrec As Long)
        Dim rstGLMaster As New ADODB.Recordset
        Dim FinanceMAster As New ADODB.Recordset
        rstGLMaster.Open "GLMaster", constring, adOpenKeyset, adLockPessimistic, adCmdTable
        If rstGLMaster.EOF = True Then
            cBeginBal = 0
          Else
          cBeginBal = rstGLMaster!Balance
        End If
        i = 0
        Me.Label1.caption = "Status: Posting " & Trim(Me.ListView1.ListItems.Item(DayNo)) & " Transactions..."
        Do Until rsJRN.EOF = True
              i = i + 1
              If i = cTotrec Then
                cVal = cVal + 1
                i = 0
                If cVal <= 100 Then
                    Me.ProgressBar1.Value = cVal
                    Me.ListView1.ListItems.Item(DayNo).SubItems(4) = cVal & "% Process"
                End If
              End If
              Ac = rsJRN!accountnumber
              Dramt = rsJRN!DebitAmount
              CrAmt = rsJRN!creditamount

              'FinanceMAster.Open "FInanceMASter", conString, adOpenKeyset, adLockPessimistic, adCmdTable
              On Error Resume Next
              FinanceMAster.close
              On Error GoTo 0
              FinanceMAster.Open "SElect * From FInanceMASter where AccountCode = " & "'" & Ac & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
              MsgBox FinanceMAster.RecordCount
              With FinanceMAster
              If FinanceMAster.BOF = False Then
               If Val(Dramt) <> 0 Then
                 If Trim(!LastTransdate) <> Trim(Format(Date, "dd/mm/yyyy")) Then
                   !BeginBal = !EndingBal
                   !Debit = Dramt
                   Else
                   !Debit = !Debit + Dramt
                  End If
                  !EndingBal = !BeginBal + Dramt
                  !TotalDebit = !TotalDebit + Dramt
                  !LastTransdate = Format(Date, "mm/dd/yyyy")
                  !LastTransType = "Debit"
                  !TotalTrans = !TotalTrans + 1
                  .Update
                  .close
                End If
               If Val(CrAmt) <> 0 Then
                 If Trim(!LastTransdate) <> Trim(Format(Date, "dd/mm/yyyy")) Then
                  !BeginBal = !EndingBal
                   !Credit = CrAmt
                   Else
                   !Credit = !Credit + CrAmt
                  End If
                  !EndingBal = !BeginBal + CrAmt
                  !TotalCredit = !TotalCredit + CrAmt
                  !LastTransdate = Format(Date, "mm/dd/yyyy")
                  !LastTransType = "Credit"
                  !TotalTrans = !TotalTrans + 1
                  .Update
                  .close
                End If
              End If
              End With
            Dim recupdatevou As New ADODB.Recordset ' this is for update the voucher table
              Dim con2 As New ADODB.Connection
              con2.Open "Dsn=Finance;Uid=Sa;Pwd=;"
              recupdatevou.Open "Select * from vouchers where post = 'no' and journalnumber = " & "'" & Trim(rsJRN!SerialNo) & "'", con2, adOpenKeyset, adLockOptimistic
                If recupdatevou.BOF = False Then
                    While recupdatevou.EOF = False
                        recupdatevou!Post = "Yes"
                        recupdatevou.Update
                        recupdatevou.MoveNext
                    Wend
                End If
                recupdatevou.close
                con2.close      'end voucher


            With rstGLMaster
                .addnew
                !JOurnalNo = rsJRN!SerialNo
                !AccountCode = rsJRN!accountnumber
                !accountname = rsJRN!accountname
               !PostDate = Date & " " & Time
                !recorddate = rsJRN!TRansDate
                 !Particulars = rsJRN!Particulars
                !DebitAmount = Dramt
                !creditamount = CrAmt
                If Val(Dramt) <> 0 Then
                    !Balance = cBeginBal + Dramt
                 Else
                    !Balance = cBeginBal - CrAmt
                End If
                cBeginBal = !Balance
                .Update
             End With
            Dramt = 0
            CrAmt = 0
            rsJRN!remarks = "Posted Last " & Date & " " & Time
            rsJRN!Status = "Posted"
            rsJRN.Update
           rsJRN.MoveNext
           DoEvents
        Loop
        rsJRN.close
        POstNextTrans
End Sub
Sub PostPyb(cTotrec As Long)
        Dim rstGLMaster As New ADODB.Recordset
        Dim FinanceMAster As New ADODB.Recordset
        Dim rsPS As New ADODB.Recordset
        rstGLMaster.Open "GLMaster", constring, adOpenKeyset, adLockPessimistic, adCmdTable
        If rstGLMaster.EOF = False Then
          rstGLMaster.MoveLast
        End If
        If rstGLMaster.EOF = True Then
            cBeginBal = 0
          Else
          cBeginBal = rstGLMaster!Balance
        End If
        i = 0
        Me.Label1.caption = "Status: Posting " & Trim(Me.ListView1.ListItems.Item(DayNo)) & " Transactions..."
        Do Until rsJRN.EOF = True
              i = i + 1
              REf = rsJRN!serno
              If i = cTotrec Then
                cVal = cVal + 1
                i = 0
                If cVal <= 100 Then
                    Me.ProgressBar1.Value = cVal
                    Me.ListView1.ListItems.Item(DayNo).SubItems(4) = cVal & "% Process"
                End If
              End If
              Ac = rsJRN!AccNo
              Dramt = rsJRN!DBamount
              CrAmt = rsJRN!CRamount

              On Error Resume Next
              FinanceMAster.close
              FinanceMAster.Open "SElect * From FInanceMASter where AccountCode = " & "'" & Ac & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
              On Error GoTo 0
              With FinanceMAster
               If Val(Dramt) <> 0 Then
                 If Trim(!LastTransdate) <> Trim(Format(Date, "dd/mm/yyyy")) Then
                   !BeginBal = !EndingBal
                   !Debit = Dramt
                   Else
                   !Debit = !Debit + Dramt
                  End If
                  !EndingBal = !BeginBal + Dramt
                  !TotalDebit = !TotalDebit + Dramt
                  !LastTransdate = Format(Date, "mm/dd/yyyy")
                  !LastTransType = "Debit"
                  !TotalTrans = !TotalTrans + 1
                  .Update
                  .close
                End If
               If Val(CrAmt) <> 0 Then
                 If Trim(!LastTransdate) <> Trim(Format(Date, "dd/mm/yyyy")) Then
                  !BeginBal = !EndingBal
                   !Credit = CrAmt
                   Else
                   !Credit = !Credit + CrAmt
                  End If
                  !EndingBal = !BeginBal + CrAmt
                  !TotalCredit = !TotalCredit + CrAmt
                  !LastTransdate = Format(Date, "mm/dd/yyyy")
                  !LastTransType = "Credit"
                  !TotalTrans = !TotalTrans + 1
                  .Update
                  .close
                End If
              End With


            With rstGLMaster
                .addnew
                !JOurnalNo = rsJRN!SerialNo
                !AccountCode = rsJRN!AccNo
                !accountname = rsJRN!AccName
                !PostDate = Date & " " & Time
                !recorddate = rsJRN!confirmeddate
                !Particulars = rsJRN!Description 'rsJRN!particulars  ' "Voucher #" & rsJRN!serno & "/" & rsJRN!InvNo & "-" & rsJRN!InvDate & "/" & rsJRN!SENumber & "-" & rsJRN!SEDate & " Posted by:" & cLogUser
                !DebitAmount = Dramt
                !creditamount = CrAmt
                !category = rsJRN!Classification
                If Val(Dramt) <> 0 Then
                    !Balance = cBeginBal + Dramt
                 Else
                    !Balance = cBeginBal - CrAmt
                End If
                cBeginBal = !Balance
                .Update
             End With
            Dramt = 0
            CrAmt = 0
            rsJRN!remarks = "Posted Last " & Date & " " & Time
            rsJRN!Status = "Posted"
            rsJRN.Update
           rsJRN.MoveNext
           DoEvents
        Loop
        rsJRN.close


        Dim rstJourns As New ADODB.Recordset
        Dim NathXPay As New ADODB.Recordset
        Dim NAthXrec As New ADODB.Recordset
        Dim NathPays As New ADODB.Recordset
        Dim rstjrn2 As New ADODB.Recordset

        xPays = "SELECT * From Xpayment  Where serialno =" & "'" & REf & "'" & " and confirmedmark = 'uc' order by serialno"
        xrecs = "SELECT * From Xreceipt  Where serialno =" & "'" & REf & "'" & " and confirmedmark = 'uc' order by serialno"
        PeeDo = "SELECT * From Payablesetup  Where serialno =" & "'" & REf & "'" & "  order by serialno"

        NathXPay.Open xPays, constring, adOpenDynamic, adLockOptimistic, adCmdText
        NAthXrec.Open xrecs, constring, adOpenDynamic, adLockOptimistic, adCmdText
        NathPays.Open PeeDo, constring, adOpenDynamic, adLockOptimistic, adCmdText

        If NAthXrec.EOF = False Then
            NAthXrec.MoveFirst
        End If
        While NAthXrec.EOF = False
         If NAthXrec!SerialNo = REf Then
             NAthXrec!Postmark = "Yes"
         End If
         NAthXrec.MoveNext
        Wend


        If NathXPay.EOF = False Then
             NathXPay.MoveFirst
        End If
        While NathXPay.EOF = False
            If NathXPay!SerialNo = REf Then
                NathXPay!Postmark = "Yes"
            End If
        NathXPay.MoveNext
        Wend

       POstNextTrans

End Sub
Sub PostSRL(cTotrec As Long)
        Dim rstGLMaster As New ADODB.Recordset
        Dim FinanceMAster As New ADODB.Recordset
        rstGLMaster.Open "GLMaster", constring, adOpenKeyset, adLockPessimistic, adCmdTable
        If rstGLMaster.EOF = True Then
            cBeginBal = 0
          Else
          cBeginBal = rstGLMaster!Balance
        End If
        i = 0
        Me.Label1.caption = "Status: Posting " & Trim(Me.ListView1.ListItems.Item(DayNo)) & " Transactions..."
        Do Until rsJRN.EOF = True
              i = i + 1
              If i = cTotrec Then
                cVal = cVal + 1
                i = 0
                If cVal <= 100 Then
                    Me.ProgressBar1.Value = cVal
                    Me.ListView1.ListItems.Item(DayNo).SubItems(4) = cVal & "% Process"
                End If
              End If
              Ac = rsJRN!accountnumber
              Dramt = rsJRN!DebitAmount
              CrAmt = rsJRN!creditamount

              'FinanceMAster.Open "FInanceMASter", conString, adOpenKeyset, adLockPessimistic, adCmdTable
              'FinanceMAster.Close
              FinanceMAster.Open "SElect * From FInanceMASter where AccountCode = " & "'" & Ac & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
              With FinanceMAster
               If Val(Dramt) <> 0 Then
                 If Trim(!LastTransdate) <> Trim(Format(Date, "dd/mm/yyyy")) Then
                   !BeginBal = !EndingBal
                   !Debit = Dramt
                   Else
                   !Debit = !Debit + Dramt
                  End If
                  !EndingBal = !BeginBal + Dramt
                  !TotalDebit = !TotalDebit + Dramt
                  !LastTransdate = Format(Date, "mm/dd/yyyy")
                  !LastTransType = "Debit"
                  !TotalTrans = !TotalTrans + 1
                  .Update
                  .close
                End If
               If Val(CrAmt) <> 0 Then
                 If Trim(!LastTransdate) <> Trim(Format(Date, "dd/mm/yyyy")) Then
                  !BeginBal = !EndingBal
                   !Credit = CrAmt
                   Else
                   !Credit = !Credit + CrAmt
                  End If
                  !EndingBal = !BeginBal + CrAmt
                  !TotalCredit = !TotalCredit + CrAmt
                  !LastTransdate = Format(Date, "mm/dd/yyyy")
                  !LastTransType = "Credit"
                  !TotalTrans = !TotalTrans + 1
                  .Update
                  .close
                End If
              End With
            With rstGLMaster
                .addnew
                !JOurnalNo = rsJRN!SerialNo
                !AccountCode = rsJRN!accountnumber
                !accountname = rsJRN!accountname
                !PostDate = Date & " " & Time
                !recorddate = rsJRN!TRansDate
                !Particulars = rsJRN!Description & "/ Posted by : " & cLogUser
                !DebitAmount = Dramt
                !creditamount = CrAmt
                If Val(Dramt) <> 0 Then
                    !Balance = cBeginBal + Dramt
                 Else
                    !Balance = cBeginBal - CrAmt
                End If
                cBeginBal = !Balance
                .Update
             End With
            Dramt = 0
            CrAmt = 0
            rsJRN!remarks = "Posted Last " & Date & " " & Time
            rsJRN!Status = "Posted"
            rsJRN.Update
           rsJRN.MoveNext
           DoEvents
        Loop
        rsJRN.close
        POstNextTrans
End Sub
Sub PostSPA(cTotrec As Long)
        Dim rstGLMaster As New ADODB.Recordset
        Dim FinanceMAster As New ADODB.Recordset
        rstGLMaster.Open "GLMaster", constring, adOpenKeyset, adLockPessimistic, adCmdTable
        If rstGLMaster.EOF = True Then
            cBeginBal = 0
          Else
          cBeginBal = rstGLMaster!Balance
        End If
        i = 0
        Me.Label1.caption = "Status: Posting " & Trim(Me.ListView1.ListItems.Item(DayNo)) & " Transactions..."
        Do Until rsJRN.EOF = True
              i = i + 1
              If i = cTotrec Then
                cVal = cVal + 1
                i = 0
                If cVal <= 100 Then
                    Me.ProgressBar1.Value = cVal
                    Me.ListView1.ListItems.Item(DayNo).SubItems(4) = cVal & "% Process"
                End If
              End If
              Ac = rsJRN!accountnumber
              Dramt = rsJRN!DebitAmount
              CrAmt = rsJRN!creditamount

              'FinanceMAster.Open "FInanceMASter", conString, adOpenKeyset, adLockPessimistic, adCmdTable
              'FinanceMAster.Close
              FinanceMAster.Open "SElect * From FInanceMASter where AccountCode = " & "'" & Ac & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
              With FinanceMAster
               If Val(Dramt) <> 0 Then
                 If Trim(!LastTransdate) <> Trim(Format(Date, "dd/mm/yyyy")) Then
                   !BeginBal = !EndingBal
                   !Debit = Dramt
                   Else
                   !Debit = !Debit + Dramt
                  End If
                  !EndingBal = !BeginBal + Dramt
                  !TotalDebit = !TotalDebit + Dramt
                  !LastTransdate = Format(Date, "mm/dd/yyyy")
                  !LastTransType = "Debit"
                  !TotalTrans = !TotalTrans + 1
                  .Update
                  .close
                End If
               If Val(CrAmt) <> 0 Then
                 If Trim(!LastTransdate) <> Trim(Format(Date, "dd/mm/yyyy")) Then
                  !BeginBal = !EndingBal
                   !Credit = CrAmt
                   Else
                   !Credit = !Credit + CrAmt
                  End If
                  !EndingBal = !BeginBal + CrAmt
                  !TotalCredit = !TotalCredit + CrAmt
                  !LastTransdate = Format(Date, "mm/dd/yyyy")
                  !LastTransType = "Credit"
                  !TotalTrans = !TotalTrans + 1
                  .Update
                  .close
                End If
              End With
            With rstGLMaster
                .addnew
                !JOurnalNo = rsJRN!SerialNo
                !AccountCode = rsJRN!accountnumber
                !accountname = rsJRN!accountname
                !PostDate = Date & " " & Time
                !recorddate = rsJRN!TRansDate
                !Particulars = rsJRN!Description & "/ Posted by : " & cLogUser
                !DebitAmount = Dramt
                !creditamount = CrAmt
                If Val(Dramt) <> 0 Then
                    !Balance = cBeginBal + Dramt
                 Else
                    !Balance = cBeginBal - CrAmt
                End If
                cBeginBal = !Balance
                .Update
             End With
            Dramt = 0
            CrAmt = 0
            rsJRN!remarks = "Posted Last " & Date & " " & Time
            rsJRN!Status = "Posted"
            rsJRN.Update
           rsJRN.MoveNext
           DoEvents
        Loop
        rsJRN.close
        POstNextTrans
End Sub
Sub PostPetty(cTotrec As Long)
'        On Error GoTo NElson
        Dim rstGLMaster As New ADODB.Recordset
        Dim FinanceMAster As New ADODB.Recordset
        rstGLMaster.Open "GLMaster", constring, adOpenKeyset, adLockPessimistic, adCmdTable
        If rstGLMaster.EOF = False Then
          rstGLMaster.MoveLast
        End If
        If rstGLMaster.EOF = True Then
            cBeginBal = 0
          Else
          cBeginBal = rstGLMaster!Balance
        End If
        i = 0
        Me.Label1.caption = "Status: Posting " & Trim(Me.ListView1.ListItems.Item(DayNo)) & " Transactions..."
        Do Until rsJRN.EOF = True
              i = i + 1
              REf = rsJRN!Journo
              If i = cTotrec Then
                cVal = cVal + 1
                i = 0
                If cVal <= 100 Then
                    Me.ProgressBar1.Value = cVal
                    Me.ListView1.ListItems.Item(DayNo).SubItems(4) = cVal & "% Process"
                End If
              End If
              Ac = rsJRN!AccoutNo
              Dramt = rsJRN!DBamount
              CrAmt = rsJRN!CRamount

              On Error Resume Next
              FinanceMAster.close
              FinanceMAster.Open "SElect * From FinanceMAster where AccountCode = " & "'" & Ac & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
              On Error GoTo 0
              With FinanceMAster
               If Val(Dramt) <> 0 Then
                 If Trim(!LastTransdate) <> Trim(Format(Date, "dd/mm/yyyy")) Then
                   !BeginBal = !EndingBal
                   !Debit = Dramt
                   Else
                   !Debit = !Debit + Dramt
                  End If
                  !EndingBal = !BeginBal + Dramt
                  !TotalDebit = !TotalDebit + Dramt
                  !LastTransdate = Format(Date, "mm/dd/yyyy")
                  !LastTransType = "Debit"
                  !TotalTrans = !TotalTrans + 1
                  .Update
                  .close
                End If
               If Val(CrAmt) <> 0 Then
                 If Trim(!LastTransdate) <> Trim(Format(Date, "dd/mm/yyyy")) Then
                  !BeginBal = !EndingBal
                   !Credit = CrAmt
                   Else
                   !Credit = !Credit + CrAmt
                  End If
                  !EndingBal = !BeginBal + CrAmt
                  !TotalCredit = !TotalCredit + CrAmt
                  !LastTransdate = Format(Date, "mm/dd/yyyy")
                  !LastTransType = "Credit"
                  !TotalTrans = !TotalTrans + 1
                  .Update
                  .close
                End If
              End With


            With rstGLMaster
                .addnew
                !JOurnalNo = rsJRN!Journo
                !AccountCode = rsJRN!AccoutNo
                !accountname = rsJRN!accountname
                !PostDate = Date & " " & Time
                !recorddate = rsJRN!confirmeddate
                !Particulars = "Voucher #" & rsJRN!SerialNo & "/" & rsJRN!Expla & "/ Posted by : " & cLogUser
                !DebitAmount = rsJRN!DBamount
                !creditamount = rsJRN!CRamount
                If Val(Dramt) <> 0 Then
                    !Balance = cBeginBal + rsJRN!DBamount
                 Else
                    !Balance = cBeginBal - rsJRN!CRamount
                End If
                cBeginBal = !Balance
                .Update

             End With
            Dramt = 0
            CrAmt = 0
            rsJRN!Postmark = "Posted"
            rsJRN.Update
           rsJRN.MoveNext
           DoEvents
        Loop


        rsJRN.close

       POstNextTrans

End Sub

Sub PostBNK(cTotrec As Long)
        Dim rstGLMaster As New ADODB.Recordset
        Dim FinanceMAster As New ADODB.Recordset
        rstGLMaster.Open "GLMaster", constring, adOpenKeyset, adLockPessimistic, adCmdTable
        If rstGLMaster.EOF = False Then
          rstGLMaster.MoveLast
        End If
        If rstGLMaster.EOF = True Then
            cBeginBal = 0
          Else
          cBeginBal = rstGLMaster!Balance
        End If
        i = 0
        Me.Label1.caption = "Status: Posting " & Trim(Me.ListView1.ListItems.Item(DayNo)) & " Transactions..."
        Do Until rsJRN.EOF = True
              i = i + 1
              If i = cTotrec Then
                cVal = cVal + 1
                i = 0
                If cVal <= 100 Then
                    Me.ProgressBar1.Value = cVal
                    Me.ListView1.ListItems.Item(DayNo).SubItems(4) = cVal & "% Process"
                End If
              End If
              Ac = rsJRN!accountnumber
              Dramt = rsJRN!DebitAmount
              CrAmt = rsJRN!creditamount

              'FinanceMAster.Open "FInanceMASter", conString, adOpenKeyset, adLockPessimistic, adCmdTable
              'FinanceMAster.Close
              FinanceMAster.Open "SElect * From FInanceMASter where AccountCode = " & "'" & Ac & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
              With FinanceMAster
               If Val(Dramt) <> 0 Then
                 If Trim(!LastTransdate) <> Trim(Format(Date, "dd/mm/yyyy")) Then
                   !BeginBal = !EndingBal
                   !Debit = Dramt
                   Else
                   !Debit = !Debit + Dramt
                  End If
                  !EndingBal = !BeginBal + Dramt
                  !TotalDebit = !TotalDebit + Dramt
                  !LastTransdate = Format(Date, "mm/dd/yyyy")
                  !LastTransType = "Debit"
                  !TotalTrans = !TotalTrans + 1
                  .Update
                  .close
                End If
               If Val(CrAmt) <> 0 Then
                 If Trim(!LastTransdate) <> Trim(Format(Date, "dd/mm/yyyy")) Then
                  !BeginBal = !EndingBal
                   !Credit = CrAmt
                   Else
                   !Credit = !Credit + CrAmt
                  End If
                  !EndingBal = !BeginBal + CrAmt
                  !TotalCredit = !TotalCredit + CrAmt
                  !LastTransdate = Format(Date, "mm/dd/yyyy")
                  !LastTransType = "Credit"
                  !TotalTrans = !TotalTrans + 1
                  .Update
                  .close
                End If
              End With


            With rstGLMaster
                .addnew
                !JOurnalNo = rsJRN!SerialNo
                !AccountCode = rsJRN!accountnumber
                !accountname = rsJRN!accountname
                !PostDate = Date & " " & Time
                !recorddate = rsJRN!TRansDate
                !Particulars = rsJRN!Description & "/ Posted by : " & cLogUser
                !DebitAmount = Dramt
                !creditamount = CrAmt
                If Val(Dramt) <> 0 Then
                    !Balance = cBeginBal + Dramt
                 Else
                    !Balance = cBeginBal - CrAmt
                End If
                cBeginBal = !Balance
                .Update
             End With
            Dramt = 0
            CrAmt = 0
            rsJRN!remarks = "Posted Last " & Date & " " & Time
            rsJRN!Status = "Posted"
            rsJRN.Update
           rsJRN.MoveNext
           DoEvents
        Loop
        rsJRN.close
        POstNextTrans



End Sub


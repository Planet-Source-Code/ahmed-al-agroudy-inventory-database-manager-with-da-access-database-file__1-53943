VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form SalesJournal 
   Caption         =   "Sales Journal"
   ClientHeight    =   3045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4425
   LinkTopic       =   "Form1"
   ScaleHeight     =   3045
   ScaleWidth      =   4425
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView ListView2 
      Height          =   2175
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   3836
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2520
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SalesJournal.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SalesJournal.frx":015A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SalesJournal.frx":02B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SalesJournal.frx":0706
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   1376
      ButtonWidth     =   1746
      ButtonHeight    =   1376
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Print ØÈÇÚÉ"
            Key             =   "Print"
            Object.ToolTipText     =   "Printing"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "&Post áÕÞ"
            Key             =   "Post"
            Object.ToolTipText     =   "Post to General Ledger"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Scroll ÍÑßÉ"
            Key             =   "Scroll"
            Object.ToolTipText     =   "Scroll through end"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Close ÛáÞ"
            Key             =   "Close"
            Object.ToolTipText     =   "&Close and go back"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
      EndProperty
      MousePointer    =   99
      MouseIcon       =   "SalesJournal.frx":0C48
   End
End
Attribute VB_Name = "SalesJournal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MItem As ListItem

Private Sub Command2_Click()
Me.ListView2.SetFocus
Dim rstSj As New ADODB.Recordset
Dim rstGLMaster As New ADODB.Recordset
Dim FinanceMAster As New ADODB.Recordset

Dim rstJOurcode As New ADODB.Recordset
rstJOurcode.Open "Select * from JOurnalCode order by Code", constring, adOpenKeyset, adLockPessimistic, adCmdText
rstJOurcode.Move 10, 1
xCode = rstJOurcode!JOurnalName
Dim trandate As Date
trandate = rstJOurcode!lastpostingdate
rstJOurcode.close

rstSj.Open "Select * from SalesJOurnal where Transdate> " & "'" & trandate & "'", constring, _
            adOpenKeyset, adLockPessimistic, adCmdText
rstGLMaster.Open "GLMaster", constring, adOpenKeyset, adLockPessimistic, adCmdTable
rstGLMaster.MoveLast

If rstGLMaster.EOF = True Then
    cBeginBal = 0
  Else
  cBeginBal = rstGLMaster!Balance
End If
i = 0
Do Until rstSj.EOF = True
      i = i + 1
      Ac = rstSj!accountnumber
      Dramt = IIf(IsNull(rstSj!DebitAmount) = True, 0, rstSj!DebitAmount)
      CrAmt = IIf(IsNull(rstSj!creditamount) = True, 0, rstSj!creditamount)
      Dim strFindMe As String
      Dim itmFound As ListItem   ' FoundItem variable.
      intSelectedOption = lvwText
      strFindMe = Trim(i)
      Set itmFound = Me.ListView2.Finditem(strFindMe, intSelectedOption, , lvwPartial)
      If itmFound Is Nothing Then  ' If no match, inform user and exit.
       Else
       itmFound.EnsureVisible
       itmFound.Selected = True   ' Select the ListItem.
       Me.ListView2.SelectedItem.Ghosted = True
       Me.ListView2.SetFocus
      End If

        
      
      If Trim(Ac) <> "" Then
        FinanceMAster.Open "SElect * From FInanceMASter where AccountCode = " & "'" & Ac & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
       Else
       Ac = "111020101000" 'for blank Acctno will post to Miscellaneous Acct
       FinanceMAster.Open "SElect * From FInanceMASter where AccountCode = " & "'" & Ac & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
      End If
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
          .Update
          .close
        Else
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
          .Update
          .close
        End If
      End With
      
     
    With rstGLMaster
        .addnew
        !JOurnalNo = rstSj!SerialNo
        !AccountCode = rstSj!accountnumber
        !accountname = rstSj!accountname
        !PostDate = rstSj!TRansDate
        !recorddate = rstSj!TRansDate
        !particulars = rstSj!Description
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
    
   rstSj.MoveNext
  DoEvents
Loop
MsgBox ("Finished")
End Sub

Private Sub Command3_Click()
Me.ListView2.SetFocus
i = Me.ListView2.SelectedItem.Index - 1
If i + 1 = Me.ListView2.ListItems.Count Then
        l = 0
        X = i
    For l = 1 To i - 1 Step 1
        X = X - 1
        MItem.EnsureVisible
        MItem.Selected = True
       
    Next
    Else
    For i = 1 To Me.ListView2.ListItems.Count Step 1
        On Error Resume Next
        MItem.EnsureVisible
        MItem.Selected = True
       
    Next
End If

End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
'put it the listview before loading the SalesJournal Window
Dim rstSj As New ADODB.Recordset
Dim TRansDate As Date
Dim trDate As Date
TRansDate = Trim(Mainform.lvListView.SelectedItem.Text) ' , "mm/dd/yyyy")
trDate = Mid(TRansDate, 4, 2) & "/" & Left(TRansDate, 2) & "/" & Mid(TRansDate, 7, 4)
If Mainform.lvListView.SelectedItem.SubItems(4) = "Unpost" Then
  rstSj.Open "Select * from SalesJournal where remarks is null and transdate= " & "'" & trDate & "'" & "order by ticket", constring, adOpenKeyset, adLockPessimistic, adCmdText
 Else
 rstSj.Open "Select * from SalesJournal where remarks is not null and transdate= " & "'" & trDate & "'" & "order by ticket", constring, adOpenKeyset, adLockPessimistic, adCmdText
End If
Me.ListView2.ListItems.clear
With rstSj
  Do Until .EOF = True
          Set MItem = SalesJournal.ListView2.ListItems.Add(, , rstSj!ticket)
          MItem.SubItems(1) = rstSj!ticket
          MItem.SubItems(2) = IIf(IsNull(rstSj!accountnumber) = True, "", rstSj!accountnumber)
          MItem.SubItems(3) = IIf(IsNull(rstSj!accountname) = True, "", rstSj!accountname)
          MItem.SubItems(4) = TRansDate
          MItem.SubItems(5) = Format(rstSj!DebitAmount, "###,###,###.#0")
          MItem.SubItems(6) = Format(rstSj!creditamount, "###,###,###.#0")
          MItem.SubItems(7) = rstSj!SerialNo
          MItem.SubItems(8) = rstSj!Description
          rstSj.MoveNext
         DoEvents
         SendKeys "{Pgdn}"
    Loop
End With
If Me.ListView2.ListItems.Count = 0 Then
    Me.Toolbar1.Buttons(1).Enabled = False
    Me.Toolbar1.Buttons(3).Enabled = False
    Me.Toolbar1.Buttons(5).Enabled = False
End If

End Sub

Private Sub Form_Resize()
Me.ListView2.Width = Me.Width - 120
Me.ListView2.Height = Me.Height - 1140

End Sub

Private Sub ListView2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Me.ListView2.SortKey = ColumnHeader.Index - 1
Me.ListView2.Sorted = True
End Sub

Private Sub ListView2_DblClick()
Dim rstSalesInfo As New ADODB.Recordset
On Error Resume Next
AcctCode = Trim(Me.ListView2.SelectedItem.SubItems(8))
AcctCode = Trim(Right(AcctCode, 10))
InvDate = Format(Me.ListView2.SelectedItem.SubItems(4), "mm/dd/yyyy")
rstSalesInfo.Open "Select * from salesjournal where InvoiceNo=" & "'" & AcctCode & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
SalesInfo.Combo1 = Format(rstSalesInfo!invoicedate, "dd/mm/yyyy")
SalesInfo.Combo2 = rstSalesInfo!invoiceno
SalesInfo.Combo3 = rstSalesInfo!ClientCode
SalesInfo.Combo4 = rstSalesInfo!profitcenter 'IIf(IsNull(rstSalesInfo!ProfitCenter) = True, 0, rstSalesInfo!Prf_No)
SalesInfo.Combo5 = Format(rstSalesInfo!tradercvble, "###,###,###.##0")

SalesInfo.Combo7 = Format(rstSalesInfo!tradeDiscamt, "###,###,###.##0") & " @ " & rstSalesInfo!TRadedisc & "%"
SalesInfo.Combo8 = Format(rstSalesInfo!GrossSales, "###,###,###.##0")
SalesInfo.Combo9 = Format(rstSalesInfo!transpoCharge)
SalesInfo.Combo10 = Format(rstSalesInfo!NetSales, "###,###,###.##0")
SalesInfo.Combo11 = Format(rstSalesInfo!vat, "###,###,###.##0")
SalesInfo.Combo12 = Format(rstSalesInfo!SurTaxAmt, "###,###,###.##0")

SalesInfo.Show 1


End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
        Case "Print"
'            On Error GoTo Nelson
'            Me.ListView2.SetFocus
'            Dim rstJourCode As New ADODB.Recordset
'            Dim rstSj As New ADODB.Recordset
'
'            rstJourCode.Open "Select * from JOurnalCode order by Code", conString, adOpenKeyset, adLockPessimistic, adCmdText
'            rstJourCode.Move 10, 1
'            xCode = rstJourCode!JOurnalName
'            Dim trandate As Date
'            trandate = rstJourCode!LastPostingDate
'            rstJourCode.Close
'
'
'
'            FinanceDE.SalesJournal trandate
'            SalesJournalRep.Show 1
             Dim SelectedDate As Date
             SelectedDate = Format(Date, "mm/dd/yyyy")
             Call PrintSalesJOurnal(SelectedDate)
Nelson:
            c = Err.Number
            If c = 3705 Then
              FinanceDE.rsSalesJournal.close
              FinanceDE.SalesJournal trandate
              SalesJournalRep.Show 1
            End If
            
     Case "Post"
          
        Dim rsInvj As New ADODB.Recordset
        Dim MItem As ListItem
        mess = MsgBox("Do you want to continue? ", vbQuestion + vbYesNo + vbDefaultButton2, "Please confirm")
        If mess = vbYes Then
             PostingJournal.Text1.Text = "SAL"
             rsInvj.Open "SELECT  TransDate, COUNT(TransDate) AS TotalTRan, SUM(DebitAmount) AS DrAmt, SUM(CreditAmount) AS CrAmt" _
             & " From SalesJournal where Remarks is null GROUP BY TransDate ORDER BY TransDate", constring, adOpenKeyset, adLockPessimistic, adCmdText
            'rsInvj.Close
            Do Until rsInvj.EOF = True
                Set MItem = PostingJournal.ListView1.ListItems.Add(, , Format(rsInvj!TRansDate, "dd/mm/yyyy"))
                MItem.SubItems(1) = rsInvj!TotalTRan
                MItem.SubItems(2) = FormatNumber(rsInvj!Dramt, 2, vbTrue, vbTrue, vbTrue)
                MItem.SubItems(3) = FormatNumber(rsInvj!CrAmt, 2, vbTrue, vbTrue, vbTrue)
                MItem.SubItems(4) = "Waiting"
                rsInvj.MoveNext
            Loop
            Unload Me
            PostingJournal.Show 1
        End If
        
     Case "Scroll"
        Me.ListView2.SetFocus
        i = Me.ListView2.SelectedItem.Index - 1
        If i + 1 = Me.ListView2.ListItems.Count Then
                l = 0
                X = i
            For l = 1 To i - 1 Step 1
                X = X - 1
                MItem.EnsureVisible
                MItem.Selected = True
               
            Next
            Else
            For i = 1 To Me.ListView2.ListItems.Count Step 1
                On Error Resume Next
                MItem.EnsureVisible
                MItem.Selected = True
               
            Next
        End If
         
    Case "Close"
     Unload Me
End Select
End Sub

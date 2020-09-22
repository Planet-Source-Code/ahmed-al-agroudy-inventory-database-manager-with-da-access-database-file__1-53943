VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form DownLoadinvoices 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Please confirm  ãä ÝÖáß ÇáÊÇßíÏ"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Height          =   350
      Left            =   120
      ScaleHeight     =   285
      ScaleWidth      =   4395
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   4455
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   330
         Left            =   -15
         TabIndex        =   2
         Top             =   -15
         Width           =   4440
         _ExtentX        =   7832
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel ÇáÛÇÁ"
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
      Left            =   2640
      TabIndex        =   4
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Continue ÇÓÊãÑÇÑ"
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
      Left            =   1080
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
   Begin MSComCtl2.Animation Animation1 
      Height          =   375
      Left            =   4440
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      _Version        =   393216
      AutoPlay        =   -1  'True
      BackColor       =   16777215
      FullWidth       =   33
      FullHeight      =   25
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Downloading Sales Invoices.... ÊÍãíá ÝæÇÊíÑ ÇáãÈíÚÇÊ"
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
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Label Label2 
      Caption         =   "Do you want to continue for Downloading invoices?"
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   320
      Width           =   3855
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "Downloadinvoices.frx":0000
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "DownLoadinvoices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsInv As New ADODB.Recordset
Dim rsInv2 As New ADODB.Recordset
Dim trandate As Date
Dim ErrorConek As Boolean
Dim i As Long
Dim reload As Boolean
Dim cTotrec As Long
Dim cVal As Long
Sub DownLoadINv()
Dim rstJOurcode As New ADODB.Recordset
Dim rstFM As New ADODB.Recordset
rstJOurcode.Open "Select * from JOurnalCode where Code='SAL'", constring, adOpenKeyset, adLockPessimistic, adCmdText
xCode = rstJOurcode!JOurnalName
Dim trandate As Date
Dim InvDate As Date
Dim cInvDate As Date
trandate = Format(rstJOurcode!lastpostingdate, "mm/dd/yyyy")
Dim cFound As Long
Static CancelDownLoading As Boolean
Me.caption = "Please wait...ÇáÑÌÇÁ ÇáÇäÊÙÇÑ"

Me.Command1.Left = 1800
Me.Picture1.Visible = True
Me.Label1.Visible = True
Me.Command2.Visible = False
Me.Label2.Visible = False
Me.Image1.Visible = False
If CancelDownLoading Then
   CancelDownLoading = False
  Else
    Me.Command1.caption = "Stop ÞÝ"
    CancelDownLoading = True
    
    Do Until rsInv2.EOF = True
        i = i + 1
        If i = cTotrec Then
            cVal = cVal + 1
            i = 0
            If cVal <= 100 Then
            Me.ProgressBar1.Value = cVal
            Me.Label3.caption = cVal & "%"
            End If
            
        End If
   
     If rsInv2!invc_date > trandate Then 'beyond the last posting date
      If Left(Trim(rsInv2!cust_code), 1) = "O" Then 'CustCode must be start w/ O'
       If Left(Trim(rsInv2!invc_no), 3) = "O 6" Then  'exclude for "OS",OJ" etc.
        If Left(Trim(rsInv2!invc_no), 6) <> "O 6029" Or Left(Trim(rsInv2!invc_no), 6) <> "O 6028" Or Left(Trim(rsInv2!invc_no), 6) <> "O 6039" Then
         If Left(rsInv2!factory, 1) = "Z" Or Left(rsInv2!factory, 1) = "W" Then 'ProfieCenter must start with "Z"
          If rsInv2!Cancelled <> True Then 'if invoice not canceled
            If rsInv2!Download <> True Then 'if it is not already download
             InvDate = Format(rsInv2!invc_date, "dd/mm/yyyy")
              cInvDate = Format(InvDate, "mm/dd/yyyy")
                  cFound = cFound + 1
                  With rsInv
                      .AddNew
                     
                      If Trim(rsInv2!mcustcode) = "" Then
                         rsInv!cust_code = rsInv2!cust_code
                       Else
                       rsInv!cust_code = rsInv2!mcustcode
                      End If
                      rsInv!invc_no = rsInv2!invc_no
                      rsInv!invc_date = rsInv2!invc_date ' Format(rsInv2!invc_date, "dd/mm/yyyy")
                      rsInv!Trans_dt = rsInv2!Trans_dt
                      rsInv!Tot_Qty = rsInv2!Tot_Qty
                      rsInv!Sub_Amt = rsInv2!Sub_Amt
                      rsInv!surcharge = rsInv2!surcharge
                      rsInv!SurTax = rsInv2!SurTax
                      rsInv!transchg = rsInv2!transchg
                      rsInv!Discount = rsInv2!Discount
                      rsInv!Tot_Vat = rsInv2!Tot_Vat
                      rsInv!Mngt_Acct = rsInv2!Mngt_Acct
                      rsInv!tot_amt = rsInv2!tot_amt
                      rsInv!amt_paid = rsInv2!amt_paid
                      rsInv!factory = rsInv2!factory
                      rsInv!acctNo = rsInv2!acctNo
                      rsInv!TDisAmount = rsInv2!TDisc
                      rstFM.Open "Select * from FinanceMaster where AccountCode=" & "'" & Trim(rsInv2!acctNo) & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
                      If rstFM.EOF = True Then
                        rsInv!Acctname = ""
                       Else
                        rsInv!Acctname = rstFM!accountnameeng
                      End If
                      rstFM.Close
                      .Update
                     
                     'for invcmain table
                     rsInv2!Download = True 'update download field in invoicemain to true
                     rsInv2.Update
                     
                  End With
                End If
               End If
              End If
             End If
            End If
           End If
          End If
             rsInv2.MoveNext
             DoEvents

            If CancelDownLoading = False Then
              Exit Do
              Exit Sub
            End If
       
    Loop
    If CancelDownLoading = False Then
        Me.Animation1.Visible = False
        Exit Sub
      Else
      Beep
       mess = MsgBox("Finished downloading ÇäÊåÇÁ ÇáÊÍãíá!  " & cFound & " Found(s)íÌÏ ", vbInformation + vbOKOnly, "Message")
       i = 0
       cVal = 0
       CancelDownLoading = False
       Me.Command2.Visible = False
       Me.Command1.Visible = False
       
       Me.Height = 1400
       If cFound = 0 Then
          Unload Me
          Exit Sub
       End If
       'Me.Label1.caption = "Now updating Client Account No..."
       'UpdateClientCode
       Me.Label1.caption = "Now Generating Sales Journal...ÇáÇä íæãíÉ ÇáÍÓÇÈÇÊ ÇáÚÇãÉ"
       GenerateSalesJournal
       
       Exit Sub
    End If
    
    Unload Me
End If
Command1.caption = "Continue ÇÓÊãÑÇÑ"
CancelDownLoading = False
Exit Sub
End Sub
Sub UpdateClientCode()
Dim RstInv As New ADODB.Recordset
Dim rstClient As New ADODB.Recordset
Dim rstJOurcode As New ADODB.Recordset
constring3 = "dsN=Clients;UID=SA;PWD=;"
rstClient.Open "Select * from Marcusfl order", constring3, adOpenKeyset, adLockPessimistic, adCmdText
RstInv.Open "Invoices", constring, adOpenKeyset, adLockPessimistic, adCmdText
cTotrec = Int(RstInv.RecordCount / 100)
Totrec = cTotrec
i = 0
cVal = 0

Do Until RstInv.EOF = True
        i = i + 1
        If i = cTotrec Then
            cVal = cVal + 1
            i = 0
            If cVal <= 100 Then
                Me.ProgressBar1.Value = cVal
                Me.Label3.caption = cVal & "%"
              End If
        End If
   ClientCode = Trim(RstInv!cust_code)
   rstClient.Open "SElect * from MarCusFL where Cust_code =" & "'" & ClientCode & "'", constring3, adOpenKeyset, adLockPessimistic, adCmdText
   If rstClient.EOF = False Then
    AcctCode = rstClient!acctNo
    Acctname = rstClient!first_name & " " & rstClient!last_name
    rstClient.Close
    RstInv!acctNo = AcctCode
    RstInv!Acctname = Acctname
    RstInv.Update
    'rstClient.Close
   Else
   rstClient.Close
   End If
   RstInv.MoveNext
   DoEvents
Loop
Me.Label1.caption = "Now Generating Sales Journal...ÇáÇä íæãíÉ ÇáÍÓÇÈÇÊ ÇáÚÇãÉ"
GenerateSalesJournal
End Sub
Sub GenerateSalesJournal()
Dim rstClient As New ADODB.Recordset
constring3 = "dsN=Clients;UID=SA;PWD=;"
rstClient.Open "Marcusfl", constring3, adOpenKeyset, adLockPessimistic, adCmdTable

Dim RstInv As New ADODB.Recordset
Dim rstSj As New ADODB.Recordset
Dim rstFM As New ADODB.Recordset
Dim rstjNo As New ADODB.Recordset
Dim rstJOurcode As New ADODB.Recordset
Dim CancelProcess As Boolean
Dim Dramt As Currency
Dim CrAmt As Currency
Dim IsInvDateEqual As Boolean
rstJOurcode.Open "Select * from JOurnalCode order by Code", constring, adOpenKeyset, adLockPessimistic, adCmdText
rstJOurcode.Move 10, 1
xCode = rstJOurcode!JOurnalName
Dim trandate As Date
trandate = rstJOurcode!lastpostingdate
IsInvDateEqual = True

'rstINV.Open "Select * from Invoices where Invc_date>" & "'" & trandate & "'", conString, adOpenKeyset, adLockPessimistic, adCmdText
RstInv.Open "SElect * from Invoices order by Invc_Date", constring, adOpenKeyset, adLockPessimistic, adCmdText
cTotrec = Int(RstInv.RecordCount / 100)
Totrec = IIf(cTotrec = 0, 1, cTotrec)
i = 0
cVal = 0

Dim sjtrandate  As Date
sjtrandate = Format(Date, "mm/dd/yyyy")

rstSj.Open "SalesJournal", constring, adOpenKeyset, adLockPessimistic, adCmdTable
Dim Jn As String
Call NewJn(Jn)

X = 0
i = 1
IDate = RstInv!invc_date 'FOR INVOICE DATE
vDate = RstInv!invc_date 'FOR INVOICE DATE
Do Until RstInv.EOF = True
    InvoiceClientCode = Trim(RstInv!cust_code)
    
    If Trim(RstInv!acctNo) = "" Then
        rstClient.MoveFirst
        Do Until rstClient.EOF = True
            If Trim(rstClient!cust_code) = InvoiceClientCode And Trim(rstClient!Grp) = "OS" Then
                RstInv!acctNo = "111041001001"
                RstInv!Acctname = "Obour Show Gen. Acct."
                RstInv.Update
                'rstClient.Close
                Exit Do
            End If
           rstClient.MoveNext
        Loop
        'rstClient.Close
     End If

    'if InvoiceDate not from the previous trans we assigned new JN
    If IDate <> vDate Then
        IDate = RstInv!invc_date 'FOR INVOICE DATE
        i = 1
       Call NewJn(Jn)
    End If
    IsInvDateEqual = True
        X = X + 1
        If X = Totrec Then
            cVal = cVal + 1
            X = 0
            If cVal <= 100 Then
            Me.ProgressBar1.Value = cVal
            Me.Label3.caption = cVal & "%"
            End If
        End If
    With rstSj
        'for dr
        On Error GoTo 0
        .AddNew
        !SerialNo = Jn
        !ticket = i
        !accountnumber = RstInv!acctNo
        !accountname = RstInv!Acctname
        !TRansDate = RstInv!invc_date 'Format(Date, "dd/mm/yyyy")
        !DebitAmount = RstInv!tot_amt + RstInv!TDisAmount     'rstINV!Sub_Amt  '* (rstINV!Mngt_Acct / 100))
        '!creditamount = (rstINV!Sub_Amt + IIf(IsNull(rstINV!Transchg) = True, 0, rstINV!Transchg) + rstINV!Tot_Vat) + IIf(IsNull(rstINV!Surcharge) = True, 0, rstINV!Surcharge)
       If RstInv!invc_no = "O 60290890" Then
        s = 0
       End If
        !creditamount = (RstInv!Sub_Amt + RstInv!transchg + RstInv!Tot_Vat) + RstInv!surcharge
        !Description = RstInv!invc_no
        !invoicedate = RstInv!invc_date
        !invoiceno = RstInv!invc_no
        !ClientCode = RstInv!cust_code
        !invoicedate = RstInv!invc_date
        !profitcenter = RstInv!factory
        !tradercvble = RstInv!tot_amt
        !TRadedisc = RstInv!Mngt_Acct
        !tradeDiscamt = RstInv!TDisAmount 'rstINV!Sub_Amt * (rstINV!Mngt_Acct / 100)
        !MgtDisc = 0
        !MgtDiscAmt = 0 ' rstINV!Sub_Amt * (rstINV!Mngt_Acct / 100)
        !GrossSales = RstInv!Sub_Amt 'post to Local Sales in Misc
        !NetSales = !GrossSales - !tradeDiscamt + RstInv!transchg
        !transpoCharge = RstInv!transchg
        !vat = RstInv!Tot_Vat
        !SurTaxRate = RstInv!SurTax
        !SurTaxAmt = RstInv!surcharge
        !taxCr = RstInv!Credit
       
        
        'cdrAmt = Format(!debitAmount, "###,###,###.#0")
        'ccramt = Format(!creditamount, "###,###,###.#0")
'        If !dEBITAmount > !creditamount Then
'           !GrossSales = Format(!GrossSales, "###,###,###.#0")
'           !vat = Format(!vat, "###,###,###.#0")
'           !NetSales = Format(!NetSales, "###,###,###.#0")
'           !creditamount = !GrossSales + !vat + !TranspoCharge + RstInv!Surcharge
'          Else
'           drAmt = Format(!dEBITAmount, "###,###,###.#0")
'           cramt = Format(!creditamount, "###,###,###.#0")
'           Diff = drAmt - cramt
'           If cramt - drAmt = 0.01 Then
'           !GrossSales = Format(!GrossSales, "###,###,###.#0")
'           !vat = Format(!vat, "###,###,###.#0")
'           !NetSales = Format(!NetSales, "###,###,###.#0")
'           !creditamount = !creditamount - 0.01
'           End If
'        End If
        !DebitAmount = Format(!DebitAmount, "###,###,###.#0")
        !creditamount = Format(!creditamount, "###,###,###.#0")
        !tradercvble = Format(!tradercvble, "###,###,###.#0")
        !tradeDiscamt = Format(!tradeDiscamt, "###,###,###.#0")
        !GrossSales = Format(!GrossSales, "###,###,###.#0")
        !vat = Format(!vat, "###,###,###.#0")
        !SurTaxAmt = Format(!SurTaxAmt, "###,###,###.#0")
        If !DebitAmount > !creditamount Then
          Dramt = !DebitAmount - 0.01
          !DebitAmount = Dramt
        End If
        '!creditamount = Format(!creditamount, "###,###,###.#0")
        .Update
        i = i + 1
        
        
     End With
    DoEvents
    RstInv.MoveNext
    If RstInv.EOF <> True Then
     vDate = RstInv!invc_date 'FOR INVOICE DATE
    End If
Loop

If cVal < 100 Then
    Me.ProgressBar1.Value = 100
End If
i = 0
cVal = 0
RstInv.Close
Unload Me


'if journal amount is not eqaul
Dim rsTotals As New ADODB.Recordset
Dim PrinterReady As Boolean
PrinterReady = True
rsTotals.Open "SELECT  TransDate, COUNT(TransDate) AS TotalTRan, SUM(DebitAmount) AS DrAmt, SUM(CreditAmount) AS CrAmt  " _
             & " From SalesJournal where remarks is null GROUP BY TransDate ORDER BY TransDate", constring, adOpenKeyset, adLockPessimistic, adCmdText
If rsTotals!Dramt <> rsTotals!CrAmt Then
   mess = MsgBox("Downloading is Finished, but Total Debit And Credit Amount is not equal!" & vbCrLf & _
                "Please ready the printer to print the unbalance invoice(s) and reconcile it with EDP", vbExclamation + vbOKOnly, "Warning")
   rsTotals.Close
   rsTotals.Open "Select * from SalesJOurnal where Debitamount <> CreditAMount order by InvoiceNo", constring, adOpenKeyset, adLockPessimistic, adCmdText
   On Error GoTo Nelson
   If PrinterReady = True Then
       Printer.FontName = "Arabic Transparent"
       Printer.Print
       Printer.Print
       Printer.Print "List of Unbalance Invoices"
       Printer.Print
       Printer.Print "Invoice#"; Tab(15); "Inv_Date"; Tab(30); "TR_Amt"; Tab(45); "TD_Amt"; Tab(60); "Gross"; Tab(75); "TC_Amt"; Tab(90); "VAT"; Tab(105); "SurTax "; Tab(120); "Debit"; Tab(140); "Credit"; Tab(160); "Diff"
       With rsTotals
        Do Until .EOF = True
            Printer.Print !invoiceno _
            ; Tab(15); Format(!invoicedate, "dd/mm/yyyy") _
            ; Tab(30); !tradercvble _
            ; Tab(45); !tradeDiscamt _
            ; Tab(60); !GrossSales _
            ; Tab(75); !transpoCharge _
            ; Tab(90); !vat _
            ; Tab(105); !SurTaxAmt _
            ; Tab(115); "=" _
            ; Tab(120); !DebitAmount _
            ; Tab(140); !creditamount _
            ; Tab(160); !DebitAmount - !creditamount
           .MoveNext
        Loop
        Printer.EndDoc
        'we delete generated sales jouranl because it not balance
        rsTotals.Close
        trDate = Format(Date, "dd/mm/yyyy")
        rsTotals.Open "Delete salesJournal where transdate= " & "'" & trDate & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
       End With
       Exit Sub
     Else
      Exit Sub
   End If
End If
    
msg = MsgBox("Finished Generating Sales Journal and it is OK", vbInformation + vbOKOnly, "Message")

Nelson:
c = Err.Number
If c = 480 Then
   PrinterReady = False
   msg = MsgBox("Printer not ready", vbExclamation + vbOKOnly, "Message")
  Else
  PrinterReady = True
End If
End Sub
Sub NewJn(Jn)
Dim rstjNo As New ADODB.Recordset
rstjNo.Open "Setup", constring, adOpenKeyset, adLockPessimistic, adCmdTable

If Val(Left(rstjNo!CurrentMoYr, 2)) <> Format(Date, "mm") Then
       rstjNo!CurrentMoYr = Format(Date, "mm") & Trim(Right(Year(Date), 2))
       rstjNo!nextjn = "00001"
       rstjNo.Update
   Else
       Jn = "SAL" & "-" & Right(Year(Date), 2) & Trim(Format(Date, "mm")) & "-" & Right(rstjNo!nextjn, 5)
       nextjn = Val(rstjNo!nextjn)
       If (Val(nextjn)) = 1 Then
        Zeros = "0000"
        ElseIf Len(nextjn) = 2 Then
        Zeros = "000"
        ElseIf Len(nextjn) = 3 Then
        Zeros = "00"
        ElseIf Len(nextjn) = 4 Then
        Zeros = "0"
        ElseIf Len(nextjn) = 5 Then
        Zeros = ""
       End If
       rstjNo!nextjn = Zeros & Trim(Val(nextjn) + 1)
       rstjNo.Update
       rstjNo.Close
End If

End Sub
Private Sub Command1_Click()

DownLoadINv

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Activate()
If reload = True Then
   msg = MsgBox("Load again the sales", vbInformation + vbOKOnly, "Message")
   reload = False
   Unload Me
   Exit Sub
End If

If ErrorConek = True Then
    ErrorConek = False
    Unload Me
End If

End Sub

Private Sub Form_Load()
If i = 0 Then
    On Error GoTo MyMsg
    Mainform.sbStatusBar.Panels(1).Text = "Status: Connecting now to external Database..."
    rsInv.Open "Delete Invoices", constring, adOpenKeyset, adLockPessimistic, adCmdText
    rsInv.Open "Invoices", constring, adOpenKeyset, adLockPessimistic, adCmdTable
    rsInv2.Open "INvcMain", "dsN=Invoices;UID=SA;PWD=;", adOpenKeyset, adLockPessimistic, adCmdTable
    cTotrec = Int(rsInv2.RecordCount / 100)
    Totrec = cTotrec
    i = 0
    cVal = 0
End If
Mainform.sbStatusBar.Panels(1).Text = "Status: Ready"

MyMsg:
 c = Err.Number
 d = Err.Description
 If c = -2147217887 Then
    reload = True
    rsInv2.Close
    Exit Sub
 End If
If c = 3705 Then
   Else
   X = Err.Description
   If c <> 0 Then
    ErrorConek = True
    MsgBox ("Maybe file is used by other user,try again later" & vbCrLf & _
            " ÑÈãÇ ÇáãáÝ íÓÊÎÏã ÈæÇÓØÉ ãÓÊÎÏã ÇÎÑ Íæá Ýí æÞÊ ÇÎÑ  ")
   Exit Sub
   End If
 End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Nelson
If rsInv2.EOF = False And cVal <> 0 Then
     mess = MsgBox("Are you sure you want to cancel downloading?" & vbCrLf & _
           " åá ÇäÊ ãÊÇßÏ ãä ÇáÛÇÁ ÇáÊÍãíá ", vbQuestion + vbYesNo, "Please confirm ãä ÝÖáß ÇáÊÇßíÏ")
    If mess = vbNo Then
        Cancel = -1
        'SendKeys "{Enter}"
      Else
         Unload Me
    End If
   Else
End If
Nelson:
Unload Me
End Sub

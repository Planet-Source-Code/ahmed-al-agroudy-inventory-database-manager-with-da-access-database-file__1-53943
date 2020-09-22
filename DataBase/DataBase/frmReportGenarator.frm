VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmReportGenarator2 
   Caption         =   "Report Generator"
   ClientHeight    =   3870
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   5715
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      DataField       =   "invc_no"
      DataSource      =   "AdoORmaster"
      Height          =   285
      Left            =   3000
      TabIndex        =   19
      Text            =   "Text2"
      Top             =   3480
      Width           =   375
   End
   Begin MSAdodcLib.Adodc AdoORmaster 
      Height          =   330
      Left            =   360
      Top             =   3480
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=BegiBal"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "BegiBal"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "sjmaster"
      Caption         =   "ORMaster"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   360
      Top             =   3120
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=BegiBal"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "BegiBal"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select * from marcusfl order by First_Name"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      DataField       =   "invc_no"
      DataSource      =   "SJ"
      Height          =   285
      Left            =   240
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSAdodcLib.Adodc SJ 
      Height          =   330
      Left            =   360
      Top             =   2760
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=BegiBal"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "BegiBal"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "sjmaster"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   240
      TabIndex        =   14
      Top             =   1560
      Width           =   1215
      Begin VB.OptionButton Option2 
         Caption         =   "Print One"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Print All"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Print one client"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1680
      TabIndex        =   9
      Top             =   1560
      Width           =   3975
      Begin VB.CheckBox Check1 
         Height          =   255
         Left            =   3480
         TabIndex        =   18
         Top             =   240
         Width           =   375
      End
      Begin VB.ComboBox CmbName 
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1200
         TabIndex        =   12
         Top             =   600
         Width           =   2655
      End
      Begin VB.ComboBox CmbCode 
         Height          =   315
         Left            =   1200
         TabIndex        =   10
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "Client Name"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Client Code"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Exit"
      Height          =   350
      Left            =   3600
      TabIndex        =   4
      ToolTipText     =   "close from window"
      Top             =   3480
      Width           =   2055
   End
   Begin VB.CommandButton cmdGenarator 
      Caption         =   "Generate Report"
      Height          =   350
      Left            =   3600
      TabIndex        =   3
      ToolTipText     =   "click to genarate Specified report"
      Top             =   3070
      Width           =   2055
   End
   Begin VB.Frame Frame3 
      Caption         =   "Period"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
      Begin MSMask.MaskEdBox Mskfrom 
         Height          =   325
         Left            =   840
         TabIndex        =   5
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MskTo 
         Height          =   330
         Left            =   840
         TabIndex        =   6
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         Caption         =   "To"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "From"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Select the option to print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "FrmReportGenarator2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RstCust As ADODB.Recordset 'this is to add the Cust-code for the Combo
Dim RstCustx As ADODB.Recordset ' this is to get the Name for that specified Code and to add the Combo
Dim RstCustY As ADODB.Recordset ' this is to get the Name for that specified Code and to add the Combo (RECEIVABKLES)
Dim rstVen As ADODB.Recordset
Dim CON1 As ADODB.Connection
Dim TempInv As ADODB.Recordset
Dim TempDeb As ADODB.Recordset
Dim TempCre As ADODB.Recordset
Dim TempBBal As ADODB.Recordset
Dim Temp As ADODB.Recordset
Dim ORmas As ADODB.Recordset
Dim RstCust2 As ADODB.Recordset
Dim CredmVouc As ADODB.Recordset 'For the Creditors Agoeong
Dim CredmInside As ADODB.Recordset 'For the Creditors Agoeong
Dim Credtemp As ADODB.Recordset 'For the Creditors Agoeong
Dim CredrstFinCr As ADODB.Recordset 'For the Creditors Agoeong
Dim rstVendor As ADODB.Recordset 'For the Creditors Agoeong
Dim rstPaySetup As ADODB.Recordset 'For the Creditors Aging
Dim CredRstCust2 As ADODB.Recordset 'For the Creditors Aging
Dim FilteredInv As ADODB.Recordset 'this is the table to save the Multi Inovices
Dim CredFilteredInv As ADODB.Recordset ' for the Creditors Aging
Dim mVouc As ADODB.Recordset 'this is the new Cash Voucher table
Dim mTemp2 As ADODB.Recordset 'this table is usfull to get the Paymode
Dim mInside As ADODB.Recordset 'this is to get the selected invoice number from the tepinvoice2
Dim rstFinCr As ADODB.Recordset 'this will be the final re.set to find out the Cr number and CheckAmt

Dim LedgVouc As ADODB.Recordset ' for Customer Ledger
Dim temp2 As ADODB.Recordset ' for Customer Ledger
Dim LMarkFl As ADODB.Recordset ' for Customer Ledger
Dim LfiltCL As ADODB.Recordset ' for Customer Ledger

Dim BeginBal As ADODB.Recordset 'for the Fox data
Dim CredMem As ADODB.Recordset 'for the Fox data
Dim DebMemo As ADODB.Recordset 'for the Fox data

Dim rstStmtFiltStatAcc As New ADODB.Recordset

Public Toootal

Dim inv, XC
Dim mfrom As Date
Dim mTo As Date

'This is for the Statement Of acc
'Public mStAgst, mStV, mStOR, mStMFL, mStFil, mStTemp, mStInvc, mStSJ  As String 'This is
Public StMFName, StMFcity, StMFTerms, StMFPTerms, StMFTelP1, StMFAdd1 As String
Public StMFcrlimit, StMFChkLmt As Currency

Public StVRecno, StVRecDt1, StVRecDt, StVPmode As String
Public StVCamt, StVCHKamt As Currency
Public StVchkDu

Public tAgtInvNo As String
Public tAgtInvAmt, tAgtUapplied As Currency
Public tAgtInvDate As String
Public DebNtNo As String
Public DebNtAmt As Currency
Public DebNtDT
Public CountTerm As Date
Public CrNtNo, CrNtDT  As String
Public CrNtAmt As Currency

Public Recno, outstandBal As String
Public invno1, invno2 As String

Public DisplayAmount As String
Public displayTerms As String
Public MyDetailVar As String
Public SupplimentInvNo As String
Public SupplimentInvDT As Date
Public SupplimentInvAmt As Currency
Public SupplimentInvDueDt As Date
Public LvBal As Currency
Public LvBalStmtAcc As Currency
Public Totalunpaidamt As Currency
Public TotalxCdays As Currency
Public Totalx30days As Currency
Public Totalx60days As Currency
Public Totalx90days As Currency
Public Totalxover As Currency
Public Totalcheckamt As Currency
Public TTotalChk_UP As Currency

Public LVDate As String
Public LVTransType As String
Public LvDocuNo As String
Public LvRemaks As String
Public LvDueDate As String
Public LvChkPyt As String
Public LvDb As String
Public LvCr As String
Public CustCode As String
Public TotCkPt As String
Public TotDbi As String
Public totCre As String
Public AccuTot As Currency
Dim XOrmas
Dim Xcust
Dim Xfilt
Dim Xvou
Dim XAnu
Dim XPaymodeTOcrno

Dim INo, IDate, IAmt, amt_paid
Dim MFName, MLName, MMName, MAdd1, MAdd2, Mcity, MTerms, MchkLimit, Mcrlimit, MTelP1, MtelP2
Dim Ono, Oamt, Odate
Dim BBal                           ' "'" & xAccount & "'"
Public SJinvNo, SjInvcDate, SjInvAmt, SjUnpaid, xDDue, xCurrent, x30, x60, x90, xover, SjCus
Dim Coll, t2CRno, t2AMT
Dim vPayMode
Dim FinalCRno
Public DateDue, ChkPayt, CR_no

Dim RStempagaintsinvoice As New ADODB.Recordset
Dim RsVoucher As New ADODB.Recordset

          Dim credmTrDate As Date


Private Sub Check1_Click()
If Me.Check1.Value = 1 Then
        Me.CmbCode.RightToLeft = True
        Me.CmbName.RightToLeft = True
Else
        Me.CmbName.RightToLeft = False
        Me.CmbCode.RightToLeft = False
End If
End Sub

Private Sub CmbCode_Click()
XC = CmbCode.Text

If Combo1.Text = "Aging of Account Payables" Then
    RstCustx.Open "Select * from VENDOR where VENCODE = " & "'" & XC & "'" & "", CON1, adOpenDynamic, adLockOptimistic

    If RstCustx.EOF = False Then
    RstCustx.MoveFirst
    While RstCustx.EOF = False
    'If CmbCode.Text = RstCust!cust_code Then
    Me.CmbName.Text = RstCustx!Vennameeng
    'End If
    RstCustx.MoveNext
    Wend
    End If
    RstCustx.Close
End If


If Combo1.Text = "Aging of Account Receivables" Or Combo1.Text = "Statement Of Account" Or Combo1.Text = "Debtors Outstanding Balances" Or Combo1.Text = "Customer Ledger Account" Then
   Me.Adodc1.Recordset.MoveFirst
   With Me.Adodc1
     While Me.Adodc1.Recordset.EOF = False
        If Trim(Me.CmbCode) = Me.Adodc1.Recordset!cust_code Then
         Me.CmbName = Me.Adodc1.Recordset!first_name '& " " & Me.Adodc1.Recordset!LAst_Name
        End If
        Me.Adodc1.Recordset.MoveNext
      Wend
    End With
    
    
End If


End Sub

Private Sub CmbCode_GotFocus()
Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(CmbCode.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)

End Sub

Private Sub CmbCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
CmbCode_Click
cmdGenarator.SetFocus
End If
End Sub

Private Sub CmbName_Click()

If Combo1.Text = "Aging of Account Receivables" Or Combo1.Text = "Statement Of Account" Or Combo1.Text = "Debtors Outstanding Balances" Or Combo1.Text = "Customer Ledger Account" Then
   Me.Adodc1.Recordset.MoveFirst
   With Me.Adodc1
     While Me.Adodc1.Recordset.EOF = False
        If Trim(Me.CmbName) = Trim(Me.Adodc1.Recordset!first_name) Then '& " " & Me.Adodc1.Recordset!LAst_Name Then
         Me.CmbCode = Me.Adodc1.Recordset!cust_code
        End If
        Me.Adodc1.Recordset.MoveNext
      Wend
    End With
    
    
End If








'If Combo1.Text = "Aging of Account Receivables" Then
'   If RstCust.EOF = False Then
'   RstCust.MoveFirst
'   End If
'
'While RstCust.EOF = False
'Me.CmbName.AddItem RstCust!firstname + RstCust!lastname
'RstCust.MoveNext
'Wend
'End If
End Sub
Private Sub CustLedg() 'This is the function used to so the customer ledger Part
 Dim LVouch
 Dim LFilLedg
 Dim Ltemp2
 Dim Lmfl
 Dim LFil
 Dim LMFName, LMLName, LMMName, LMAdd1, LMAdd2, LMcity, LMTerms, LMchkLimit, LMcrlimit, LMTelP1, LMtelP2
' Dim LVDate, LVTransType, LvDocuNo, LvRemaks, LvDueDate, LvChkPyt, LvDb, LvCr, LvBal
 Dim BBOBal, BBcCode, BBAmt
 Dim CMtransDt, CMType, CMcrNo, CMRem
 Dim DMtransDT, DMType, DMcrNo, DMRem
 Dim TmpAg
Dim TempAGinv As New ADODB.Recordset

'LVouch = "select * from vouchers where custNO = " & "'" & inv & "'" & " and receiptDate < " & "'" & mfrom & "'" & " and receiptDate > " & "'" & mTo & "'" & "  order by receiptDate"
LVouch = "select * from vouchers  where custNO = " & "'" & inv & "'" & " and deleted = '0'"

Ltemp2 = "select * from Temporary2"
Lmfl = "select * from marcusfl where cust_code = " & "'" & inv & "'" & ""
TmpAg = "select * from tempagaintsinvoice where custNO = " & "'" & inv & "'" & " and invoiceno <>  'Invoice Number' And invoiceno <> 'Invoice SubTotal' And invoiceno <> 'Un Applied Amount'"

LedgVouc.Open LVouch, CON1, adOpenDynamic, adLockOptimistic
temp2.Open Ltemp2, CON1, adOpenDynamic, adLockOptimistic
LMarkFl.Open Lmfl, CON1, adOpenDynamic, adLockOptimistic
TempAGinv.Open TmpAg, CON1, adOpenDynamic, adLockOptimistic

'Followings are the Foxpro tables this is the way to refer it
BeginBal.Open "select * from BEGINBALANCES where cust_code = " & "'" & inv & "'" & "", "dsn=BegiBal;uid=sa;pwd=;", adOpenKeyset, adLockOptimistic, adCmdText
'CredMem.Open "select * from credmain where cust_code = " & "'" & inv & "'" & " and Trans_DT < " & "'" & mfrom & "'" & " and Trans_DT > " & "'" & mTo & "'" & "", "dsn=CredMainX;uid=sa;pwd=;", adOpenKeyset, adLockOptimistic, adCmdText


'----------------------------------------------------------
'H E R E    I  H A V E   T O   D O  T H E   ADJUSTING ENTRY
'-----------------------------------------------------------




CustCode = inv

    If BeginBal.EOF = False Then
    BeginBal.MoveFirst
    While BeginBal.EOF = False   'Starts First Loop <BeginBalances>
    If BeginBal!Debit = 0 Then
    BBOBal = BeginBal!Credit
    Else
    BBOBal = BeginBal!Debit
    End If

    'This is to find out the Accumulated Balance for the Openning Balance
    Call AccuOpenBal
    '--------------------------------

   CredMem.Open "select * from credmain where cust_code = " & "'" & inv & "'" & "", "dsn=BegiBal;uid=sa;pwd=;", adOpenKeyset, adLockOptimistic, adCmdText
   DebMemo.Open "select * from debmain where cust_code = " & "'" & inv & "'" & "", "dsn=BegiBal;uid=sa;pwd=;", adOpenKeyset, adLockOptimistic, adCmdText

    LVDate = Format(mfrom, "dd/mm/yyyy")
    LVTransType = "Opening Balance"
    LvBal = Val(BBOBal) + Val(AccuTot)
    LvDb = Val(BBOBal) + Val(AccuTot)
    BeginBal.MoveNext
    Wend                        'End of First Loop <BeginBalances>

    Call AddnewFilteCustLedger
    End If

If LMarkFl.EOF = False Then
LMarkFl.MoveFirst
While LMarkFl.EOF = False       'Starts second Loop <MARCUSFL>
 LMFName = LMarkFl!first_name
 LMMName = LMarkFl!mid_name
 LMLName = LMarkFl!last_name
 LMAdd1 = LMarkFl!address1
 LMAdd2 = LMarkFl!address2
 LMcity = LMarkFl!city
 LMTerms = LMarkFl!Terms
 LMchkLimit = LMarkFl!chkclimit
 LMcrlimit = LMarkFl!crlimit
 LMTelP1 = LMarkFl!tel_no1
 LMtelP2 = LMarkFl!tel_no2

 'This is for the BegiBalance
 LvBal = IIf(IsNull(BBOBal), "0", (BBOBal))

LMarkFl.MoveNext
Wend                            'end of second Loop <MARCUSFL>
End If
LMarkFl.Close

 '-0-0-0-0-0-0-




SjCus = Trim(inv)
SJ.Recordset.MoveFirst
While SJ.Recordset.EOF = False   'Starts second

If (Trim(SJ.Recordset!cust_code) = Trim(inv) Or Trim(SJ.Recordset!mcustcode) = Trim(inv)) And Format(SJ.Recordset!invc_date, "dd/mm/yyyy") > Format(mfrom, "dd/mm/yyyy") And Format(SJ.Recordset!invc_date, "dd/mm/yyyy") < Format(mTo, "dd/mm/yyyy") Then

LvDocuNo = SJ.Recordset!invc_no
LVDate = SJ.Recordset!invc_date
LvDb = SJ.Recordset!unpaidamt
LVTransType = "Invoice"
LvBal = Val(LvBal) + Val(LvDb)
TotDbi = Val(TotDbi) + Val(LvDb)
Call AddnewFilteCustLedger
End If
SJ.Recordset.MoveNext            'End of second Loop
Wend


'-0--0-0-0-0-0-
'----------------------------------------------------------
'H E R E      PREVIOS O R ( DELPI'S CASH RECIEPT)
'-----------------------------------------------------------


AdoORmaster.Recordset.MoveFirst
While AdoORmaster.Recordset.EOF = False   'Starts second

If Trim(AdoORmaster.Recordset!orcustno) = Trim(inv) And Format(AdoORmaster.Recordset!ordate, "dd/mm/yyyy") > Format(mfrom, "dd/mm/yyyy") And Format(AdoORmaster.Recordset!ordate, "dd/mm/yyyy") < Format(mTo, "dd/mm/yyyy") Then

LvDocuNo = AdoORmaster.Recordset!ORno
LVDate = AdoORmaster.Recordset!ordate
LvDb = AdoORmaster.Recordset!oramt
LVTransType = "Receipt - 2002"
LvBal = Val(LvBal) + Val(LvDb)
TotDbi = Val(TotDbi) + Val(LvDb)
Call AddnewFilteCustLedger
End If
AdoORmaster.Recordset.MoveNext            'End of second Loop
Wend


'----------------------------------------------



    'This is for the CHECK
    If LedgVouc.EOF = False Then
    LedgVouc.MoveFirst
    End If
    While LedgVouc.EOF = False      'Statrt of Third Loop <Voucher>
    If LedgVouc!receiptdate > mfrom And LedgVouc!receiptdate < mTo And LedgVouc!paymode = "03     Check" Then

    ' If LedgVouc!custNO = inv And LedgVouc!receiptDate > mfrom And LedgVouc!receiptDate < mTo Then
    LVDate = LedgVouc!receiptdate
    LVTransType = "Cash Receipt" 'Mid(LedgVouc!payopt, 5)
    LvDocuNo = LedgVouc!receiptno
    LvRemaks = LedgVouc!remarks
    On Error Resume Next
    LvRemaks = LedgVouc!chkdue
    On Error GoTo 0 'This is when we not put the Date Due

    LvChkPyt = LedgVouc!checkreceipt
   ' LvDb = IIf((LedgVouc!dEBITAmount) = 0, "", (LedgVouc!dEBITAmount))
    LvCr = LedgVouc!checkreceipt
    TotCkPt = Val(TotCkPt) + Val(LvChkPyt)
    totCre = Val(TotCkPt) + Val(totCre)
'''''''''    TotDbi = Val(TotCkPt) + Val(TotDbi)
    ' LvBal = LvBal + Val(LvDb) - Val(LvCr)

    If LvCr = "" And LvChkPyt = "" Then
    LvBal = Val(LvBal) + Val(LvDb)
    ElseIf LvDb = "" And LvChkPyt = "" Then
    LvBal = Val(LvBal) - Val(LvCr)
    ElseIf LvCr = "" And LvDb = "" Then
    LvBal = Val(LvBal) + Val(LvChkPyt)

    End If
    Call AddnewFilteCustLedger
    End If
    LedgVouc.MoveNext
    Wend                                 'End of Third Loop <Voucher>



             'here it wold be one Loop for the Invoces <Tempagainstinvoices>
'        If TempAGinv.EOF = False Then
'        TempAGinv.MoveFirst
'        End If
'        While TempAGinv.EOF = False
'         LvDocuNo = TempAGinv!invoiceno
'         LVDate = TempAGinv!invoicedate
'         LVTransType = "Invoice"
'         LvDb = TempAGinv!Applied
'         Call AddnewFilteCustLedger
'        TempAGinv.MoveNext
'        Wend



'this is internal loop for to check the Credmain table
If CredMem.EOF = False Then
CredMem.MoveFirst
End If
While CredMem.EOF = False                   'Statrt of Forth Loop <credmain>
If CredMem!Trans_dt > mfrom And CredMem!Trans_dt < mTo Then
LVDate = CredMem!Trans_dt
LVTransType = "Credit Memo"
LvCr = IIf(IsNull(CredMem!tot_amt), 0, (CredMem!tot_amt))
LvDocuNo = CredMem!trans_no 'Always this is the Cerdit Balns
LvRemaks = CredMem!rem1
totCre = Val(totCre) + Val(LvCr)
LvBal = Val(LvBal) - Val(LvCr)
Call AddnewFilteCustLedger
End If

CredMem.MoveNext
Wend                                    'End of Forth Loop <credmain>


    'this is internal loop for to check the Debmain table
    If DebMemo.EOF = False Then
    DebMemo.MoveFirst
    End If
    While DebMemo.EOF = False               'Statrt of Fifth Loop <debmain>

     'credmTrDate = DebMemo!Trans_dt

    If DebMemo!Trans_dt > mfrom And DebMemo!Trans_dt < mTo Then
    LVDate = DebMemo!Trans_dt
    LVTransType = "Debit Note"
    LvDb = IIf(IsNull(DebMemo!tot_amt), 0, (DebMemo!tot_amt))
    LvDocuNo = DebMemo!trans_no
    LvRemaks = DebMemo!rem1
    TotDbi = Val(TotDbi) + Val(LvDb)
    LvBal = Val(LvBal) - Val(LvDb)
    Call AddnewFilteCustLedger

    End If
    DebMemo.MoveNext
    Wend                                      'End of Fifth Loop <debmain>



Call Reflect



Call Trigerit


'This is to addNew Records    <Temporary2>
With temp2
.AddNew
!cust_code = inv
!first_name = LMFName
!last_name = LMLName
!Address = LMAdd1
'!contperson=lmc
!Tel_No = LMTelP1
!Terms = LMTerms
!openingBal = BBOBal
!openbaldate = mfrom
!CreditLimit = LMcrlimit
!CheckLimit = LMchkLimit
!xFrom = mfrom
!xTo = mTo
'!cPAY =
!cUNPAID = Totalunpaidamt
!CCurrent = TotalxCdays
!c30 = Totalx30days
!c60 = Totalx60days
!c90 = Totalx90days
!c90Over = Totalxover
!TotsForAdding = TTotalChk_UP
!cCkPay = Totalcheckamt
On Error Resume Next
!TotBalance = LvBal
!TotDebit = TotDbi
!TotCredit = totCre
!TotChkpayment = TotCkPt
!lastbalance = Toootal
!LAstBalPLUSchk = Val(Toootal) + Val(Totalcheckamt)
On Error GoTo 0
' Totalcheckamt
.Update

End With

  passdate2 RepLedgerLast.Sections(2).Controls("label4")
  passdate3 RepLedgerLast.Sections(2).Controls("Label44")
  'passdate4 RepLedgerLast.Sections(5).Controls("Label45")

RepLedgerLast.Show 1

FrmReportGenarator.Command2.caption = "Delete Table"
'End If
End Sub
Private Sub Trigerit()
Dim CON1 As New ADODB.Connection
Dim Earlys As New ADODB.Recordset
Dim Lates As New ADODB.Recordset
Dim LAtu, EArlu
conStr1 = "Provider=MSDASQL;DSN=Finance;UID=sa; PWD=;"
CON1.Open conStr1
EArlu = "Select * from FilteredCustLedger2 order by Ldate"
Earlys.Open EArlu, CON1, adOpenDynamic, adLockOptimistic

LAtu = "Select * from FilteredCustLedger order by Ldate"
Lates.Open LAtu, CON1, adOpenDynamic, adLockOptimistic

If Earlys.EOF = False Then
Earlys.MoveFirst
End If
While Earlys.EOF = False
With Lates
mmmmm = IIf(IsNull(Earlys.Fields(7)), 0, (Earlys.Fields(7)))
lllll = IIf(IsNull(Earlys.Fields(8)), 0, (Earlys.Fields(8)))

Toootal = Val(Toootal) + Val(mmmmm) - Val(lllll)

.AddNew
        !cust_code = Earlys.Fields(0)
        !Ldate = Earlys.Fields(1)
        !LTransType = Earlys.Fields(2)
        !LdocuNum = Earlys.Fields(3)
        !LRemarks = Earlys.Fields(4)
        If Earlys.Fields(5) = "" Then
        Else
        !LDueDate = Earlys.Fields(5)
        End If
        !LChk = Earlys.Fields(6)
        !Ldb = IIf(IsNull(Earlys.Fields(7)), 0, (Earlys.Fields(7)))
        !Lcr = IIf(IsNull(Earlys.Fields(8)), 0, (Earlys.Fields(8)))
        !Lbal = Toootal
    .Update
End With

Earlys.MoveNext
Wend

End Sub
Private Sub AccuOpenBal()
Dim LALAL As New ADODB.Recordset
Dim CON1 As New ADODB.Connection
Dim CredMem2 As New ADODB.Recordset
Dim DebMem2 As New ADODB.Recordset

conStr1 = "Provider=MSDASQL;DSN=Finance;UID=sa; PWD=;"
CON1.Open conStr1

CustCode = inv

Dim Mambalam

Mambalam = "SELECT CUSTNO, SUM(debitamount) AS cashamount, SUM(checkreceipt) AS checkamount, SUM(debitamount) + SUM(checkreceipt) AS totalamount From vouchers WHERE (deleted = '0') AND (receiptDate <= " & "'" & Format(mfrom, "mm/dd/yyyy") & "'" & ") AND (CUSTNO = " & "'" & inv & "'" & ") GROUP BY CUSTNO"
LALAL.Open Mambalam, CON1, adOpenDynamic, adLockOptimistic
CredMem2.Open "select * from credmain where cust_code = " & "'" & inv & "'" & "", "dsn=BegiBal;uid=sa;pwd=;", adOpenKeyset, adLockOptimistic, adCmdText
DebMem2.Open "select * from debmain where cust_code = " & "'" & inv & "'" & "", "dsn=BegiBal;uid=sa;pwd=;", adOpenKeyset, adLockOptimistic, adCmdText

Dim Masir
If LALAL.EOF = False Then
Masir = LALAL!totalamount
End If

Dim LvCrm As Currency
Dim totCre2 As Currency
'this is internal loop for to check the Credmain table
If CredMem2.EOF = False Then
CredMem2.MoveFirst
End If

While CredMem2.EOF = False
If CredMem2!Trans_dt > mfrom Then
LvCrm = IIf(IsNull(CredMem2!tot_amt), 0, (CredMem2!tot_amt))
totCre2 = Val(totCre2) + Val(LvCrm)
'LvBal = Val(LvBal) - Val(LvCr)
End If
CredMem2.MoveNext
Wend
CredMem2.Close

Dim TotDbi2 As Currency
Dim LvDb2 As Currency

    'this is internal loop for to check the Debmain table
    If DebMem2.EOF = False Then
    DebMem2.MoveFirst
    End If
    While DebMem2.EOF = False
    If DebMem2!Trans_dt > mfrom Then
    LvDb2 = IIf(IsNull(DebMem2!tot_amt), 0, (DebMem2!tot_amt))
    'LvRemaks = DebMem2!rem1
    TotDbi2 = Val(TotDbi2) + Val(LvDb2)
  '  LvBal = Val(LvBal) - Val(LvDb)
     End If
    DebMem2.MoveNext
    Wend
    DebMem2.Close

AccuTot = Val(Masir) + Val(totCre2) - Val(TotDbi2)

End Sub
Private Sub AddnewFilteCustLedger()

Dim LfiltCL As New ADODB.Recordset
LFil = "select * from filteredCustLedger2"
LfiltCL.Open LFil, CON1, adOpenDynamic, adLockOptimistic

'This is to addNew records <filteredCustLedger>
With LfiltCL ' this will add the multi records of invoice for the printout for the particual customer
.AddNew
On Error Resume Next
!cust_code = inv
On Error GoTo 0
!Ldate = LVDate
!LTransType = Trim(LVTransType)
!LdocuNum = LvDocuNo
!LRemarks = LvRemaks
tt = "12:00:00 ?"
If Val(LvDueDate) = Val(tt) Then
!LDueDate = ""
Else
!LDueDate = LvDueDate
End If
!LChk = LvChkPyt
'On Error Resume Next
If LvDb = "" Then
Else
!Ldb = LvDb
End If

If LvCr = "" Then
Else
!Lcr = LvCr
End If

'''''''''''''''''''''''''''''''''''''''''&&&!lbal = LvBal


.Update

LVDate = ""
LVTransType = ""
LvDocuNo = ""
LvRemaks = ""
LvDueDate = ""
LvChkPyt = ""
LvDb = ""
LvCr = ""
'LvBal = 0
End With

End Sub

Private Sub Reflect() 'this is to get the Reflected part from Customer Aging of Account TO Customer Ledger Account
Dim Xcust
Dim XXfilt
Dim XXvou
Dim XXAnu
Dim XXPaymodeTOcrno

Dim YINo, YIDate, YAmt, Yamt_paid
'Dim MFName, MLName, MMName, MAdd1, MAdd2, Mcity, MTerms, MchkLimit, Mcrlimit, MTelP1, MtelP2
Dim YOno, YOamt, YOdate
Dim YBBal                           ' "'" & xAccount & "'"
Dim YSJinvNo, YSjInvcDate, YSjInvAmt, YSjUnpaid, YxDDue, YxCurrent, Yx30, Yx60, Yx90, Yxover, YSjCus
Dim YColl, Yt2CRno, Yt2AMT
Dim YvPayMode
Dim YFinalCRno
Dim YDateDue, YChkPayt, YCR_no



YColl = "Collections" 'this variable is for the XXvou definition


XXOrmas = "Select * from ormaster_2001 where orcustno = " & "'" & inv & "'" & ""
Xcust = "Select * from MARCUSFL where cust_code = " & "'" & inv & "'" & ""
XXfilt = "Select * from filteredInvoice "

ORmas.Open XXOrmas, CON1, adOpenDynamic, adLockOptimistic
RstCust2.Open Xcust, CON1, adOpenDynamic, adLockOptimistic
FilteredInv.Open XXfilt, CON1, adOpenDynamic, adLockOptimistic


If ORmas.EOF = True And RstCust2.EOF Then
 MsgBox "No Records for the Report", vbOKOnly, "Empty Records"

TempCre.Close
TempBBal.Close
ORmas.Close
RstCust2.Close
FilteredInv.Close
Exit Sub
End If



lla = 0

    Dim TotalChecks
    TotalChecks = 0
    Pmode = "03     Check"
    Dim RsVoucher2 As New ADODB.Recordset
    Xvaw = "select CUSTNO, SUM(checkreceipt) AS MothaChecku from vouchers where (deleted = '0') AND (receiptDate <= " & "'" & Format(mTo, "mm/dd/yyyy") & "'" & ") and(paymode = " & "'" & Pmode & "'" & ") AND (CUSTNO = " & "'" & inv & "'" & ") GROUP BY CUSTNO"
    RsVoucher2.Open Xvaw, CON1, adOpenDynamic, adLockOptimistic
    If RsVoucher2.EOF = False Then
    YChkPayt = RsVoucher2!MothaChecku
    End If





'This will check all the INVOICE NUMBERS under the condition "If SJ.Recordset!cust_code = Trim(inv) And SJ.Recordset!invc_date < mTo Then"
'and will put all the multy Invoice details in the FILTERINVOICE table

YSjCus = Trim(inv)
SJ.Recordset.MoveFirst
While SJ.Recordset.EOF = False
'If SJ.Recordset!cust_code = Trim(inv) And SJ.Recordset!invc_date > mfrom And SJ.Recordset!invc_date < mTo Then
If SJ.Recordset!cust_code = Trim(inv) And Format(SJ.Recordset!invc_date, "dd/mm/yyyy") > Format(mfrom, "dd/mm/yyyy") And Format(SJ.Recordset!invc_date, "dd/mm/yyyy") < Format(mTo, "dd/mm/yyyy") Then
YSJinvNo = SJ.Recordset!invc_no
YSjInvcDate = SJ.Recordset!invc_date
YSjInvAmt = SJ.Recordset!tot_amt
YSjUnpaid = SJ.Recordset!unpaidamt
'YxDDue = Format(YSjInvcDate, "dd/mm/yyyy") + MTerms
YxDDue = DateAdd("d", MTerms, YSjInvcDate)

  Dim Myval

Myval = DateDiff("d", YxDDue, mTo) 'this deducts the Date(As of Date) from Due Date to count
                                 '- which colomn said to be taken the Amout

If Myval < 30 Then
YxCurrent = YSjUnpaid
ElseIf Myval < 61 And Myval > 30 Then
Yx30 = YSjUnpaid
ElseIf Myval < 91 And Myval > 60 Then
Yx60 = YSjUnpaid
ElseIf Myval > 92 Then
Xxover = YSjUnpaid
End If



       ' With FilteredInv ' this will add the multi records of invoice for the printout for the particual customer
        '.AddNew
         Totalunpaidamt = Val(Totalunpaidamt) + Val(YSjUnpaid)
         TotalxCdays = Val(TotalxCdays) + Val(IIf(IsNull(YxCurrent), "0", (YxCurrent)))

         Totalx30days = Val(Totalx30days) + Val(Yx30)
         Totalx60days = Val(Totalx60days) + Val(Yx60)
         Totalx90days = Val(Totalx90days) + Val(Yx90)
         Totalxover = Val(Totalxover) + Val(Xxover)
         Totalcheckamt = Val(Totalcheckamt) + Val(YChkPayt)
         TTotalChk_UP = Val(Totalcheckamt) + Val(Totalunpaidamt) 'this is to find the Total Balace


        YChkPayt = ""
       ' End With
End If
SJ.Recordset.MoveNext
Wend

RstCust2.Close
FilteredInv.Close

End Sub
Private Sub CreditorsAging()
Dim CredXOrmas
Dim CredXcust
Dim CredXfilt
Dim CredXvou
Dim CredXAnu
Dim CredXPaymodeTOcrno

Dim CredINo, CredIDate, CredIAmt, Credamt_Credpaid
Dim CredMFName, CredMLName, CredMMName, CredMAdd1, CredMAdd2, CredMcity, CredMTerms, CredMchkLimit, CredMcrlimit, CredMTelP1, CredMtelP2
Dim CredOno, CredOamt, CredOdate
Dim CredBBal                           ' "'" & xAccount & "'"
Dim CredSJinvNo, CredSjInvcDate, CredSjInvAmt, CredSjUnpaid, CredxDDue, CredxCurrent, Credx30, Credx60, Credx90, Credxover, CredSjCus
Dim CredColl, Credt2CRno, Credt2AMT
Dim CredvPayMode
Dim CredFinalCRno
Dim CredDateDue, CredChkPayt, CredCR_no
Dim Cred
Coll = "Collections" 'this variable is for the Xvou definition


'CredXOrmas = "Select * from ormaster_2001 where orcustno = " & "'" & inv & "'" & ""
CredXcust = "Select * from vendor where vencode = " & "'" & inv & "'" & ""
CredXfilt = "Select * from filteredInvoiceCreditors "
PAyset = "Select * from Payablesetup where vencode = " & "'" & inv & "'" & ""
Cred = "select * from TempoCredAge"

'ORmas.Open XOrmas, con1, adOpenDynamic, adLockOptimistic
CredRstCust2.Open CredXcust, CON1, adOpenDynamic, adLockOptimistic
CredFilteredInv.Open CredXfilt, CON1, adOpenDynamic, adLockOptimistic
rstPaySetup.Open PAyset, CON1, adOpenDynamic, adLockOptimistic
Credtemp.Open Cred, CON1, adOpenDynamic, adLockOptimistic




'This is to get the Payee Details from table Payee

If CredRstCust2.EOF = False Then
CredRstCust2.MoveFirst
While CredRstCust2.EOF = False
 CredMFName = CredRstCust2!Vennameeng
' CredMMName = CredRstCust2!Mid_name
' CredMLName = CredRstCust2!last_name
' CredMAdd1 = CredRstCust2!address1
' CredMAdd2 = CredRstCust2!address2
 CredMcity = CredRstCust2!venhomecty
 CredMTerms = CredRstCust2!venTerms
' CredMchkLimit = CredRstCust2!chklimit
 CredMcrlimit = CredRstCust2!vencrlimit
 CredMTelP1 = CredRstCust2!venhometel
 'CredMtelP2 = CredRstCust2!tel_no2
CredRstCust2.MoveNext
Wend
Else
End If



'This will check all the INVOICE NUMBERS under the condition "If SJ.Recordset!cust_code = Trim(inv) And SJ.Recordset!invc_date < mTo Then"
'and will put all the multy Invoice details in the FILTERINVOICE table

 If rstPaySetup.EOF = False Then
 rstPaySetup.MoveFirst
 While rstPaySetup.EOF = False
            credsjrefno = rstPaySetup!RefNo
            CredSJinvNo = rstPaySetup!invoiceno
            Dim YG As Date

            YG = Format(rstPaySetup!invoicedate, "DD/MM/YYYY")
            CredSjInvcDate = Format(YG, "DD/MM/YYYY")     ' Format(rstPaySetup!invoicedate, "DD/MM/YYYY")
            CredSjInvAmt = rstPaySetup!invAmt
            CredSjUnpaid = rstPaySetup!outbal
       '     CredxDDue = Format(CredSjInvcDate, "mm/dd/yyyy") - CredMTerms

            CredxDDue = DateAdd("d", CredMTerms, Format(CredSjInvcDate, "dd/mm/yyyy"))

Dim CredMyval
CredMyval = DateDiff("d", CredxDDue, mTo) 'this deducts the Date(As of Date) from Due Date to count
                                  '- which colomn said to be taken the Amout
If CredMyval < 30 Then
CredxCurrent = CredSjUnpaid
ElseIf CredMyval < 61 And CredMyval > 30 Then
Credx30 = CredSjUnpaid
ElseIf CredMyval < 91 And CredMyval > 60 Then
Credx60 = CredSjUnpaid
End If

    Dim CredxMyMode
    CredxMyMode = "03     Check"
    CredXvou = "select * from vouchers where CUSTNAME = " & "'[" & inv & "]'" & " and PAYMODE = " & "'" & CredxMyMode & "'" & ""
    CredmVouc.Open CredXvou, CON1, adOpenDynamic, adLockOptimistic
    If CredmVouc.EOF = False Then
    CredmVouc.MoveFirst
    End If
    While CredmVouc.EOF = False
    CredDateDue = CredmVouc!chkdue
    t2CRno = CredmVouc!receiptno
    CredChkPayt = CredmVouc!checkreceipt



        With CredFilteredInv ' this will add the multi records of invoice for the printout for the particual customer
        .AddNew
         !cust_code = inv
         !Inum = CredSJinvNo
         If CredSjInvcDate = "" Then
         Else
         !IDate = CredSjInvcDate
         End If
         !IAmt = CredSjInvAmt
         If CredSjUnpaid = "" Then
         Else
         !unpaidamt = CredSjUnpaid
         End If
         If CredxDDue = "" Then
         Else
         !DateDue = CredxDDue
         End If
         !RefNo = credsjrefno
         On Error Resume Next
         !XCdays = IIf(IsNull(CredxCurrent), "0", (CredxCurrent))

         !x30days = Credx30
         !x60days = Credx60
         !x90days = Credx90
         '!xover = Credxover
         On Error GoTo 0
         !CRno = Credt2CRno
         !CheckAmt = CredChkPayt
         !CHKValueDate = CredDateDue
          !TotalChk_UP = Val(CredChkPayt) + Val(CredSjUnpaid) 'this is to find the Total Balace
         .Update
         Credt2CRno = ""
         CredChkPayt = ""
         CredDateDue = ""

'        CredSjCus = ""
'        CredSjInvAmt = ""
'        CredSJinvNo = ""
'        CredSjInvcDate = ""
'        CredSjUnpaid = ""
'        CredxDDue = ""
'        CredxCurrent = ""
'        credsjrefno = ""
'        Credx30 = ""
'        Credx60 = ""
'        Credx90 = ""
'        Credxover = ""



        End With

    CredmVouc.MoveNext
    Wend
    CredmVouc.Close

        CredSjCus = ""
        CredSjInvAmt = ""
        CredSJinvNo = ""
        CredSjInvcDate = ""
        CredSjUnpaid = ""
        CredxDDue = ""
        CredxCurrent = ""
        credsjrefno = ""
        Credx30 = ""
        Credx60 = ""
        Credx90 = ""
        Credxover = ""


rstPaySetup.MoveNext
Wend
End If
'CredRstCust2.Close
CredFilteredInv.Close

With Credtemp
.AddNew
            Credtemp!cust_code = inv
            Credtemp!first_name = CredMFName
            Credtemp!last_name = CredMLName
            Credtemp!Address = CredMAdd1
            Credtemp!Tel_No = CredMTelP1
            Credtemp!Terms = CredMTerms
            Credtemp!openingBal = CredBBal
            Credtemp!CreditLimit = CredMcrlimit
            Credtemp!CheckLimit = CredMchkLimit
            '!xfrom = Me.Mskfrom.Text
            !xTo = Me.MskTo.Text
                .Update
End With
  Credtemp.Close
  passdate AgingCreditors.Sections(2).Controls("lblCrAsOf")
  Me.Command2.caption = "Delete Table"
  AgingCreditors.Show 1 'this is to show the Report
End Sub
Private Sub StatementOfAcc()
 Dim rstStmtAgstInv As New ADODB.Recordset
 Dim rstStmtVoucher As New ADODB.Recordset
 Dim rstStmtOrmas As New ADODB.Recordset
 Dim rstStmtMFL As New ADODB.Recordset
' Dim rstStmtFiltStatAcc As New ADODB.Recordset
 Dim rstStmtTemp3 As New ADODB.Recordset
 Dim rstStmtInvc As New ADODB.Recordset
 Dim rstStmtCredMem As New ADODB.Recordset
 Dim rstStmtDebMemo As New ADODB.Recordset
 Dim rstStmtSJ1 As New ADODB.Recordset
 Dim rstStmtSJ As New ADODB.Recordset
 Dim rstKAsolai As New ADODB.Recordset

 Dim BeginBalFORStatmentAC As New ADODB.Recordset
 
 Dim RstOrmast As New ADODB.Recordset
 
Dim mStAgst, mStV, mStOR, mStMFL, mStFil, mStTemp, mStInvc, mStSJ
invno1 = "Invoice Number"

'mStInvc = "Select * from invoices where cust_code = " & "'" & inv & "'" & " and  trans_Dt > " & "'" & mTo & "'" & " "
mStInvc = "Select * from invoices where cust_code = " & "'" & inv & "'" & " "
'mStOR = "Select * from ormaster_2001 where orcustno = " & "'" & inv & "'" & ""
mStMFL = "Select * from MARCUSFL where cust_code = " & "'" & inv & "'" & ""

mStTemp = "select * from Temporary3"      'Format(CredSjInvcDate, "dd/mm/yyyy"))


rstStmtInvc.Open mStInvc, CON1, adOpenDynamic, adLockOptimistic
rstStmtMFL.Open mStMFL, CON1, adOpenDynamic, adLockOptimistic
rstStmtTemp3.Open mStTemp, CON1, adOpenDynamic, adLockOptimistic

rstStmtCredMem.Open "select * from credmain where cust_code = " & "'" & inv & "'" & "", "dsn=begibal;uid=sa;pwd=;", adOpenKeyset, adLockOptimistic, adCmdText


'This is to get the Customer Details
    
'Dim StMFName, StMFcity, StMFTerms, StMFcrlimit, StMFTelP1, StMFAdd1
If rstStmtMFL.EOF = False Then
rstStmtMFL.MoveFirst
End If

'-----------------------------------
'C U S T O M E R     D E T A I L S
'-----------------------------------


While rstStmtMFL.EOF = False
 StMFName = rstStmtMFL!first_name
 StMFAdd1 = IIf(IsNull(rstStmtMFL!address1), "", (rstStmtMFL!address1))
 StMFcity = IIf(IsNull(rstStmtMFL!city), "", (rstStmtMFL!city))
 StMFTerms = IIf(IsNull(rstStmtMFL!Terms), "", (rstStmtMFL!Terms))
 StMFChkLmt = IIf(IsNull(rstStmtMFL!chkclimit), 0#, (rstStmtMFL!chkclimit))
 StMFcrlimit = IIf(IsNull(rstStmtMFL!crlimit), 0#, (rstStmtMFL!crlimit))
 StMFTelP1 = IIf(IsNull(rstStmtMFL!tel_no1), "", (rstStmtMFL!tel_no1))
 StMFPTerms = IIf(IsNull(rstStmtMFL!pTerms), "", (rstStmtMFL!pTerms))
'Here do the Loop for the display allowance and display terms
                   'and passs the variable to rststmtTemp3
'                    mStSJ = "Select * from Sjmaster "
                    rstStmtSJ.Open "select * from Sjmaster where cust_code = " & "'" & inv & "'" & "", "dsn=BegiBal;uid=sa;pwd=;", adOpenKeyset, adLockOptimistic, adCmdText
                    Dim IJKLM
                    
                    If rstStmtSJ.EOF = False Then
                    rstStmtSJ.MoveFirst
                    End If
                    
                    While rstStmtSJ.EOF = False
                    If rstStmtSJ!LDIsp = True Then
                    IJKLM = OK
                    End If
                    
                    rstStmtSJ.MoveNext
                    Wend
                    rstStmtSJ.Close
 If IJKLM = OK Then
 DisplayAmount = IIf(IsNull(rstStmtMFL!displamt), 0#, (rstStmtMFL!displamt))
 displayTerms = IIf(IsNull(rstStmtMFL!dispdays), 0#, (rstStmtMFL!dispdays))
 End If


'-------------------------------------
'B E G I N I N G    B A L A N C E
'-------------------------------------

BeginBalFORStatmentAC.Open "select * from BEGINBALANCES where cust_code = " & "'" & inv & "'" & "", "dsn=BegiBal;uid=sa;pwd=;", adOpenKeyset, adLockOptimistic, adCmdText
CustCode = inv

    If BeginBalFORStatmentAC.EOF = False Then
    BeginBalFORStatmentAC.MoveFirst
    While BeginBalFORStatmentAC.EOF = False   'Starts First Loop <BeginBalances>
    If BeginBalFORStatmentAC!Debit = 0 Then
    BBOBal = BeginBalFORStatmentAC!Credit
    Else
    BBOBal = BeginBalFORStatmentAC!Debit
    End If


    tAgtInvDate = Format(mTo, "dd/mm/yyyy")
    MyDetailVar = "Begining Balance"
    LvBalStmtAcc = Val(BBOBal) + Val(LvBalStmtAcc)
    
    BeginBalFORStatmentAC.MoveNext
    Wend                        'End of First Loop <BeginBalances>
    tAgtInvAmt = LvBalStmtAcc
    Call AddnewFiltStatAcc
    End If


'-------------------------------------
'I N V O I C E   D E T A I L S
'-------------------------------------
'I HAVE TO GET THE  INVOICE DETAILS FROM SJMASTER

rstStmtSJ1.Open "select * from Sjmaster where cust_code = " & "'" & inv & "'" & "", "dsn=Begibal;uid=sa;pwd=;", adOpenKeyset, adLockOptimistic, adCmdText
  If rstStmtSJ1.EOF = False Then
rstStmtSJ1.MoveFirst
End If
While rstStmtSJ1.EOF = False
'If rstStmtSJ1!invc_no <> "Un Applied Amount" Then
tAgtInvNo = rstStmtSJ1!invc_no
tAgtInvAmt = rstStmtSJ1!tot_amt
tAgtInvDate = rstStmtSJ1!invc_date

    '--------------------------------
    'THIS IS TO COUNT THE DATE DUE
    '--------------------------------
    If StMFPTerms = "COD" Or StMFPTerms = "CIA" Then
    CountTerm = rstStmtSJ1!invc_date
    ElseIf StMFPTerms = "O30" Then
    CountTerm = DateAdd("d", 30, rstStmtSJ1!invc_date)
    ElseIf StMFPTerms = "O60" Then
    CountTerm = DateAdd("d", 60, rstStmtSJ1!invc_date)
    ElseIf StMFPTerms = "O90" Then
    CountTerm = DateAdd("d", 90, rstStmtSJ1!invc_date)
    ElseIf StMFPTerms = "ONN" Then
    CountTerm = DateAdd("d", StMFTerms, rstStmtSJ1!invc_date)
    End If

'-----------------------------------
'P A Y M E N T    D E T A I L S          f r o m       T e m p o r a r y A g a i n s t   I n v o i c e
'-----------------------------------

                    mStAgst = "select * from tempagaintsinvoice where invoiceno <> " & " '" & tAgtInvNo & "'" & " And invoiceno <> 'Invoice SubTotal'"
                    rstStmtAgstInv.Open mStAgst, CON1, adOpenDynamic, adLockOptimistic
                    If rstStmtAgstInv.EOF = False Then
                    rstStmtAgstInv.MoveFirst
                    End If
                    While rstStmtAgstInv.EOF = False
                 '   If rstStmtAgstInv!display = "F" Then
                         StVRecno = rstStmtAgstInv!receiptno
                         StVCamt = rstStmtAgstInv!Applied
             
             
'-----------------------------------
'P A Y M E N T    D E T A I L S     FROM VOUCHER  ( NOT CHECK)
'-----------------------------------
                   Dim VariabVoucher
                  VariabVoucher = "select * from vouchers where receiptno = " & " '" & StVRecno & "'" & " and Paymode <> '03 Check' "
                    rstStmtVoucher.Open VariabVoucher, CON1, adOpenDynamic, adLockOptimistic
                    
                     While rstStmtVoucher.EOF = False
                             StVRecDt = IIf(IsNull(rstStmtVoucher!receiptdate), "", (rstStmtVoucher!receiptdate))
                             StVPmode = IIf(IsNull(rstStmtVoucher!paymode), "", (rstStmtVoucher!paymode))
                     '         If rstStmtAgstInv!Paymode = "03 Check" Then
                     '         StVCHKamt = rstStmtVoucher!checkreceipt
                     '         StVchkDu = rstStmtVoucher!chkdue
                     '         end if
                     rstStmtVoucher.MoveNext
                     Wend
                    rstStmtVoucher.Close
             Call AddnewFiltStatAcc

            rstStmtAgstInv.MoveNext
            Wend
            rstStmtAgstInv.Close
 
'PAYMENT DETAILS FROM THE OR MASTER

'RstOrmast.Open "select * from ormaster where orcustno = " & "'" & inv & "'" & "", "dsn=Begibal;uid=sa;pwd=;", adOpenKeyset, adLockOptimistic, adCmdText
 
 Call AddnewFiltStatAcc
rstStmtSJ1.MoveNext
'End If
Wend



rstStmtDebMemo.Open "select * from debmain where cust_code = " & "'" & inv & "'" & "", "dsn=begibal;uid=sa;pwd=;", adOpenKeyset, adLockOptimistic, adCmdText
If rstStmtDebMemo.EOF = False Then
rstStmtDebMemo.MoveFirst
Else
'Call AddnewFiltStatAcc
End If
While rstStmtDebMemo.EOF = False
tAgtInvNo = rstStmtDebMemo!SERIES_NO
'DebNtAmt = rstStmtDebMemo!tot_amt
tAgtInvAmt = rstStmtDebMemo!Balance 'Here I have to Makesure which field they used for the amount
tAgtInvDate = rstStmtDebMemo!Trans_dt


                         If MyFirstCheck = "Yes2" Then
                         Else
                         MyDetailVar = "DEBIT NOTE"
                         MyFirstCheck = "Yes2"
                         End If
                         


Call AddnewFiltStatAcc
rstStmtDebMemo.MoveNext
Wend
rstStmtDebMemo.Close

'-------------------------------------
'UNUPPLIED PAYMENT PART FOR CHECK
'-------------------------------------

Dim Natham
Natham = "select * from vouchers where  Paymode = '03 Check' and  custno = " & "'" & inv & "'" & ""
rstKAsolai.Open Natham, CON1, adOpenDynamic, adLockOptimistic
On Error Resume Next
rstKAsolai.MoveFirst
While rstKAsolai.EOF = False
 StVRecno = rstKAsolai!receiptno
 StVCHKamt = rstKAsolai!checkreceipt
 StVchkDu = rstKAsolai!chkdue
Call AddnewFiltStatAcc
rstKAsolai.MoveNext
Wend

'-------------------------------------
' PEE UNUPPLIED PAYMENT PART FOR CASH  HENA
'-------------------------------------


With rstStmtTemp3
    .AddNew
!cust_code = inv
!first_name = StMFName
!Address = StMFAdd1
!Tel_No = StMFTelP1
!Terms = StMFTerms
!CreditLimit = StMFcrlimit
!CheckLimit = StMFChkLmt
!xTo = mTo
!InvoiceNumber = tAgtInvNo
'!InvDate=
If tAgtInvAmt = "" Then
Else
!invAmt = tAgtInvAmt
End If
'!TotInvAmt=
!DebNtNo = DebNtNo
If DebNtDT = "" Then
Else
!DebNtDate = DebNtDT
End If

!DebNtAmt = DebNtAmt
!Unapplied = tAgtUapplied
!Display_Amt = DisplayAmount
!Display_Terms = displayTerms
!a = "Supplimental Statement of Display Account"
!b = "Display Allowance"
!d = "Display Terms"

  .Update
End With

rstStmtMFL.MoveNext
Wend


rstStmtTemp3.Requery
rstStmtTemp3.Close
'DataEnvironment1.rsStatementOfAcc.Requery
passdate5 RepStatementOfAcc.Sections(2).Controls("lblforPer")
RepStatementOfAcc.Show
Command2.caption = "Delete Table"
Command2.Width = 2000
End Sub
Private Sub SupplimentDisplay()

 Dim rstStmtSJ1 As New ADODB.Recordset

If rstStmtSJ.EOF = False Then
rstStmtSJ.MoveFirst
End If


End Sub
Private Sub AddnewFiltStatAcc()

Dim rstStmtFiltStatAcc As New ADODB.Recordset
'mStFil = "Select * from FilterdStatementOfAc order by Invoice_no"
rstStmtFiltStatAcc.Open "Select * from FilterdStatementOfAc ", CON1, adOpenDynamic, adLockOptimistic, adCmdText

outstandBal = Val(tAgtInvAmt) - Val(StVCHKamt) - Val(CrNtAmt) + Val(DebNtAmt)

            'This is to add the Details
            '******* add new here
            'This is to add the Filter
                With rstStmtFiltStatAcc
                .AddNew
                !xDueDate = CountTerm
                !Invoice_No = tAgtInvNo
                If tAgtInvDate = "" Then
                Else
                !Inv_Date = tAgtInvDate
                End If
                '!Inv_Amt =
                !Inv_Settl_OrRef = StVRecno
                !Inv_Settl_Date = StVRecDt
                !Inv_Settl_Mode = StVPmode
                If StVCamt = "" Then

                Else
                !Inv_Settl_CAmt = StVCamt
                End If
                !Inv_Settl_Chk = StVCHKamt
                If StVchkDu = "" Then

                Else
                !Inv_Settl_ValDate = StVchkDu
                End If
                !Inv_Settl_CN = CrNtNo
                !Inv_Settl_CNamt = CrNtAmt
                !Inv_Out_Bal_tot = outstandBal
                !cust_code = inv
                !Inv_DN_No = DebNtNo
                !Inv_DN_Date = DebNtDT
                !Inv_DN_Amt = DebNtAmt
                !Invoice_No = tAgtInvNo
                !Explanation = MyDetailVar
                '!Inv_Date =
                If tAgtInvAmt = "" Then

                Else
                !Inv_Amt = tAgtInvAmt
                End If
                    !Inv_Out_Bal_tot = outstandBal
                    !TotCHKandOutstanding = Val(StVCHKamt) + Val(outstandBal)
                .Update
                End With

        tAgtInvDate = ""
        tAgtInvNo = ""
        StVRecDt = ""
        StVRecno = ""
        StVRecDt = ""
        StVPmode = ""
        StVCamt = ""
        On Error Resume Next
        StVCHKamt = ""
        StVchkDu = ""
        CrNtNo = ""
        CrNtAmt = ""
        outstandBal = ""
        MyDetailVar = ""
        DebNtNo = ""
        DebNtDT = ""
        tAgtInvNo = ""
        tAgtInvAmt = ""








End Sub


Private Sub cmdGenarator_Click()
If MskTo.Text = "__/__/____" Then
MsgBox "Enter the Period"
Exit Sub
End If

inv = CmbCode.Text
mTo = Format(MskTo.Text, "dd/mm/yyyy")
If Combo1.Text <> "Aging of Account Receivables" And Combo1.Text <> "Aging of Account Payables" And Combo1.Text <> "Statement Of Account" And Combo1.Text <> "Debtors Outstanding Balances" Then
mfrom = Mskfrom.Text
End If

If Combo1.Text = "Customer Ledger Account" Then
Call CustLedg
Exit Sub

ElseIf Combo1.Text = "Aging of Account Payables" Then
Call CreditorsAging
Exit Sub

ElseIf Combo1.Text = "Statement Of Account" Then
Call StatementOfAcc
Exit Sub

ElseIf Combo1.Text = "Aging of Account Receivables" Then
Call AgingReceivable
Exit Sub
End If

End Sub
Private Sub passdate2(a As RptLabel)
a.caption = Mskfrom.Text
End Sub
Private Sub passdate(a As RptLabel)
a.caption = "As Of  " & MskTo.Text
End Sub
Private Sub passdate3(b As RptLabel)
b.caption = "To " & MskTo.Text
End Sub
Private Sub passdate4(c As RptLabel)
b.caption = Toootal
End Sub
Private Sub passdate5(d As RptLabel)
d.caption = MskTo.Text
End Sub

Private Sub AgingReceivable()

Coll = "Collections" 'this variable is for the Xvou definition


Xcust = "Select * from MARCUSFL where cust_code = " & "'" & inv & "'" & ""
Xfilt = "Select * from filteredInvoice "

RstCust2.Open Xcust, CON1, adOpenDynamic, adLockOptimistic
FilteredInv.Open Xfilt, CON1, adOpenDynamic, adLockOptimistic


If RstCust2.EOF Then     '@Doubt
 MsgBox "No Records for the Report", vbOKOnly, "Empty Records"


'TempDeb.Close
TempCre.Close
TempBBal.Close
ORmas.Close
RstCust2.Close
FilteredInv.Close
Exit Sub
End If



If RstCust2.EOF = False Then
RstCust2.MoveFirst
End If

While RstCust2.EOF = False            'Starts First Loop <MarcusFL>
MFName = RstCust2!first_name
MMName = RstCust2!mid_name
MLName = RstCust2!last_name
MAdd1 = RstCust2!address1
MAdd2 = RstCust2!address2
Mcity = RstCust2!city
MTerms = RstCust2!Terms
MchkLimit = RstCust2!chkclimit
Mcrlimit = RstCust2!crlimit
MTelP1 = RstCust2!tel_no1
MtelP2 = RstCust2!tel_no2

RstCust2.MoveNext
Wend                                'End of First Loop <MarcusFL>
RstCust2.Close


    Dim TotalChecks
    TotalChecks = 0
    Pmode = "03     Check"

    '    Mambalam = "SELECT CUSTNO, SUM(debitamount) AS cashamount, SUM(checkreceipt) AS checkamount, SUM(debitamount) + SUM(checkreceipt) AS totalamount From vouchers WHERE (deleted = '0') AND (receiptDate <= " & "'" & Format(mfrom, "mm/dd/yyyy") & "'" & ") AND (CUSTNO = " & "'" & inv & "'" & ") GROUP BY CUSTNO"
    Xvaw = "select CUSTNO, SUM(checkreceipt) AS MothaChecku from vouchers where (deleted = '0') AND (receiptDate <= " & "'" & Format(mTo, "mm/dd/yyyy") & "'" & ") and(paymode = " & "'" & Pmode & "'" & ") AND (CUSTNO = " & "'" & inv & "'" & ") GROUP BY CUSTNO"
    RsVoucher.Open Xvaw, CON1, adOpenDynamic, adLockOptimistic
    If RsVoucher.EOF = False Then
    ChkPayt = RsVoucher!MothaChecku
    End If





                'This will check all the INVOICE NUMBERS under the condition "If SJ.Recordset!cust_code = Trim(inv) And SJ.Recordset!invc_date < mTo Then"
                'and will put all the multy Invoice details in the FILTERINVOICE table
                SjCus = Trim(inv)
                SJ.Recordset.MoveFirst
              While SJ.Recordset.EOF = False   'Starts second Loop
              If SJ.Recordset!cust_code = Trim(inv) And Format(SJ.Recordset!invc_date, "dd/mm/yyyy") < Format(mTo, "dd/mm/yyyy") Then
                SJinvNo = SJ.Recordset!invc_no
                SjInvcDate = SJ.Recordset!invc_date
                SjInvAmt = SJ.Recordset!tot_amt
                SjUnpaid = SJ.Recordset!unpaidamt
                'xDDue = Format(SjInvcDate, "dd/mm/yyyy") + MTerms
                xDDue = DateAdd("d", MTerms, SjInvcDate)
                Dim Myval

                Myval = DateDiff("d", xDDue, mTo) 'this deducts the Date(As of Date) from Due Date to count
                '- which colomn said to be taken the Amout

                xCurrent = ""
                x30 = ""
                x60 = ""
                xover = ""

                If Myval < 30 Then
                xCurrent = SjUnpaid
                ElseIf Myval < 61 And Myval > 30 Then
                x30 = SjUnpaid
                ElseIf Myval < 91 And Myval > 60 Then
                x60 = SjUnpaid
                ElseIf Myval > 92 Then
                xover = SjUnpaid
                End If





'On Error Resume Next
'dr = "Select * from tempagaintsinvoice where invoiceno = " & "'" & SJinvNo & "'" & ""
'RStempagaintsinvoice.Open dr, con1, adOpenDynamic, adLockOptimistic
'If RStempagaintsinvoice.EOF = False Then
'RStempagaintsinvoice.MoveFirst
'Else
'Call FilteredInvoice
'End If
'On Error GoTo 0
'While RStempagaintsinvoice.EOF = False  'Starts Third Loop
't2CRno = RStempagaintsinvoice!Receiptno
'
''RsVoucher.Close

'
      Call FilteredInvoice


'
'RStempagaintsinvoice.MoveNext
'Wend                              'End of third Loop

End If
SJ.Recordset.MoveNext            'End of second Loop
Wend



With Temp
If Err.Number = 3704 Then
Temp.Open
End If

.AddNew
            Temp!cust_code = inv
            Temp!first_name = MFName
            Temp!last_name = MLName
            Temp!Address = MAdd1
            Temp!Tel_No = MTelP1
            Temp!Terms = MTerms
            Temp!openingBal = BBal
            Temp!CreditLimit = Mcrlimit
            Temp!CheckLimit = MchkLimit
            '!xfrom = Me.Mskfrom.Text
            !xTo = Me.MskTo.Text

                .Update
     End With
  Temp.Close
  FilteredInv.Close

  passdate CustLedger.Sections(2).Controls("label7")
  FrmReportGenarator.Command2.caption = "Delete Table"
  FrmReportGenarator.Command2.Width = 2055
  CustLedger.Show 1 'this is to show the Report

End Sub
Private Sub FilteredInvoice()
        With FilteredInv ' this will add the multi records of invoice for the printout for the particual customer
        .AddNew
         !cust_code = inv
         !Inum = SJinvNo
         !IDate = SjInvcDate
         !IAmt = SjInvAmt
         !unpaidamt = SjUnpaid
         !DateDue = xDDue
         On Error Resume Next
         !XCdays = IIf(IsNull(xCurrent), "0", (xCurrent))

         !x30days = x30
         !x60days = x60
         !x90days = x90
         !xover = xover
         On Error GoTo 0
         '!CRno = t2CRno
         If ChkPayt = "" Then
         Else
         !CheckAmt = ChkPayt
         End If

         '!CHKValueDate = DateDue
        '  !TotalChk_UP = Val(ChkPayt) + Val(SjUnpaid) 'this is to find the Total Balace
         .Update

        SjCus = ""
        SjInvAmt = ""
        SJinvNo = ""
        SjInvcDate = ""
        SjUnpaid = ""
        xDDue = ""
        xCurrent = ""
        x30 = ""
        x60 = ""
        x90 = ""
        xover = ""
        ChkPayt = ""
        CR_no = ""
        DateDue = ""
        End With


End Sub
Private Sub Combo1_Click()
CmbCode.Clear
If Combo1.Text = "Aging of Account Receivables" Or Combo1.Text = "Statement Of Account" Or Combo1.Text = "Debtors Outstanding Balances" Then
    Me.Label1.Visible = False
    Me.Label2.caption = "As of"
    Me.Mskfrom.Visible = False
    Me.MskTo.Top = 360
    Me.Label2.Top = 360
    Me.Frame3.Height = 800

    Me.CmbName.Clear
    On Error Resume Next
    RstCust.MoveFirst 'for the marcusfl
    While RstCust.EOF = False
    Me.CmbCode.AddItem RstCust!cust_code
    Me.CmbCode.AddItem RstCust!first_name
    RstCust.MoveNext
    Wend
    
    Dim rspayee As New ADODB.Recordset
    rspayee.Open "Select * from Payee", constring, adOpenDynamic, adLockOptimistic
    
    While rspayee.EOF = False
    Me.CmbCode.AddItem RstCust!cust_code
    Me.CmbCode.AddItem RstCust!first_name
    rspayee.MoveNext
    Wend
    
    
 On Error GoTo 0
ElseIf Combo1.Text = "Customer Ledger Account" Then
    Me.Label1.Visible = True
    Me.Label2.caption = "To"
    Me.Mskfrom.Visible = True
    Me.MskTo.Top = 720
    Me.Label2.Top = 720
    Me.Frame3.Height = 1215

    Me.CmbName.Clear
'    RstCust.MoveFirst 'for the marcusfl
'    While RstCust.EOF = False
'        Me.CmbCode.AddItem RstCust!cust_code
'        If IsNull(RstCust!FirsT_Name) <> True Then
'          Me.CmbName.AddItem RstCust.Fields(4)
'         Else
'         Me.CmbName.AddItem "Blank Name"
'        End If
'        RstCust.MoveNext
'    Wend
   With Me.Adodc1
     While Me.Adodc1.Recordset.EOF = False
        On Error Resume Next
        Me.CmbCode.AddItem Me.Adodc1.Recordset!cust_code
        Me.CmbName.AddItem Me.Adodc1.Recordset!first_name '& " " & Me.Adodc1.Recordset!LAst_Name
        On Error GoTo 0
        Me.Adodc1.Recordset.MoveNext
      Wend
    End With
'
'
'ElseIf Combo1.Text = "Aging of Account Payables" Then
'    rstVen.MoveFirst
'    While rstVen.EOF = False
'    CmbCode.AddItem rstVen!vencode
'    CmbName.AddItem rstVen!Venameeng
'    rstVen.MoveNext
'    Wend
End If
'
'Me.CmbName.clear


End Sub

Private Sub Command1_Click()
End Sub
Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Combo1.Text = "Aging of Account Receivables" Then
MskTo.SetFocus
End If
Combo1_Click
End If


End Sub
Private Sub Command2_Click()

Dim rsTe2 As New ADODB.Recordset
Dim filcus As New ADODB.Recordset

Dim filcus2 As New ADODB.Recordset


Dim rstTemporary3 As New ADODB.Recordset
Dim rstFilterdStatementOfAc As New ADODB.Recordset

Dim tempy As New ADODB.Recordset
Dim filtInv As New ADODB.Recordset

rstTemporary3.Open "select * from temporary3", CON1, adOpenDynamic, adLockOptimistic
rstFilterdStatementOfAc.Open "select * from FilterdStatementOfAc", CON1, adOpenDynamic, adLockOptimistic

rsTe2.Open "select * from temporary2", CON1, adOpenDynamic, adLockOptimistic
filcus.Open "select * from FilteredCustLedger", CON1, adOpenDynamic, adLockOptimistic
filcus2.Open "select * from FilteredCustLedger2", CON1, adOpenDynamic, adLockOptimistic

tempy.Open "select * from temporary", CON1, adOpenDynamic, adLockOptimistic
filtInv.Open "select * from Filteredinvoice", CON1, adOpenDynamic, adLockOptimistic


If Command2.caption = "&Exit" Then
Unload Me

ElseIf Command2.caption = "Delete Table" And Combo1.Text = "Aging of Account Receivables" Then
filtInv.Close
tempy.Close
filtInv.Open "delete from Filteredinvoice"
tempy.Open "delete from Temporary"
Command2.caption = "&Exit"


ElseIf Command2.caption = "Delete Table" And Combo1.Text = "Customer Ledger Account" Then
'On Error Resume Next
filcus.Close
rsTe2.Close
filcus2.Close
rsTe2.Open "delete from temporary2"
filcus.Open "delete from FilteredCustLedger"
filcus2.Open "delete from FilteredCustLedger2"

Command2.caption = "&Exit"


ElseIf Command2.caption = "Delete Table" And Combo1.Text = "Aging of Account Payables" Then
CredFilteredInv.Open
Credtemp.Open
On Error Resume Next
CredFilteredInv.Delete
Credtemp.Delete
On Error GoTo 0
CredFilteredInv.Close
Credtemp.Close
Command2.caption = "&Exit"


ElseIf Command2.caption = "Delete Table" And Combo1.Text = "Statement Of Account" Then
rstFilterdStatementOfAc.Close
rstTemporary3.Close
rstFilterdStatementOfAc.Open "delete from FilterdStatementOfAc"
rstTemporary3.Open "delete from temporary3"
Command2.caption = "&Exit"
End If

End Sub
Private Sub Form_Load()
Set CON1 = New ADODB.Connection
Set RstCust = New ADODB.Recordset
Set RstCust2 = New ADODB.Recordset
Set CredmVouc = New ADODB.Recordset 'For the Creditors Aging
Set CredmInside = New ADODB.Recordset 'For the Creditors Aging
Set rstVendor = New ADODB.Recordset 'For the Creditors Aging
Set CredRstCust2 = New ADODB.Recordset 'For the Creditors Aging
Set rstPaySetup = New ADODB.Recordset 'For the Creditors Aging
Set Credtemp = New ADODB.Recordset 'For the Creditors Aging
Set rstVen = New ADODB.Recordset
Set TempInv = New ADODB.Recordset
Set TempCre = New ADODB.Recordset
Set TempDeb = New ADODB.Recordset
Set TempBBal = New ADODB.Recordset
Set Temp = New ADODB.Recordset
Set temp2 = New ADODB.Recordset 'for the CustLedger
Set LedgVouc = New ADODB.Recordset 'for the CustLedger
Set ORmas = New ADODB.Recordset
Set FilteredInv = New ADODB.Recordset
Set CredFilterdInv = New ADODB.Recordset
Set mVouc = New ADODB.Recordset
Set mTemp2 = New ADODB.Recordset
Set mInside = New ADODB.Recordset
Set rstFinCr = New ADODB.Recordset
Set RstCustx = New ADODB.Recordset
Set RstCustY = New ADODB.Recordset
Set LMarkFl = New ADODB.Recordset
Set LfiltCL = New ADODB.Recordset
Set CredFilteredInv = New ADODB.Recordset
Set BeginBal = New ADODB.Recordset
Set CredMem = New ADODB.Recordset
Set DebMemo = New ADODB.Recordset
Set CredrstFinCr = New ADODB.Recordset

conStr1 = "Provider=MSDASQL;DSN=Finance;UID=sa; PWD=;"

'conStr2 = "for odbc"
CON1.Open conStr1
RstCust.Open "Select * from MARCUSFL ", CON1, adOpenDynamic, adLockOptimistic
rstVen.Open "Select * from VENDOR ", CON1, adOpenDynamic, adLockOptimistic
Temp.Open "select * from Temporary", CON1, adOpenDynamic, adLockOptimistic



Me.Combo1.AddItem "Aging of Account Receivables"
Me.Combo1.AddItem "Customer Ledger Account"
Me.Combo1.AddItem "Aging of Account Payables"
'Me.Combo1.AddItem "Debtors Outstanding Balances"
Me.Combo1.AddItem "Statement Of Account"

Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(Combo1.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub Mskfrom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
MskTo.SetFocus
End If
End Sub

Private Sub Mskfrom_LostFocus()
If Me.Mskfrom.Text = "__/__/____" Then
    Exit Sub
End If
Me.Mskfrom.Text = Format(Me.Mskfrom.Text, "dd/mm/yyyy")
cDay = Val(Left(Me.Mskfrom.Text, 2))
cMonth = Val(Mid(Me.Mskfrom.Text, 4, 2))
cYear = Val(Right(Me.Mskfrom.Text, 4))
If cDay > 31 Or cDay < 1 Then
    mess = MsgBox("Invalid Date", vbInformation + vbOKOnly, "Message")
    Me.Mskfrom.SetFocus
  ElseIf cMonth > 12 Or cMonth < 1 Then
    mess = MsgBox("Invalid Month", vbInformation + vbOKOnly, "Message")
    Me.Mskfrom.SetFocus
ElseIf cYear < 1900 Or cYear > Year(Date) Then
    mess = MsgBox("Invalid Year", vbInformation + vbOKOnly, "Message")
    Me.Mskfrom.SetFocus
End If

End Sub

Private Sub MskTo_KeyPress(KeyAscii As Integer)
If Frame1.Enabled = True Then
If KeyAscii = 13 Then
CmbCode.SetFocus
End If
End If
End Sub

Private Sub MskTo_LostFocus()
If Me.MskTo.Text = "__/__/____" Then
    Exit Sub
End If
Me.MskTo.Text = Format(Me.MskTo.Text, "dd/mm/yyyy")
cDay = Val(Left(Me.MskTo.Text, 2))
cMonth = Val(Mid(Me.MskTo.Text, 4, 2))
cYear = Val(Right(Me.MskTo.Text, 4))
If cDay > 31 Or cDay < 1 Then
    mess = MsgBox("Invalid Date", vbInformation + vbOKOnly, "Message")
    Me.MskTo.SetFocus
  ElseIf cMonth > 12 Or cMonth < 1 Then
    mess = MsgBox("Invalid Month", vbInformation + vbOKOnly, "Message")
    Me.MskTo.SetFocus
ElseIf cYear < 1900 Or cYear > Year(Date) Then
    mess = MsgBox("Invalid Year", vbInformation + vbOKOnly, "Message")
    Me.MskTo.SetFocus
End If

End Sub

Private Sub Option1_Click()
Me.Frame1.Enabled = False
Me.CmbCode.Text = ""
Me.CmbName.Text = ""
End Sub

Private Sub Option2_Click()
Me.Frame1.Enabled = True
End Sub

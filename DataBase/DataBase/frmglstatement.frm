VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmglstatement 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "General Ledger Balance  ÃÑÕÏÉ ÇáÇÓÊÇÐ ÇáÚÇã"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8055
   ControlBox      =   0   'False
   Icon            =   "frmglstatement.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   8055
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Left            =   1200
      Top             =   4920
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Group Printing ÇáØÈÇÚÉ ÈÇáãÌãæÚÇÊ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3960
      TabIndex        =   13
      Top             =   0
      Width           =   3735
   End
   Begin VB.Frame Frame2 
      Height          =   1935
      Left            =   120
      TabIndex        =   15
      Top             =   2880
      Visible         =   0   'False
      Width           =   7815
      Begin VB.CommandButton cmdprint1 
         Caption         =   "P&rint ØÈÇÚÉ"
         Height          =   375
         Left            =   5060
         TabIndex        =   34
         Top             =   1440
         Width           =   1335
      End
      Begin VB.ComboBox comsub 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   840
         Width           =   3255
      End
      Begin VB.ComboBox commain 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   480
         Width           =   2415
      End
      Begin VB.CommandButton cmdactivity1 
         Caption         =   "&Preview  ÚÑÖ"
         Height          =   375
         Left            =   3720
         TabIndex        =   17
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton cmdclose1 
         Caption         =   "&Close  ÇÛáÇÞ"
         Height          =   375
         Left            =   6380
         TabIndex        =   16
         Top             =   1440
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker dtpfrom1 
         Height          =   315
         Left            =   2400
         TabIndex        =   20
         Top             =   1200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   51970049
         CurrentDate     =   37578
      End
      Begin MSComCtl2.DTPicker dtpto1 
         Height          =   315
         Left            =   2400
         TabIndex        =   21
         Top             =   1560
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   51970049
         CurrentDate     =   37578
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Main Group ÇáãÌãæÚÉ ÇáÑÆíÓíÉ"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   26
         Top             =   480
         Width           =   2085
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sub Group ÇáãÌæÚÇÊ ÇáÝÑÚíÉ "
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   25
         Top             =   840
         Width           =   2025
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Çáí "
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   24
         Top             =   1560
         Width           =   525
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From ãä "
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   23
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Group Details ÊÝÇÕíá ÇáãÌãæÚÇÊ"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   2460
      End
   End
   Begin VB.OptionButton optall 
      Caption         =   "Series Printing ØÈÇÚÉ ãÊÓáÓáÉ "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   14
      Top             =   0
      Value           =   -1  'True
      Width           =   3735
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   300
      Width           =   7815
      Begin VB.CommandButton cmdprint 
         Caption         =   "P&rint ØÈÇÚÉ"
         Height          =   375
         Left            =   5040
         TabIndex        =   33
         Top             =   1440
         Width           =   1365
      End
      Begin VB.CommandButton cmdclose 
         Caption         =   "&Close ÇÛáÇÞ"
         Height          =   375
         Left            =   6360
         TabIndex        =   12
         Top             =   1440
         Width           =   1365
      End
      Begin VB.ComboBox comtoname 
         Height          =   315
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   840
         Width           =   4695
      End
      Begin VB.ComboBox comfromname 
         Height          =   315
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   480
         Width           =   4695
      End
      Begin VB.ComboBox comto 
         Height          =   315
         Left            =   1440
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   840
         Width           =   1575
      End
      Begin VB.ComboBox comfrom 
         Height          =   315
         Left            =   1440
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton cmdactivity 
         Caption         =   "&Preview ÚÑÖ"
         Height          =   375
         Left            =   3680
         TabIndex        =   2
         Top             =   1440
         Width           =   1365
      End
      Begin MSComCtl2.DTPicker dtpfrom 
         Height          =   315
         Left            =   1440
         TabIndex        =   3
         Top             =   1200
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Format          =   51970049
         CurrentDate     =   37530
      End
      Begin MSComCtl2.DTPicker dtpto 
         Height          =   315
         Left            =   1440
         TabIndex        =   4
         Top             =   1560
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Format          =   51970049
         CurrentDate     =   37621
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Numbers ÑÞã ÇáÍÓÇÈ "
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Çáí "
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   525
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From ãä "
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From ãä "
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Çáí "
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   1560
         Width           =   525
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   150
      Left            =   120
      TabIndex        =   27
      Top             =   2520
      Visible         =   0   'False
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   265
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.Label Label33 
      BackStyle       =   0  'Transparent
      Caption         =   "checking time"
      Height          =   255
      Left            =   1800
      TabIndex        =   32
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Label Label51 
      BackStyle       =   0  'Transparent
      Caption         =   "checkingtime"
      Height          =   255
      Left            =   3720
      TabIndex        =   31
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label lp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100 %"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7320
      TabIndex        =   29
      Top             =   2520
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Label lblstatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Choose The Level Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   28
      Top             =   2280
      Visible         =   0   'False
      Width           =   2250
   End
End
Attribute VB_Name = "frmglstatement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim closeme As String
Dim recfin As New ADODB.Recordset
Dim recglb As New ADODB.Recordset
Dim recope As New ADODB.Recordset
Dim recope1 As New ADODB.Recordset
Dim recinquery As New ADODB.Recordset
Dim con2 As New ADODB.Connection
Dim level As String
Dim activity As Integer
Dim summery As Integer

Private Sub prcgetdata()

If Trim(cLogUser) = "" Then
    cLogUser = "xx"
End If

'this is all to create the temptable
Dim reccreatetable As New ADODB.Recordset
On Error GoTo errcreatetable
reccreatetable.Open "Select * into " & cLogUser & "openingbalance from openingbalance", con2
reccreatetable.Open "Select * into " & cLogUser & "financemaster from FinanceMaster", con2
reccreatetable.Open "Select * into " & cLogUser & "glmaster from glmaster", con2

errcreatetable:
If Err.Number <> 0 Then
    If MsgBox("Your User Name Has Been Using In Another Workstation;" & vbCrLf & "Please LogOff And Try Again or Do You Want To Continue... ?", vbInformation + vbYesNo, "Multi Login") = vbNo Then
        closeme = 1
        Exit Sub
    End If
End If
On Error GoTo 0
'end create the temptable

On Error Resume Next
con2.Close
On Error GoTo 0
con2.Open "Dsn=Finance;Uid=Sa;Pwd=;" ' this is for refresh the connection

'to delete only that user details
Dim recdelete As New ADODB.Recordset
recdelete.Open "delete from glprinttable where loguser ='" & cLogUser & "'", con2
recinquery.Open "Select * from glprinttable", con2, adOpenDynamic, adLockOptimistic, adCmdText


recfin.Open "Select * from " & cLogUser & "FinanceMaster where active <> '0' and " & "Accountcode >= " & "'" & Trim(comfrom.Text) & "'" & " And accountcode <= " & "'" & Trim(comto.Text) & "'", con2, adOpenKeyset, adLockOptimistic
ProgressBar1.Min = 0
ProgressBar1.Max = recfin.RecordCount

ProgressBar1.Value = 0
Dim recchange As New ADODB.Recordset
recchange.Open "update " & cLogUser & "glmaster set printed='0'", con2, adOpenKeyset, adLockPessimistic
recchange.Open "update " & cLogUser & "openingbalance set printed = '0'", con2, adOpenKeyset, adLockOptimistic

'end clear
            Dim recglbdebitamount As Currency
            Dim recglbcreditamount As Currency
            Dim recopebeginningDebit As Currency
            Dim recopebeginningCredit As Currency
            'this three for update the printed colum in those table
            Dim recupglb As New ADODB.Recordset
            Dim recupope1 As New ADODB.Recordset
            Dim recupope As New ADODB.Recordset
               
If recfin.BOF = False Then
recfin.MoveFirst
End If
While recfin.EOF = False
recfinaccountcode = recfin!AccountCode
'this is all for GLMaster table
        recglb.Open "SELECT AccountCode, SUM(CreditAmount) AS creditamount, SUM(DebitAmount) AS debitamount, CASE WHEN SUM(DebitAmount) >= SUM(creditamount) tHEN SUM(debitamount) - SUM(creditamount)else 0 END AS endingDebit, CASE WHEN SUM(creditamount) >= SUM(DebitAmount) THEN SUM(creditamount) - SUM(debitamount) Else 0 END AS endingCredit from " & cLogUser & "GLMaster where accountcode = " & "'" & recfinaccountcode & "'" & " and recorddate >= " & "'" & Format(dtpfrom.Value, "mm/dd/yyyy") & "'" & " and recorddate <= " & "'" & Format(dtpto.Value, "mm/dd/yyyy") & "'" & "GROUP BY AccountCode", con2, adOpenKeyset, adLockOptimistic
        recope1.Open "SELECT AccountCode, SUM(CreditAmount) AS creditamount, SUM(DebitAmount) AS debitamount from " & cLogUser & "GLMaster where accountcode = " & "'" & recfinaccountcode & "'" & " And recorddate < " & "'" & Format(dtpfrom.Value, "mm/dd/yyyy") & "'" & " GROUP BY AccountCode", con2, adOpenKeyset, adLockOptimistic
        recope.Open "select sum(beginningdebit) as beginningdebit,sum(beginningcredit) as beginningcredit from " & cLogUser & "OpeningBalance where accountcode = " & "'" & recfinaccountcode & "'", con2, adOpenKeyset, adLockOptimistic
        
        'this is for update printed colum on glmaster table
        recupglb.Open "update " & cLogUser & "GLMaster  set printed = '1' where recorddate >= " & "'" & Format(dtpfrom.Value, "mm/dd/yyyy") & "'" & " and recorddate <= " & "'" & Format(dtpto.Value, "mm/dd/yyyy") & "'", con2, adOpenKeyset, adLockOptimistic
        recupope1.Open "update " & cLogUser & "GLMaster  set printed = '1' where recorddate < " & "'" & Format(dtpfrom.Value, "mm/dd/yyyy") & "'", con2, adOpenKeyset, adLockOptimistic
        recupope.Open "update " & cLogUser & "openingbalance set printed ='1' where accountcode = " & "'" & recfinaccountcode & "'", con2, adOpenKeyset, adLockOptimistic
        
        With recinquery
            .AddNew
            lblstatus.Visible = True
            lblstatus.caption = "Adding Level Base Account Numbers...."
            ProgressBar1.Visible = True
            !LogUser = cLogUser
            !AccountCode = recfin!AccountCode
            !accountname = recfin!accountnameeng
            !accountnameara = recfin!accountnamearab
            'take the begin balance from glmaster upto specific date
            If recope1.BOF = False Then
                recopebeginningDebit11 = Val(IIf(IsNull(recope1!DebitAmount), 0, recope1!DebitAmount))
                recopebeginningCredit11 = Val(IIf(IsNull(recope1!creditamount), 0, recope1!creditamount))
            Else
                recopebeginningDebit11 = 0
                recopebeginningCredit11 = 0
            End If
 
            ' take the begin balance from opening balance table
            If recope.BOF = False Then
                recopebeginningDebit = Val(IIf(IsNull(recope!beginningdebit), 0, recope!beginningdebit)) + Val(recopebeginningDebit11)
                recopebeginningCredit = Val(IIf(IsNull(recope!beginningcredit), 0, recope!beginningcredit)) + Val(recopebeginningCredit11)
            Else
                recopebeginningDebit = 0 + Val(recopebeginningDebit11)
                recopebeginningCredit = 0 + Val(recopebeginningCredit11)
            End If
            
            'check for whether debit more or credit more
            If recopebeginningDebit >= recopebeginningCredit Then
                recopebeginningDebit = recopebeginningDebit - recopebeginningCredit
                recopebeginningCredit = 0
            Else
                recopebeginningCredit = recopebeginningCredit - recopebeginningDebit
                recopebeginningDebit = 0
            End If
            
                !beginningdebit = recopebeginningDebit
                !beginningcredit = recopebeginningCredit

            If recglb.BOF = False Then
                !activitydebit = recglb!DebitAmount
                !ActivityCredit = recglb!creditamount
                recglbdebitamount = IIf(IsNull(recglb!DebitAmount), 0, recglb!DebitAmount)
                recglbcreditamount = IIf(IsNull(recglb!creditamount), 0, recglb!creditamount)

            Else
                !activitydebit = 0
                !ActivityCredit = 0
                recglbdebitamount = 0
                recglbcreditamount = 0
            End If
            
            ed = Val(recglbdebitamount) + Val(recopebeginningDebit) ' Ending Debit
            ec = Val(recglbcreditamount) + Val(recopebeginningCredit) ' Ending Credit
            If ed >= ec Then
                ed = ed - ec
                ec = 0
            Else
                ec = ec - ed
                ed = 0
            End If
            
            !endingDebit = ed
            !endingCredit = ec
            .Update
        End With
    DoEvents
    ProgressBar1.Value = ProgressBar1.Value + 1
    lp.caption = Format(((ProgressBar1.Value * 100) / ProgressBar1.Max), "###") & " %"

    recglb.Close
    recope.Close
    recope1.Close
recfin.MoveNext
Wend
ProgressBar1.Visible = False
lblstatus.Visible = False
lp.Visible = False
recinquery.Requery
recfin.Close

'this is drop table
Dim recdroptable As New ADODB.Recordset
'On Error GoTo ar
recdroptable.Open "Drop table " & cLogUser & "openingbalance", con2
recdroptable.Open "Drop table " & cLogUser & "financemaster", con2
recdroptable.Open "Drop table " & cLogUser & "glmaster", con2
'ar:
'MsgBox Err.Description

On Error Resume Next
con2.Close
On Error GoTo 0
con2.Open "Dsn=Finance;Uid=Sa;Pwd=;" ' this is for refresh the connection

End Sub
Private Sub passlevel(a As RptLabel)
If level = "All" Then
    level = "All Main Accounts"
ElseIf level = "2" Then
    level = "2nd Level"
ElseIf level = "3" Then
    level = "3rd Level"
ElseIf level = "4" Then
    level = "4th Level"
ElseIf level = "5" Then
    level = "5th Level"
ElseIf level = "6" Then
    level = "6th Level"
End If

a.caption = "Level :  " & level
End Sub

Private Sub passlevel1(a As RptLabel)
a.caption = " From :  " & Format(dtpfrom.Value, "dd/mm/yyyy") & "  To :  " & Format(dtpto.Value, "dd/mm/yyyy")
End Sub

Private Sub cmdactivity_Click()
activity = 1
'this is to check the account number whether this is correct
Dim checkcomfrom As Double
Dim checkcomto As Double

checkcomfrom = Val(Trim(comfrom.Text))
checkcomto = Val(Trim(comto.Text))

If checkcomfrom <= 0 Or checkcomto <= 0 Then
    MsgBox "Please check the Account Number From The List", vbInformation, "Invalid Account Number"
    comfrom.SetFocus
    Exit Sub
End If

If checkcomfrom > checkcomto Then
    MsgBox "Please check Accont Number That is Incorrect Format", vbInformation, "In Order Numbers"
    Exit Sub
End If
'errfromto:


If dtpfrom.Value > dtpto.Value Then
    MsgBox "Please Enter Your Date Correctly", vbInformation, "Disorder Date"
    Exit Sub
End If

DoEvents
cmdactivity.Enabled = False
cmdclose.Enabled = False
cmdprint.Enabled = False

DoEvents
addheight Me ' this is for add the form height
cmdactivity.Enabled = False
lp.Visible = True

Call prcgetdata

If closeme = "1" Then
    closeme = 2
    lp.Visible = False
    cmdactivity.Enabled = True
    cmdclose.Enabled = True
    cmdprint.Enabled = True
    lessheight Me
    Exit Sub
End If

lessheight Me
cmdPrint_Click
lp.Visible = False
cmdactivity.Enabled = True
cmdclose.Enabled = True
cmdprint.Enabled = True
End Sub

Private Sub cmdsummery_Click()
activity = 2
If Trim(comlevel.Text) = "" Then
    MsgBox "Please check The Level That You Select", vbInformation, "Empty Level"
    Exit Sub
End If
lp.Visible = True
level = Trim(comlevel.Text)
       passlevel re_TrialBalance_summerybased.Sections(2).Controls("Label15")
       passlevel1 re_TrialBalance_summerybased.Sections(2).Controls("Label3") ' to pass the from to date
activity = 6
On Error Resume Next
dataanu.rstrial_balance_summery.Requery
On Error GoTo 0
re_TrialBalance_summerybased.Show 1
cmdsummery.Enabled = False
lp.Visible = False
End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdclose1_Click()
Unload Me
End Sub

Private Sub cmdactivity1_Click()
activity = 1
'this is to check the account number whether this is correct
Dim checkcomfrom As Double
Dim checkcomto As Double

If Trim(commain.Text) = "" Or Trim(comsub.Text) = "" Then
    MsgBox "please Choose the Main Code & Sub Code Correctly", vbInformation, "Invalid Codes or Names"
    Exit Sub
End If

If dtpfrom1.Value > dtpto1.Value Then
    MsgBox "Please Enter Your Date Correctly", vbInformation, "Disorder Date"
    Exit Sub
End If

DoEvents
cmdactivity1.Enabled = False
cmdclose1.Enabled = False
cmdprint1.Enabled = False

DoEvents
addheight Me ' this is for add the form height
cmdactivity1.Enabled = False
lp.Visible = True

Call prcgetdata1
If closeme = "1" Then
    closeme = 2
    lp.Visible = False
    cmdactivity1.Enabled = True
    cmdclose1.Enabled = True
    cmdprint1.Enabled = True
    lessheight Me
    Exit Sub
End If

lessheight Me
cmdprint1_Click
lp.Visible = False
cmdactivity1.Enabled = True
cmdclose1.Enabled = True
cmdprint1.Enabled = True
End Sub

Private Sub cmdPrint_Click()

On Error Resume Next
dataanu.rscom_gl_statement.Close
On Error GoTo 0

dataanu.com_gl_statement cLogUser

If MsgBox("Select the Language to Print      ÇÎÊÇÑ áÛÉ ÇáØÈÇÚÉ  " & vbCrLf & "Yes :- For Arabic                                       ÈÇáÚÑÈí " & vbCrLf & "No  :- For English                                  ÈÇáÇäÌáíÒí", vbInformation + vbYesNo, "Language Option  ÇÎÊíÇÑ ÇááÛÉ") = vbYes Then
    lbldateara re_ara_glstatement.Sections(2).Controls("lbldate")
    re_ara_glstatement.Show 1
Else
    lbldate re_eng_glstatement.Sections(2).Controls("lbldate")
    re_eng_glstatement.Show 1
End If

End Sub

Private Sub cmdprint1_Click()
On Error Resume Next
dataanu.rscom_gl_statement.Close
On Error GoTo 0

dataanu.com_gl_statement cLogUser

If MsgBox("Select the Language to Print      ÇÎÊÇÑ áÛÉ ÇáØÈÇÚÉ  " & vbCrLf & "Yes :- For Arabic                                       ÈÇáÚÑÈí " & vbCrLf & "No  :- For English                                  ÈÇáÇäÌáíÒí", vbInformation + vbYesNo, "Language Option  ÇÎÊíÇÑ ÇááÛÉ") = vbYes Then
    lbldateara re_ara_glstatement.Sections(2).Controls("lbldate")
    re_ara_glstatement.Show 1
Else
    lbldate re_eng_glstatement.Sections(2).Controls("lbldate")
    re_eng_glstatement.Show 1
End If


End Sub

Private Sub lbldate(a As RptLabel)
a.caption = "As End Of :  " & Format(dtpto.Value, "dd/mm/yyyy")
End Sub
Private Sub lbldateara(a As RptLabel)
Dim dfrom As Date
Dim dto As Date
allword = "Ýí äåÇíÉ íæã  " & DatePart("yyyy", Format(dtpto.Value, "dd/mm/yyyy")) & "/" & DatePart("m", Format(dtpto.Value, "dd/mm/yyyy")) & "/" & DatePart("d", Format(dtpto.Value, "dd/mm/yyyy"))
'allword = "  ãä  ÊÇÑíÎ   " & Format(dtpfrom.Value, "dd/mm/yyyy") & "  áÛÇíÉ  " & Format(dtpto.Value, "dd/mm/yyyy")
a.caption = allword
End Sub


Private Sub comfrom_Click()
'this is for find the account name
Dim recfindaccount As New ADODB.Recordset
recfindaccount.Open "select * from financemaster where active <> '0' and accountcode = " & "'" & Trim(comfrom.Text) & "'", con2, adOpenKeyset, adLockOptimistic
If recfindaccount.BOF = False Then
    comfromname.Text = recfindaccount!accountnamearab & "\" & recfindaccount!accountnameeng
End If
recfindaccount.Close
'end find the account name
    

End Sub

Private Sub comfrom_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyPageDown
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(comfrom.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Select

End Sub

Private Sub comfrom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    comfromname.SetFocus
End If
End Sub

Private Sub comfrom_LostFocus()
'this is for find the account name
If Trim(comfrom.Text) <> "" Then
Dim recfindaccount As New ADODB.Recordset
recfindaccount.Open "select * from financemaster where active <> '0' and accountcode = " & "'" & Trim(comfrom.Text) & "'", con2, adOpenKeyset, adLockOptimistic
If recfindaccount.BOF = False Then
    comfromname.Text = recfindaccount!accountnamearab & "\" & recfindaccount!accountnameeng
    comto.Text = Trim(comfrom.Text)
Else
    comfrom.Text = "  "
    comfrom.SetFocus
    Exit Sub
End If
recfindaccount.Close
End If
'end find the account name

End Sub

Private Sub comfromname_Click()
' to change the account code
Dim recfindacc As New ADODB.Recordset
namenamenum = InStr(1, Trim(comfromname.Text), "\", vbTextCompare)
If namenamenum > 0 Then
namename = Mid(Trim(comfromname.Text), (namenamenum + 1), Len(Trim(comfromname.Text)))
Else
namename = Trim(comfromname.Text)
End If
recfindacc.Open "Select * from financemaster where active <> '0' and accountnameeng = " & "'" & namename & "'", con2, adOpenKeyset, adLockOptimistic
If recfindacc.BOF = False Then
    comfrom.Text = recfindacc!AccountCode
End If
 recfindacc.Close
' end change the account number

End Sub

Private Sub comfromname_GotFocus()
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(comfromname.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)

End Sub

Private Sub comfromname_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyPageDown
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(comfromname.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Select


End Sub

Private Sub comfromname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    comto.SetFocus
End If
End Sub

Private Sub comfromname_LostFocus()
' to change the account code
Dim recfindacc As New ADODB.Recordset
namenamenum = InStr(1, Trim(comfromname.Text), "\", vbTextCompare)
If namenamenum > 0 Then
namename = Mid(Trim(comfromname.Text), (namenamenum + 1), Len(Trim(comfromname.Text)))
Else
namename = Trim(comfromname.Text)
End If
recfindacc.Open "Select * from financemaster where active <> '0' and accountnameeng = " & "'" & namename & "'", con2, adOpenKeyset, adLockOptimistic
If recfindacc.BOF = False Then
    comfrom.Text = recfindacc!AccountCode
Else
    MsgBox "Please Choose Correct Account Name", vbInformation, "Invalid Name"
    comfromname.SetFocus
    Exit Sub
End If
 recfindacc.Close
' end change the account number

End Sub

Private Sub commain_Click()
comsub.Clear

Dim recsub As New ADODB.Recordset
recsub.Open "select * from groupname where groupcatcode = '" & Trim(Mid(Trim(commain.Text), 1, 4)) & "'", con2, adOpenKeyset, adLockOptimistic
    If recsub.BOF = False Then
        If recsub.RecordCount > 1 Then
            comsub.AddItem "0000     All"
        End If
        While recsub.EOF = False
            comsub.AddItem recsub!Idno & "     " & recsub!GroupNameEng
            recsub.MoveNext
        Wend
    End If
recsub.Close

End Sub

Private Sub comto_Click()
'this is for find the account name
Dim recfindaccount As New ADODB.Recordset
recfindaccount.Open "select * from financemaster where active <> '0' and accountcode = " & "'" & Trim(comto.Text) & "'", con2, adOpenKeyset, adLockOptimistic
If recfindaccount.BOF = False Then
    comtoname.Text = recfindaccount!accountnamearab & "\" & recfindaccount!accountnameeng
End If
recfindaccount.Close
'end find the account name
    

End Sub

Private Sub comto_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyPageDown
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(comto.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Select

End Sub

Private Sub comto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    comtoname.SetFocus
End If
End Sub

Private Sub comto_LostFocus()
'this is for find the account name
If Trim(comfrom.Text) <> "" Then
Dim recfindaccount As New ADODB.Recordset
recfindaccount.Open "select * from financemaster where active <> '0' and accountcode = " & "'" & Trim(comto.Text) & "'", con2, adOpenKeyset, adLockOptimistic
If recfindaccount.BOF = False Then
    comtoname.Text = recfindaccount!accountnamearab & "\" & recfindaccount!accountnameeng
Else
comto.Text = " "
Exit Sub
End If
recfindaccount.Close
End If
'end find the account name

End Sub

Private Sub comtoname_Click()
' to change the account code
Dim recfindacc As New ADODB.Recordset
namenamenum = InStr(1, Trim(comtoname.Text), "\", vbTextCompare)
If namenamenum > 0 Then
namename = Mid(Trim(comtoname.Text), (namenamenum + 1), Len(Trim(comtoname.Text)))
Else
namename = Trim(comtoname.Text)
End If
recfindacc.Open "Select * from financemaster where active <> '0' and accountnameeng = " & "'" & namename & "'", con2, adOpenKeyset, adLockOptimistic
If recfindacc.BOF = False Then
    comto.Text = recfindacc!AccountCode
End If
 recfindacc.Close
' end change the account number

End Sub

Private Sub comtoname_GotFocus()
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(comtoname.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)

End Sub

Private Sub comtoname_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyPageDown
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(comtoname.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Select


End Sub

Private Sub comtoname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdactivity.SetFocus
End If
End Sub

Private Sub comtoname_LostFocus()
' to change the account code
Dim recfindacc As New ADODB.Recordset
namenamenum = InStr(1, Trim(comtoname.Text), "\", vbTextCompare)
If namenamenum > 0 Then
namename = Mid(Trim(comtoname.Text), (namenamenum + 1), Len(Trim(comtoname.Text)))
Else
namename = Trim(comtoname.Text)
End If
recfindacc.Open "Select * from financemaster where active <> '0' and accountnameeng = " & "'" & namename & "'", con2, adOpenKeyset, adLockOptimistic
If recfindacc.BOF = False Then
    comto.Text = recfindacc!AccountCode
Else
    MsgBox "Please Choose Correct Account Name", vbInformation, "Invalid Name"
    comtoname.SetFocus
    Exit Sub
End If
 recfindacc.Close
' end change the account number


End Sub

Private Sub optall_Click()
If optall.Value = True Then
    Frame2.Visible = False
    Frame1.Visible = True
    comfrom.Text = ""
    comfromname.Text = ""
    comto.Text = ""
    comtoname.Text = ""
End If
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
    Frame1.Visible = False
    Frame2.Visible = True
'    commain.Text = ""  ' becasue this is read only property
'    comsub.Text = ""
End If

End Sub

Private Sub Form_Load()
Label33.caption = Time
Timer2.Interval = 1
Mainform.sbStatusBar.Panels(1).Text = "Status : Account Inquery ..."
con2.Open "Dsn=Finance;Uid=Sa;Pwd=;"

Dim recaddfin As New ADODB.Recordset
Dim constring As String
Dim myclass As New HabitatClass
Dim sqltable As Boolean
Dim xtable As String
Dim conrecfin As New ADODB.Connection

constring = "Dsn=finance;Uid=sa;Pwd=;"
xtable = "SELECT * from financemaster where active <> '0' order by accountcode"
sqltable = True
myclass.GetTables recaddfin, conrecfin, xtable, constring, sqltable

comfrom.Clear
comto.Clear
While recaddfin.EOF = False
    comfrom.AddItem recaddfin!AccountCode
    comto.AddItem recaddfin!AccountCode
    anu = recaddfin!accountnamearab & "\" & recaddfin!accountnameeng
    comfromname.AddItem anu
    comtoname.AddItem anu
    recaddfin.MoveNext
Wend
recaddfin.Close
conrecfin.Close


dtpto.Value = Date
dtpto1.Value = Date
Dim fromdate As Date
datefromdate = "01/01/" & DatePart("yyyy", Date)
fromdate = Format(datefromdate, "mm/dd/yyyy")
dtpfrom.Value = fromdate
dtpfrom1.Value = fromdate
Mainform.sbStatusBar.Panels(1).Text = "Status : Ready."

'this is for grouping
Dim recmain As New ADODB.Recordset
recmain.Open "select * from groupcat", con2, adOpenDynamic, adLockOptimistic, adCmdText

    If recmain.BOF = False Then
        While recmain.EOF = False
            commain.AddItem Trim(recmain!Code) & "     " & Trim(recmain!NameEng)
            recmain.MoveNext
        Wend
    End If
recmain.Close
Frame2.Top = Frame1.Top
Frame2.Left = Frame1.Left
lessheight Me
Timer2.Interval = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
con2.Close
End Sub

Sub addheight(anuForm As Form)
    anuForm.Height = 3270
End Sub
Sub lessheight(anuForm As Form)
    anuForm.Height = 2760
End Sub
Private Sub prcgetdata1()

If Trim(cLogUser) = "" Then
    cLogUser = "xx"
End If

Dim reccreatetable As New ADODB.Recordset
On Error GoTo errcreatetable
reccreatetable.Open "Select * into " & cLogUser & "openingbalance from openingbalance", con2
reccreatetable.Open "Select * into " & cLogUser & "financemaster from FinanceMaster", con2
reccreatetable.Open "Select * into " & cLogUser & "glmaster from glmaster", con2

errcreatetable:
If Err.Number <> 0 Then
    If MsgBox("Your User Name Has Been Using In Another Workstation;" & vbCrLf & "Please LogOff And Try Again or Do You Want To Continue... ?", vbInformation + vbYesNo, "Multi Login") = vbNo Then
        closeme = 1
        Exit Sub
    End If
End If
On Error GoTo 0

'end create the temptable
On Error Resume Next
con2.Close
On Error GoTo 0
con2.Open "Dsn=Finance;Uid=Sa;Pwd=;" ' this is for refresh the connection


If Trim(Mid(Trim(comsub.Text), 1, 7)) <> "0000" Then
'if it is not zero then should opent only the specific account numbers
recfin.Open "SELECT * from groupmember where groupcatcode = '" & Trim(Mid(Trim(commain.Text), 1, 4)) & "' and  groupnamecode = '" & Trim(Mid(Trim(comsub.Text), 1, 7)) & "'", con2, adOpenKeyset, adLockOptimistic
Else
'if it is zero it should open all codes
recfin.Open "SELECT * from groupmember where groupcatcode = '" & Trim(Mid(Trim(commain.Text), 1, 4)) & "'", con2, adOpenKeyset, adLockOptimistic
End If

'MsgBox recfin.RecordCount

ProgressBar1.Min = 0
ProgressBar1.Max = recfin.RecordCount

ProgressBar1.Value = 0
' this is for clear the glprinttable table
recinquery.Open "delete from glprinttable where loguser ='" & cLogUser & "'", con2, adOpenKeyset, adLockOptimistic
recinquery.Open "select * from glprinttable", con2, adOpenKeyset, adLockOptimistic

'end clear

Dim recchange As New ADODB.Recordset
recchange.Open "update " & cLogUser & "glmaster set printed='0'", con2, adOpenKeyset, adLockPessimistic
recchange.Open "update " & cLogUser & "openingbalance set printed = '0'", con2, adOpenKeyset, adLockOptimistic

            Dim recglbdebitamount As Currency
            Dim recglbcreditamount As Currency
            Dim recopebeginningDebit As Currency
            Dim recopebeginningCredit As Currency
            
            'this three for update the printed colum in those table
            Dim recupglb As New ADODB.Recordset
            Dim recupope1 As New ADODB.Recordset
            Dim recupope As New ADODB.Recordset
               
If recfin.BOF = False Then
recfin.MoveFirst
End If
While recfin.EOF = False
recfinaccountcode = recfin!memberAcctCode
'this is all for GLMaster table
        recglb.Open "SELECT AccountCode, SUM(CreditAmount) AS creditamount, SUM(DebitAmount) AS debitamount, CASE WHEN SUM(DebitAmount) >= SUM(creditamount) tHEN SUM(debitamount) - SUM(creditamount)else 0 END AS endingDebit, CASE WHEN SUM(creditamount) >= SUM(DebitAmount) THEN SUM(creditamount) - SUM(debitamount) Else 0 END AS endingCredit from " & cLogUser & "GLMaster where accountcode = " & "'" & recfinaccountcode & "'" & " and recorddate >= " & "'" & Format(dtpfrom.Value, "mm/dd/yyyy") & "'" & " and recorddate <= " & "'" & Format(dtpto.Value, "mm/dd/yyyy") & "'" & "GROUP BY AccountCode", con2, adOpenKeyset, adLockOptimistic
        recope1.Open "SELECT AccountCode, SUM(CreditAmount) AS creditamount, SUM(DebitAmount) AS debitamount from " & cLogUser & "GLMaster where accountcode = " & "'" & recfinaccountcode & "'" & " And recorddate < " & "'" & Format(dtpfrom.Value, "mm/dd/yyyy") & "'" & " GROUP BY AccountCode", con2, adOpenKeyset, adLockOptimistic
        recope.Open "select sum(beginningdebit) as beginningdebit,sum(beginningcredit) as beginningcredit from " & cLogUser & "OpeningBalance where accountcode = " & "'" & recfinaccountcode & "'", con2, adOpenKeyset, adLockOptimistic
        
        'this is for update printed colum on glmaster table
        recupglb.Open "update " & cLogUser & "GLMaster  set printed = '1' where recorddate >= " & "'" & Format(dtpfrom.Value, "mm/dd/yyyy") & "'" & " and recorddate <= " & "'" & Format(dtpto.Value, "mm/dd/yyyy") & "'", con2, adOpenKeyset, adLockOptimistic
        recupope1.Open "update " & cLogUser & "GLMaster  set printed = '1' where recorddate < " & "'" & Format(dtpfrom.Value, "mm/dd/yyyy") & "'", con2, adOpenKeyset, adLockOptimistic
        recupope.Open "update " & cLogUser & "openingbalance set printed ='1' where accountcode = " & "'" & recfinaccountcode & "'", con2, adOpenKeyset, adLockOptimistic
        
        With recinquery
            .AddNew
            lblstatus.Visible = True
            lblstatus.caption = "Adding Level Base Account Numbers...."
            ProgressBar1.Visible = True
            !LogUser = cLogUser
            !AccountCode = recfin!memberAcctCode
            !accountname = recfin!memberNameEng
            !accountnameara = recfin!memberNameArab
            'take the begin balance from glmaster upto specific date
            If recope1.BOF = False Then
                recopebeginningDebit11 = Val(IIf(IsNull(recope1!DebitAmount), 0, recope1!DebitAmount))
                recopebeginningCredit11 = Val(IIf(IsNull(recope1!creditamount), 0, recope1!creditamount))
            Else
                recopebeginningDebit11 = 0
                recopebeginningCredit11 = 0
            End If
 
            ' take the begin balance from opening balance table
            If recope.BOF = False Then
                recopebeginningDebit = Val(IIf(IsNull(recope!beginningdebit), 0, recope!beginningdebit)) + Val(recopebeginningDebit11)
                recopebeginningCredit = Val(IIf(IsNull(recope!beginningcredit), 0, recope!beginningcredit)) + Val(recopebeginningCredit11)
            Else
                recopebeginningDebit = 0 + Val(recopebeginningDebit11)
                recopebeginningCredit = 0 + Val(recopebeginningCredit11)
            End If
            
            'check for whether debit more or credit more
            If recopebeginningDebit >= recopebeginningCredit Then
                recopebeginningDebit = recopebeginningDebit - recopebeginningCredit
                recopebeginningCredit = 0
            Else
                recopebeginningCredit = recopebeginningCredit - recopebeginningDebit
                recopebeginningDebit = 0
            End If
            
                !beginningdebit = recopebeginningDebit
                !beginningcredit = recopebeginningCredit

            If recglb.BOF = False Then
                !activitydebit = recglb!DebitAmount
                !ActivityCredit = recglb!creditamount
                recglbdebitamount = IIf(IsNull(recglb!DebitAmount), 0, recglb!DebitAmount)
                recglbcreditamount = IIf(IsNull(recglb!creditamount), 0, recglb!creditamount)

            Else
                !activitydebit = 0
                !ActivityCredit = 0
                recglbdebitamount = 0
                recglbcreditamount = 0
            End If
            
            ed = Val(recglbdebitamount) + Val(recopebeginningDebit) ' Ending Debit
            ec = Val(recglbcreditamount) + Val(recopebeginningCredit) ' Ending Credit
            If ed >= ec Then
                ed = ed - ec
                ec = 0
            Else
                ec = ec - ed
                ed = 0
            End If
            
            !endingDebit = ed
            !endingCredit = ec
            .Update
            
        End With
    DoEvents
    ProgressBar1.Value = ProgressBar1.Value + 1
    lp.caption = Format(((ProgressBar1.Value * 100) / ProgressBar1.Max), "###") & " %"

    recglb.Close
    recope.Close
    recope1.Close
recfin.MoveNext
Wend
ProgressBar1.Value = 0

'**************** end opening balance adding
ProgressBar1.Visible = False
lblstatus.Visible = False
lp.Visible = False
recinquery.Requery
recfin.Close

'this is drop table
Dim recdroptable As New ADODB.Recordset
'On Error GoTo ar
recdroptable.Open "Drop table " & cLogUser & "openingbalance", con2
recdroptable.Open "Drop table " & cLogUser & "financemaster", con2
recdroptable.Open "Drop table " & cLogUser & "glmaster", con2
'ar:
'MsgBox Err.Description

On Error Resume Next
con2.Close
On Error GoTo 0
con2.Open "Dsn=Finance;Uid=Sa;Pwd=;" ' this is for refresh the connection

End Sub


Private Sub Timer2_Timer()
Label51.caption = Time
End Sub

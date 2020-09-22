VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmtrialbalance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trial Balance ãíÒÇä ÇáãÑÇÌÚÉ"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6840
   ControlBox      =   0   'False
   Icon            =   "frmtrialbalance.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   6840
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      Begin VB.ComboBox comlevel 
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
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   480
         Width           =   735
      End
      Begin MSComCtl2.DTPicker dtpfrom 
         Height          =   315
         Left            =   3120
         TabIndex        =   2
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   50659329
         CurrentDate     =   37530
      End
      Begin MSComCtl2.DTPicker dtpto 
         Height          =   315
         Left            =   5160
         TabIndex        =   3
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   50659329
         CurrentDate     =   37621
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   1080
         Visible         =   0   'False
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   344
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.Frame Frame4 
         Caption         =   " Printing Option"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   6375
         Begin VB.CommandButton cmdactivity 
            Caption         =   "&Preview"
            Height          =   375
            Left            =   120
            TabIndex        =   16
            Top             =   480
            Width           =   975
         End
         Begin VB.CommandButton cmdprintactivity 
            Caption         =   "Print"
            Enabled         =   0   'False
            Height          =   375
            Left            =   1200
            TabIndex        =   15
            Top             =   480
            Width           =   975
         End
         Begin VB.CommandButton cmdsummery 
            Caption         =   "Pre&view"
            Enabled         =   0   'False
            Height          =   375
            Left            =   2760
            TabIndex        =   13
            Top             =   480
            Width           =   975
         End
         Begin VB.CommandButton cmdclose 
            Caption         =   "&Close"
            Height          =   375
            Left            =   5280
            TabIndex        =   11
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Activity Based     ÇáÇÓÇÓ ÇáÝÚÇáíÉ "
            BeginProperty Font 
               Name            =   "Arabic Transparent"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   2460
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Summary Based ÇáÇÓÇÓ ÇáÕíÝí"
            BeginProperty Font 
               Name            =   "Arabic Transparent"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   2760
            TabIndex        =   14
            Top             =   240
            Width           =   2415
         End
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
         Left            =   6000
         TabIndex        =   12
         Top             =   840
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
         TabIndex        =   6
         Top             =   840
         Visible         =   0   'False
         Width           =   2250
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Choose The Period ÇÎÊíÇÑ ÇáÝÊÑÉ "
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
         Left            =   2160
         TabIndex        =   9
         Top             =   120
         Width           =   2775
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Çáí "
         Height          =   195
         Left            =   4560
         TabIndex        =   8
         Top             =   480
         Width           =   525
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From ãä "
         Height          =   195
         Left            =   2400
         TabIndex        =   7
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Choose The Level"
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
         TabIndex        =   5
         Top             =   120
         Width           =   1560
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Level  ÇáãÓÊæí "
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmtrialbalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim closeme As String
Dim recfin As New ADODB.Recordset
Dim recglb As New ADODB.Recordset
Dim recope As New ADODB.Recordset
Dim recope1 As New ADODB.Recordset
Dim rectri As New ADODB.Recordset
Dim con2 As New ADODB.Connection
Dim level As String
Dim activity As Integer
Dim summery As Integer

Private Sub prcgetdata()

If Trim(cLogUser) = "" Then
    cLogUser = "xx"
End If

'this is line can be vary depand on the level number
no = Trim(comlevel.Text)
level = no
 a = Switch(no = 2, "Level2", no = 3, "Level3", no = 4, "Level4", _
 no = 5, "Level5", no = 6, "Level6", no = "All", "FinanceMaster where active <> '0'")


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
con2.close
On Error GoTo 0
con2.Open "Dsn=Finance;Uid=Sa;Pwd=;" ' this is for refresh the connection
 
 
recfin.Open "Select * from " & a, con2, adOpenKeyset, adLockOptimistic
ProgressBar1.Min = 0
ProgressBar1.Max = recfin.RecordCount + 1
ProgressBar1.Value = 0

    Dim recdelete As New ADODB.Recordset
    recdelete.Open "delete from TrialBalance where loguser ='" & cLogUser & "'", con2, adOpenKeyset, adLockOptimistic
    rectri.Open "Select * from TrialBalance", con2, adOpenKeyset, adLockOptimistic


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
               
recfin.MoveFirst
While recfin.EOF = False
recfinaccountcode = recfin!AccountCode
'this is all for GLMaster table
    If no = 2 Then
    'this is for all transaction
        recglb.Open "SELECT substring(AccountCode, 1, 3), SUM(CreditAmount) AS creditamount, SUM(DebitAmount) AS debitamount, CASE WHEN SUM(DebitAmount) >= SUM(creditamount) tHEN SUM(debitamount) - SUM(creditamount)else 0 END AS endingDebit, CASE WHEN SUM(creditamount) >= SUM(DebitAmount) THEN SUM(creditamount) - SUM(debitamount) Else 0 END AS endingCredit from " & cLogUser & "GLMaster where substring(accountcode,1,3) + '000000000' = " & "'" & recfinaccountcode & "'" & " and recorddate >= " & "'" & Format(dtpfrom.Value, "mm/dd/yyyy") & "'" & " and recorddate <= " & "'" & Format(dtpto.Value, "mm/dd/yyyy") & "'" & " GROUP BY substring(AccountCode, 1, 3)", con2, adOpenKeyset, adLockOptimistic
    'this is for opening balance upto specific date
         recope1.Open "SELECT substring(AccountCode, 1, 3), SUM(CreditAmount) AS creditamount, SUM(DebitAmount) AS debitamount from " & cLogUser & "GLMaster where substring(accountcode,1,3) + '000000000' = " & "'" & recfinaccountcode & "'" & " And recorddate < " & "'" & Format(dtpfrom.Value, "mm/dd/yyyy") & "'" & " GROUP BY substring(AccountCode, 1, 3)", con2, adOpenKeyset, adLockOptimistic
        'this is for opening and ending balance
        recope.Open "select sum(beginningdebit) as beginningdebit,sum(beginningcredit) as beginningcredit from " & cLogUser & "OpeningBalance where substring(accountcode,1,3) + '000000000' = " & "'" & recfinaccountcode & "'", con2, adOpenKeyset, adLockOptimistic
        
        'this is for update printed colum on glmaster table
        recupglb.Open "update " & cLogUser & "GLMaster  set printed = '1' where substring(accountcode,1,3) + '000000000' = " & "'" & recfinaccountcode & "'" & " and recorddate >= " & "'" & Format(dtpfrom.Value, "mm/dd/yyyy") & "'" & " and recorddate <= " & "'" & Format(dtpto.Value, "mm/dd/yyyy") & "'", con2, adOpenKeyset, adLockOptimistic
        recupope1.Open "update " & cLogUser & "GLMaster  set printed = '1' where substring(accountcode,1,3) + '000000000' = " & "'" & recfinaccountcode & "'" & " And recorddate < " & "'" & Format(dtpfrom.Value, "mm/dd/yyyy") & "'", con2, adOpenKeyset, adLockOptimistic
        recupope.Open "update " & cLogUser & "openingbalance set printed ='1' where substring(accountcode,1,3) + '000000000' = " & "'" & recfinaccountcode & "'", con2, adOpenKeyset, adLockOptimistic
    End If
    
    If no = 3 Then
        recglb.Open "SELECT substring(AccountCode, 1, 5), SUM(CreditAmount) AS creditamount, SUM(DebitAmount) AS debitamount, CASE WHEN SUM(DebitAmount) >= SUM(creditamount) tHEN SUM(debitamount) - SUM(creditamount)else 0 END AS endingDebit, CASE WHEN SUM(creditamount) >= SUM(DebitAmount) THEN SUM(creditamount) - SUM(debitamount) Else 0 END AS endingCredit from " & cLogUser & "GLMaster where substring(accountcode,1,5) + '0000000' = " & "'" & recfinaccountcode & "'" & " and recorddate >= " & "'" & Format(dtpfrom.Value, "mm/dd/yyyy") & "'" & " and recorddate <= " & "'" & Format(dtpto.Value, "mm/dd/yyyy") & "'" & "GROUP BY substring(AccountCode, 1, 5)", con2, adOpenKeyset, adLockOptimistic
        recope1.Open "SELECT substring(AccountCode, 1, 5), SUM(CreditAmount) AS creditamount, SUM(DebitAmount) AS debitamount from " & cLogUser & "GLMaster where substring(accountcode,1,5) + '0000000' = " & "'" & recfinaccountcode & "'" & " And recorddate < " & "'" & Format(dtpfrom.Value, "mm/dd/yyyy") & "'" & " GROUP BY substring(AccountCode, 1, 5)", con2, adOpenKeyset, adLockOptimistic
        recope.Open "select sum(beginningdebit) as beginningdebit,sum(beginningcredit) as beginningcredit from " & cLogUser & "OpeningBalance where substring(accountcode,1,5) + '0000000' = " & "'" & recfinaccountcode & "'", con2, adOpenKeyset, adLockOptimistic
        
        'this is for update printed colum on glmaster table
        recupglb.Open "update " & cLogUser & "GLMaster  set printed = '1' where substring(accountcode,1,5) + '0000000' = " & "'" & recfinaccountcode & "'" & " and recorddate >= " & "'" & Format(dtpfrom.Value, "mm/dd/yyyy") & "'" & " and recorddate <= " & "'" & Format(dtpto.Value, "mm/dd/yyyy") & "'", con2, adOpenKeyset, adLockOptimistic
        recupope1.Open "update " & cLogUser & "GLMaster  set printed = '1' where substring(accountcode,1,5) + '0000000' = " & "'" & recfinaccountcode & "'" & " And recorddate < " & "'" & Format(dtpfrom.Value, "mm/dd/yyyy") & "'", con2, adOpenKeyset, adLockOptimistic
        recupope.Open "update " & cLogUser & "openingbalance set printed ='1' where substring(accountcode,1,5) + '0000000' = " & "'" & recfinaccountcode & "'", con2, adOpenKeyset, adLockOptimistic
    End If
    
    If no = 4 Then
        recglb.Open "SELECT substring(AccountCode, 1, 7), SUM(CreditAmount) AS creditamount, SUM(DebitAmount) AS debitamount, CASE WHEN SUM(DebitAmount) >= SUM(creditamount) tHEN SUM(debitamount) - SUM(creditamount)else 0 END AS endingDebit, CASE WHEN SUM(creditamount) >= SUM(DebitAmount) THEN SUM(creditamount) - SUM(debitamount) Else 0 END AS endingCredit from " & cLogUser & "GLMaster where substring(accountcode,1,7) + '00000' = " & "'" & recfinaccountcode & "'" & " and recorddate >= " & "'" & Format(dtpfrom.Value, "mm/dd/yyyy") & "'" & " and recorddate <= " & "'" & Format(dtpto.Value, "mm/dd/yyyy") & "'" & "GROUP BY substring(AccountCode, 1, 7)", con2, adOpenKeyset, adLockOptimistic
         recope1.Open "SELECT substring(AccountCode, 1, 7), SUM(CreditAmount) AS creditamount, SUM(DebitAmount) AS debitamount from " & cLogUser & "GLMaster where substring(accountcode,1,7) + '00000' = " & "'" & recfinaccountcode & "'" & " And recorddate < " & "'" & Format(dtpfrom.Value, "mm/dd/yyyy") & "'" & " GROUP BY substring(AccountCode, 1, 7)", con2, adOpenKeyset, adLockOptimistic
        recope.Open "select sum(beginningdebit) as beginningdebit,sum(beginningcredit) as beginningcredit from " & cLogUser & "OpeningBalance where substring(accountcode,1,7) + '00000' = " & "'" & recfinaccountcode & "'", con2, adOpenKeyset, adLockOptimistic
        
        'this is for update printed colum on glmaster table
        recupglb.Open "update " & cLogUser & "GLMaster  set printed = '1' where substring(accountcode,1,7) + '00000' = " & "'" & recfinaccountcode & "'" & " and recorddate >= " & "'" & Format(dtpfrom.Value, "mm/dd/yyyy") & "'" & " and recorddate <= " & "'" & Format(dtpto.Value, "mm/dd/yyyy") & "'", con2, adOpenKeyset, adLockOptimistic
        recupope1.Open "update " & cLogUser & "GLMaster  set printed = '1' where substring(accountcode,1,7) + '00000' = " & "'" & recfinaccountcode & "'" & " And recorddate < " & "'" & Format(dtpfrom.Value, "mm/dd/yyyy") & "'", con2, adOpenKeyset, adLockOptimistic
        recupope.Open "update " & cLogUser & "openingbalance set printed ='1' where substring(accountcode,1,7) + '00000' = " & "'" & recfinaccountcode & "'", con2, adOpenKeyset, adLockOptimistic
    End If
    
    If no = 5 Then
        recglb.Open "SELECT substring(AccountCode, 1, 9), SUM(CreditAmount) AS creditamount, SUM(DebitAmount) AS debitamount, CASE WHEN SUM(DebitAmount) >= SUM(creditamount) tHEN SUM(debitamount) - SUM(creditamount)else 0 END AS endingDebit, CASE WHEN SUM(creditamount) >= SUM(DebitAmount) THEN SUM(creditamount) - SUM(debitamount) Else 0 END AS endingCredit from " & cLogUser & "GLMaster where substring(accountcode,1,9) + '000' = " & "'" & recfinaccountcode & "'" & " and recorddate >= " & "'" & Format(dtpfrom.Value, "mm/dd/yyyy") & "'" & " and recorddate <= " & "'" & Format(dtpto.Value, "mm/dd/yyyy") & "'" & "GROUP BY substring(AccountCode, 1, 9)", con2, adOpenKeyset, adLockOptimistic
        recope1.Open "SELECT substring(AccountCode, 1, 9), SUM(CreditAmount) AS creditamount, SUM(DebitAmount) AS debitamount from " & cLogUser & "GLMaster where substring(accountcode,1,9) + '000' = " & "'" & recfinaccountcode & "'" & " And recorddate < " & "'" & Format(dtpfrom.Value, "mm/dd/yyyy") & "'" & " GROUP BY substring(AccountCode, 1, 9)", con2, adOpenKeyset, adLockOptimistic
        recope.Open "select sum(beginningdebit) as beginningdebit,sum(beginningcredit) as beginningcredit from " & cLogUser & "OpeningBalance where substring(accountcode,1,9) + '000' = " & "'" & recfinaccountcode & "'", con2, adOpenKeyset, adLockOptimistic
        
        'this is for update printed colum on glmaster table
        recupglb.Open "update " & cLogUser & "GLMaster  set printed = '1' where substring(accountcode,1,9) + '000' = " & "'" & recfinaccountcode & "'" & " and recorddate >= " & "'" & Format(dtpfrom.Value, "mm/dd/yyyy") & "'" & " and recorddate <= " & "'" & Format(dtpto.Value, "mm/dd/yyyy") & "'", con2, adOpenKeyset, adLockOptimistic
        recupope1.Open "update " & cLogUser & "GLMaster  set printed = '1' where substring(accountcode,1,9) + '000' = " & "'" & recfinaccountcode & "'" & " And recorddate < " & "'" & Format(dtpfrom.Value, "mm/dd/yyyy") & "'", con2, adOpenKeyset, adLockOptimistic
        recupope.Open "update " & cLogUser & "openingbalance set printed ='1' where substring(accountcode,1,9) + '000' = " & "'" & recfinaccountcode & "'", con2, adOpenKeyset, adLockOptimistic
    End If
       
    If no = "All" Then
        recglb.Open "SELECT AccountCode, SUM(CreditAmount) AS creditamount, SUM(DebitAmount) AS debitamount, CASE WHEN SUM(DebitAmount) >= SUM(creditamount) tHEN SUM(debitamount) - SUM(creditamount)else 0 END AS endingDebit, CASE WHEN SUM(creditamount) >= SUM(DebitAmount) THEN SUM(creditamount) - SUM(debitamount) Else 0 END AS endingCredit from " & cLogUser & "GLMaster where accountcode = " & _
        "'" & recfinaccountcode & "'" & " and recorddate >= " & "'" & Format(dtpfrom.Value, "mm/dd/yyyy") & "'" & _
        " and recorddate <= " & "'" & Format(dtpto.Value, "mm/dd/yyyy") & "'" & _
        "GROUP BY AccountCode", con2, adOpenKeyset, adLockOptimistic
        
        recope1.Open "SELECT AccountCode, SUM(CreditAmount) AS creditamount, SUM(DebitAmount) AS debitamount from " & cLogUser & "GLMaster where accountcode = " & _
        "'" & recfinaccountcode & "'" & " And recorddate < " & "'" & Format(dtpfrom.Value, "mm/dd/yyyy") & "'" & _
        " GROUP BY AccountCode", con2, adOpenKeyset, adLockOptimistic
        
        recope.Open "select sum(beginningdebit) as beginningdebit,sum(beginningcredit) as beginningcredit from " & cLogUser & "OpeningBalance where accountcode = " & _
        "'" & recfinaccountcode & "'", con2, adOpenKeyset, adLockOptimistic
        
        'this is for update printed colum on glmaster table
        recupglb.Open "update " & cLogUser & "GLMaster  set printed = '1' where accountcode = '" & recfinaccountcode & "' and recorddate >= " & "'" & Format(dtpfrom.Value, "mm/dd/yyyy") & "'" & " and recorddate <= " & "'" & Format(dtpto.Value, "mm/dd/yyyy") & "'", con2, adOpenKeyset, adLockOptimistic
        recupope1.Open "update " & cLogUser & "GLMaster  set printed = '1' where accountcode = '" & recfinaccountcode & "' and recorddate < " & "'" & Format(dtpfrom.Value, "mm/dd/yyyy") & "'", con2, adOpenKeyset, adLockOptimistic
        recupope.Open "update " & cLogUser & "openingbalance set printed ='1' where accountcode = " & "'" & recfinaccountcode & "'", con2, adOpenKeyset, adLockOptimistic
    End If
        
        With rectri
            .addnew
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
            
            'this for profit and loss statement
            If no = "All" Then
                !levelno = "All"
            End If
            .Update
            
        End With
    DoEvents
    ProgressBar1.Value = ProgressBar1.Value + 1
    lp.caption = Format(((ProgressBar1.Value * 100) / ProgressBar1.Max), "###") & " %"
    recglb.close
    recope.close
    recope1.close

recfin.MoveNext
Wend

ProgressBar1.Value = 0

'**************** ' this is for the balance accountnumber the details from glmaster
'which is not included

Dim recupdatesecond As New ADODB.Recordset  'update recordset
Dim recglbbb As New ADODB.Recordset
recglb.Open "SELECT AccountCode, SUM(DebitAmount) AS debitamount, SUM(CreditAmount) AS creditamount from " & cLogUser & "GLMaster where printed = '0'" & " and recorddate >= " & "'" & Format(dtpfrom.Value, "mm/dd/yyyy") & "'" & " and recorddate <= " & "'" & Format(dtpto.Value, "mm/dd/yyyy") & "' GROUP BY AccountCode", con2, adOpenKeyset, adLockOptimistic

'MsgBox recglb.RecordCount & " " & recglb!accountcode
If recglb.BOF = False Then

    Dim recglbfin As New ADODB.Recordset
    'this is for get the accountname and arab name dont change this if you trial balance will be the wrong value
    recglbfin.Open "select * from " & cLogUser & "financemaster where active <> '0' and accountcode = " & "'" & Trim(recglb!AccountCode) & "'", con2, adOpenKeyset, adLockOptimistic
    
    If recglbfin.BOF = False Then
        recglbfinenglish = recglbfin!accountnameeng
        recglbfinarab = recglbfin!accountnamearab
    End If
    recglbfin.close
End If

'MsgBox recglb.RecordCount
ProgressBar1.Min = 0
ProgressBar1.Max = recglb.RecordCount + 1
While recglb.EOF = False
lblstatus.caption = "Adding Uncoverd Account Numbers...."
recfinaccountcode = Trim(recglb!AccountCode)
        recope.Open "select beginningdebit as beginningdebit,beginningcredit as beginningcredit from " & cLogUser & "OpeningBalance where accountcode = " & "'" & recfinaccountcode & "' and printed = '0'", con2, adOpenKeyset, adLockOptimistic
        recope1.Open "SELECT AccountCode, SUM(CreditAmount) AS creditamount, SUM(DebitAmount) AS debitamount from " & cLogUser & "GLMaster where printed = '0' and accountcode = " & "'" & recfinaccountcode & "'" & " And recorddate < " & "'" & Format(dtpfrom.Value, "mm/dd/yyyy") & "'" & " GROUP BY AccountCode", con2, adOpenKeyset, adLockOptimistic
        
        recglbbb.Open "SELECT sum(debitamount) as debitamount,sum(creditamount) as creditamount from " & cLogUser & "GLMaster where  printed = '0' and accountcode = " & "'" & recglb!AccountCode & "'" & " and  recorddate >= " & "'" & Format(dtpfrom.Value, "mm/dd/yyyy") & "'" & " and recorddate <= " & "'" & Format(dtpto.Value, "mm/dd/yyyy") & "' group by accountcode", con2, adOpenKeyset, adLockOptimistic

        'MsgBox recglbbb.RecordCount
        'MsgBox recope1.RecordCount
       'for update the printed column in opening and glmaster tables
        recupdatesecond.Open "update " & cLogUser & "openingbalance set printed ='1' where accountcode = " & "'" & recfinaccountcode & "'", con2, adOpenKeyset, adLockOptimistic
        recupdatesecond.Open "update " & cLogUser & "glmaster set printed ='1' where accountcode = " & "'" & recfinaccountcode & "'" & " And recorddate < " & "'" & Format(dtpfrom.Value, "mm/dd/yyyy") & "'", con2, adOpenKeyset, adLockOptimistic

        With rectri
            .addnew
            !LogUser = cLogUser
            !AccountCode = recglb!AccountCode
            
            '*********** this is come from finance master you have to care about this
            'MsgBox recglbfin!accountcode
            !accountname = recglbfinenglish
            !accountnameara = recglbfinarab
            '************************************
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

            If recglbbb.BOF = False Then
                !activitydebit = recglbbb!DebitAmount
                !ActivityCredit = recglbbb!creditamount
                recglbdebitamount = IIf(IsNull(recglbbb!DebitAmount), 0, recglbbb!DebitAmount)
                recglbcreditamount = IIf(IsNull(recglbbb!creditamount), 0, recglbbb!creditamount)

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
    
    recope.close
    recope1.close
    recglbbb.close
recglb.MoveNext
Wend
recglb.close
'*****************  this for unapplied opening balance

recope.Open "Select * from " & cLogUser & "openingbalance where printed ='0'", con2, adOpenKeyset, adLockOptimistic
ProgressBar1.Value = 0
ProgressBar1.Max = recope.RecordCount + 1
While recope.EOF = False
lblstatus.caption = "Adding Uncoverd Opening Balance Account Numbers...."
        With rectri
            .addnew
            !LogUser = cLogUser
            !AccountCode = recope!AccountCode
            !accountname = recope!accountnameeng
            !accountnameara = recope!accountnamearab
            ' take the begin balance from opening balance table
            If recope.BOF = False Then
                recopebeginningDebit = Val(IIf(IsNull(recope!beginningdebit), 0, recope!beginningdebit))
                recopebeginningCredit = Val(IIf(IsNull(recope!beginningcredit), 0, recope!beginningcredit))
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

                recglbdebitamount = 0
                recglbcreditamount = 0
            
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
    
recope.MoveNext
Wend
recope.close

'**************** end opening balance adding
ProgressBar1.Visible = False
lblstatus.Visible = False
lp.Visible = False
recfin.close

'this is drop table
Dim recdroptable As New ADODB.Recordset
'On Error GoTo ar
recdroptable.Open "Drop table " & cLogUser & "openingbalance", con2
recdroptable.Open "Drop table " & cLogUser & "financemaster", con2
recdroptable.Open "Drop table " & cLogUser & "glmaster", con2
'ar:
'MsgBox Err.Description

On Error Resume Next
con2.close
On Error GoTo 0
con2.Open "Dsn=Finance;Uid=Sa;Pwd=;" ' this is for refresh the connection
cmdprintactivity_Click ' call the print button
activity = 6
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
If cLogUser = "" Then
    cLogUser = "xx"
End If

activity = 1
If Trim(comlevel.Text) = "" Then
    MsgBox "Please check The Level That YOu Select", vbInformation, "Empty Level"
    Exit Sub
End If
If dtpfrom.Value > dtpto.Value Then
    MsgBox "Please Enter Your Date Correctly", vbInformation, "Disorder Date"
    Exit Sub
End If
cmdactivity.Enabled = False
cmdclose.Enabled = False
lp.Visible = True
Call prcgetdata
cmdprintactivity.Enabled = True
If closeme = "1" Then
    closeme = 2
    lp.Visible = False
    cmdactivity.Enabled = True
    cmdclose.Enabled = True
    Exit Sub
End If

cmdsummery.Enabled = True
cmdactivity.Enabled = True
lp.Visible = False
cmdclose.Enabled = True
End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdprintactivity_Click()

On Error Resume Next
dataanu.rsTrial_balance.close
On Error GoTo 0

dataanu.Trial_balance cLogUser
    If MsgBox("Select the Language to Print      ÇÎÊÇÑ áÛÉ ÇáØÈÇÚÉ  " & vbCrLf & "Yes :- For Arabic                                       ÈÇáÚÑÈí " & vbCrLf & "No  :- For English                                  ÈÇáÇäÌáíÒí", vbInformation + vbYesNo, "Language Option  ÇÎÊíÇÑ ÇááÛÉ") = vbYes Then
        passlevel re_ara_trialBalance_activity.Sections(2).Controls("Label15")
        passlevel1 re_ara_trialBalance_activity.Sections(2).Controls("Label3")
        re_ara_trialBalance_activity.Show 1
    Else
        passlevel re_trialBalance_activity.Sections(2).Controls("Label15")
        passlevel1 re_trialBalance_activity.Sections(2).Controls("Label3")
        re_trialBalance_activity.Show 1
    End If

End Sub


Private Sub cmdsummery_Click()
If cLogUser = "" Then
    cLogUser = "xx"
End If

activity = 2
If Trim(comlevel.Text) = "" Then
    MsgBox "Please check The Level That You Select", vbInformation, "Empty Level"
    Exit Sub
End If
lp.Visible = True

On Error Resume Next
dataanu.rstrial_balance_summery.close
On Error GoTo 0

dataanu.trial_balance_summery cLogUser

level = Trim(comlevel.Text)

If MsgBox("Select the Language to Print      ÇÎÊÇÑ áÛÉ ÇáØÈÇÚÉ  " & vbCrLf & "Yes :- For Arabic                                       ÈÇáÚÑÈí " & vbCrLf & "No  :- For English                                  ÈÇáÇäÌáíÒí", vbInformation + vbYesNo, "Language Option  ÇÎÊíÇÑ ÇááÛÉ") = vbYes Then
    passlevel re_TrialBalance_summerybased.Sections(2).Controls("Label15")
    passlevel1 re_TrialBalance_summerybased.Sections(2).Controls("Label3") ' to pass the from to date
    re_ara_TrialBalance_summerybased.Show 1
Else
    passlevel re_TrialBalance_summerybased.Sections(2).Controls("Label15")
    passlevel1 re_TrialBalance_summerybased.Sections(2).Controls("Label3") ' to pass the from to date
    re_TrialBalance_summerybased.Show 1

End If

activity = 6

'cmdsummery.Enabled = False
lp.Visible = False
End Sub

Private Sub comlevel_Click()
cmdsummery.Enabled = False
cmdprintactivity.Enabled = False
End Sub

Private Sub Form_Load()

con2.Open "Dsn=Finance;Uid=Sa;Pwd=;"

Dim reccheck As New ADODB.Recordset
reccheck.Open "select sum(debitamount)as debitamount ,sum(creditamount) as creditamount from glmaster", con2, adOpenKeyset, adLockOptimistic

If Val(reccheck!DebitAmount) <> Val(reccheck!creditamount) Then
    MsgBox "Your General Ledger is Not Balance Please Contact Administrator", vbInformation, "Un Balance"
    reccheck.close
    con2.close
    Unload Me
    Exit Sub
End If

For i = 2 To 5
    comlevel.AddItem i
Next
comlevel.AddItem "All"
 comlevel.ListIndex = 4
 
 dtpto.Value = Date
Dim fromdate As Date
datefromdate = "01/01/" & DatePart("yyyy", Date)
fromdate = Format(datefromdate, "mm/dd/yyyy")
dtpfrom.Value = fromdate

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
con2.close
End Sub


VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmprofitandlost 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "  Profit & (Lost) Statement  ÞÇÆãÉ ÇáÇÑÈÇÍ æÇáÎÓÇÆÑ"
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   ControlBox      =   0   'False
   Icon            =   "frmprofitandlost.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   5175
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      Begin VB.CommandButton cmdprint 
         Caption         =   "&Show"
         Height          =   375
         Left            =   3720
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdclose 
         Caption         =   "&Close"
         Height          =   375
         Left            =   3720
         TabIndex        =   3
         Top             =   840
         Width           =   975
      End
      Begin VB.ComboBox comchoice 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   855
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   135
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Visible         =   0   'False
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   238
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.Label lblprogress 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "progress..."
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
         TabIndex        =   7
         Top             =   720
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label lblper 
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
         Left            =   3120
         TabIndex        =   6
         Top             =   840
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Level"
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
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1080
      End
   End
End
Attribute VB_Name = "frmprofitandlost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim recfromtrial As New ADODB.Recordset
Dim reclev As New ADODB.Recordset
Dim recpro As New ADODB.Recordset
Dim con As New ADODB.Connection

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Dim reccheck As New ADODB.Recordset
reccheck.Open "Select * from trialbalance where loguser ='" & cLogUser & "' and levelno = 'All'", con, adOpenKeyset, adLockOptimistic
If reccheck.BOF = True Then
    MsgBox "Please Run The Trial Balance First With Option All", vbInformation, "No Data"
    reccheck.close
    con.close
    Unload Me
    Exit Sub
End If
reccheck.close

'reccheck.Open "select sum(debitamount)as debitamount ,sum(creditamount) as creditamount from glmaster", con, adOpenKeyset, adLockOptimistic
'
'If Val(reccheck!debitamount) <> Val(reccheck!creditamount) Then
'    MsgBox "Your General Ledger is Not Balance Please Contact Administrator", vbInformation, "Un Balance"
'    reccheck.Close
'    con.Close
'    Unload Me
'    Exit Sub
'End If

End Sub

Private Sub Form_Load()
On Error Resume Next
con.close
On Error GoTo 0
con.Open "dsn=Finance;Uid=Sa;Pwd=;"
For i = 2 To 5
comchoice.AddItem i
Next
comchoice.AddItem "All"
comchoice.ListIndex = 0
ProgressBar1.Min = 0
ProgressBar1.Max = 100
End Sub

Private Sub cmdPrint_Click()
If Trim(cLogUser) = "" Then
    cLogUser = "xx"
End If

If MsgBox("Are You Sure You Want to Calculate The TrialBalance On Currect Date", vbYesNo, "Conformation") = vbNo Then
    con.close
    Unload Me
    Exit Sub
End If


On Error Resume Next
recpro.close
On Error GoTo 0
recpro.Open "update TrialBalance set okprint = '0' where loguser = '" & cLogUser & "'", con, adOpenKeyset, adLockOptimistic

recpro.Open "Delete from profitandloss where loguser = '" & cLogUser & "'", con, adOpenKeyset, adLockOptimistic

recpro.Open "select * from profitandloss", con, adOpenKeyset, adLockOptimistic

b = Trim(comchoice.Text)
a = Switch(b = "2", "SUBSTRING(AccountCode, 1, 3)", _
            b = "3", "SUBSTRING(AccountCode, 1, 5)", _
            b = "4", "SUBSTRING(AccountCode, 1, 7)", _
            b = "5", "SUBSTRING(AccountCode, 1, 9)", _
            b = "All", "SUBSTRING(AccountCode, 1, 12)")
                                                                
l = Switch(b = 2, "Level2", b = 3, "Level3", b = 4, "Level4", b = 5, "Level5", b = "All", "FinanceMaster")

Dim recupdate As New ADODB.Recordset

'recfromtrial!code is showing the particular substring

With recpro
    .addnew
    !LogUser = cLogUser
    !details = "SALES"
    .Update
        'first
        recfromtrial.Open "SELECT " & a & "AS code, SUM(EndingDebit) - SUM(EndingCredit) AS endingbalance from TrialBalance" _
        & " WHERE loguser = '" & cLogUser & "' and  SUBSTRING(AccountCode, 1, 1) = '2' and ((SUBSTRING(AccountCode, 1, 3) >= '211' and" _
        & " SUBSTRING(AccountCode, 1, 3) <= '214')  or SUBSTRING(AccountCode, 1, 3) = '216') GROUP BY " _
        & a & " order by " & a, con, adOpenKeyset, adLockOptimistic
        
        reclev.Open "Select *," & a & " as code from " & l & "  WHERE SUBSTRING(AccountCode, 1, 1) = '2' and ((SUBSTRING(AccountCode, 1, 3) >= '211' and" _
        & " SUBSTRING(AccountCode, 1, 3) <= '214') or SUBSTRING(AccountCode, 1, 3) = '216') order by " _
        & a, con, adOpenKeyset, adLockOptimistic
        
        'if trial balancetable is null then
        If recpro.BOF = True Then
            MsgBox "Please Run The Trial Balance First Before Run This", vbInformation, "Profit & Lost"
            recfromtrial.close
            reclev.close
            recpro.close
            Exit Sub
        End If
    
         If recfromtrial.BOF = False Then
         recupdate.Open "Update TrialBalance set okprint = '1' WHERE loguser = '" & cLogUser & "' and SUBSTRING(AccountCode, 1, 1) = '2' and ((SUBSTRING(AccountCode, 1, 3) >= '211' and SUBSTRING(AccountCode, 1, 3) <= '214')  or SUBSTRING(AccountCode, 1, 3) = '216')", con, adOpenKeyset, adLockOptimistic
            recfromtrial.MoveFirst
        End If
        If reclev.BOF = False Then
            reclev.MoveFirst
        End If
        
ProgressBar1.Value = 0
ProgressBar1.Visible = True
lblprogress.Visible = True
lblper.Visible = True

cmdprint.Enabled = False
cmdclose.Enabled = False
comchoice.Enabled = False

        While recfromtrial.EOF = False
                ProgressBar1.Max = recfromtrial.RecordCount
                    .addnew
                    !LogUser = cLogUser
                    If b = 4 Then
                        getaccounntnumber Trim(recfromtrial!Code) & "00000", Acctname
                        !AccountCode = Trim(recfromtrial!Code) & "00000"
                        !accountname = Acctname
                    ElseIf b = 5 Then
                        getaccounntnumber Trim(recfromtrial!Code) & "000", Acctname
                        !AccountCode = Trim(recfromtrial!Code) & "000"
                        !accountname = Acctname
                    Else
                    !AccountCode = reclev!AccountCode
                    !accountname = reclev!accountnameeng
                    reclev.MoveNext
                    End If
                    !amount = (recfromtrial!endingbalance) * -1
                    firsttotal1 = firsttotal1 + Val(recfromtrial!endingbalance) * -1
                    .Update
                    recfromtrial.MoveNext
                If ProgressBar1.Value <> ProgressBar1.Max Then
                    ProgressBar1.Value = ProgressBar1.Value + 1
                Else
                    ProgressBar1.Value = 0
                End If
            lblper.caption = Format(((ProgressBar1.Value * 100) / ProgressBar1.Max), "###") & " %"
            lblprogress.caption = "Adding Sales Infomation ..."
        DoEvents
        Wend
        
    .addnew
    !LogUser = cLogUser
    !details = "Total Sales "
    !lastbalance = firsttotal1
    .Update
    recfromtrial.close
    reclev.close
    
    .addnew
    !LogUser = cLogUser
    !details = "Less:"
    .Update
    
    'second
        recfromtrial.Open "SELECT " & a & "AS code, SUM(EndingDebit) - SUM(EndingCredit) AS endingbalance from TrialBalance WHERE loguser = '" & cLogUser & "' and (SUBSTRING(AccountCode, 1, 1) = '2') and (SUBSTRING(AccountCode, 1, 3) >= '231') and (SUBSTRING(AccountCode, 1, 3) <= '233') GROUP BY " & a & " order by " & a, con, adOpenKeyset, adLockOptimistic
        reclev.Open "Select *," & a & " as code from " & l & " where (SUBSTRING(AccountCode, 1, 3) >= '231') and (SUBSTRING(AccountCode, 1, 3) <= '233') order by " & a, con, adOpenKeyset, adLockOptimistic
         'MsgBox recfromtrial.RecordCount
        'MsgBox reclev.RecordCount
         If recfromtrial.BOF = False Then
                  recupdate.Open "Update TrialBalance set okprint = '1'  WHERE loguser = '" & cLogUser & "' and (SUBSTRING(AccountCode, 1, 1) = '2') and (SUBSTRING(AccountCode, 1, 3) >= '231') and (SUBSTRING(AccountCode, 1, 3) <= '233')", con, adOpenKeyset, adLockOptimistic
            recfromtrial.MoveFirst
        End If
        If reclev.BOF = False Then
            reclev.MoveFirst
        End If
        
        While recfromtrial.EOF = False
                ProgressBar1.Max = recfromtrial.RecordCount
                    .addnew
                    !LogUser = cLogUser
                    If b = 4 Then
                        getaccounntnumber Trim(recfromtrial!Code) & "00000", Acctname
                        !AccountCode = Trim(recfromtrial!Code) & "00000"
                        !accountname = Acctname
                    ElseIf b = 5 Then
                        getaccounntnumber Trim(recfromtrial!Code) & "000", Acctname
                        !AccountCode = Trim(recfromtrial!Code) & "000"
                        !accountname = Acctname
                    Else
                    !AccountCode = reclev!AccountCode
                    !accountname = reclev!accountnameeng
                    reclev.MoveNext
                    End If
                    !amount = recfromtrial!endingbalance * -1
                    firsttotal2 = firsttotal2 + Val(recfromtrial!endingbalance) * -1
                    .Update
                    recfromtrial.MoveNext
                If ProgressBar1.Value <> ProgressBar1.Max Then
                    ProgressBar1.Value = ProgressBar1.Value + 1
                Else
                    ProgressBar1.Value = 0
                End If
    lblper.caption = Format(((ProgressBar1.Value * 100) / ProgressBar1.Max), "###") & " %"
           lblprogress.caption = "Calculating Total Sales ..."
        DoEvents
        Wend
    
    .addnew
    !LogUser = cLogUser
    !details = "Net Sales "
    firsttotal2 = (firsttotal1 + firsttotal2)
    !lastbalance = firsttotal2
    .Update
    
    .addnew
    !LogUser = cLogUser
    !details = "Less:"
    .Update
    recfromtrial.close
    reclev.close
    
    'third
        recfromtrial.Open "SELECT " & a & "AS code, SUM(EndingDebit) - SUM(EndingCredit) AS endingbalance from TrialBalance WHERE loguser = '" & cLogUser & "' and  (SUBSTRING(AccountCode, 1, 1) = '2') and (SUBSTRING(AccountCode, 1, 3) >= '221') and (SUBSTRING(AccountCode, 1, 3) <= '222') GROUP BY " & a & " order by " & a, con, adOpenKeyset, adLockOptimistic
        reclev.Open "Select *," & a & " as code from " & l & " where (SUBSTRING(AccountCode, 1, 3) >= '221') and (SUBSTRING(AccountCode, 1, 3) <= '222') order by " & a, con, adOpenKeyset, adLockOptimistic
        'MsgBox recfromtrial.RecordCount
        'MsgBox reclev.RecordCount
         If recfromtrial.BOF = False Then
                  recupdate.Open "Update TrialBalance set okprint = '1'  WHERE loguser = '" & cLogUser & "' and (SUBSTRING(AccountCode, 1, 1) = '2') and (SUBSTRING(AccountCode, 1, 3) >= '221') and (SUBSTRING(AccountCode, 1, 3) <= '222')", con, adOpenKeyset, adLockOptimistic

            recfromtrial.MoveFirst
        End If
        If reclev.BOF = False Then
            reclev.MoveFirst
        End If
        
        While recfromtrial.EOF = False
                ProgressBar1.Max = recfromtrial.RecordCount
                    .addnew
                    !LogUser = cLogUser
                    If b = 4 Then
                        getaccounntnumber Trim(recfromtrial!Code) & "00000", Acctname
                        !AccountCode = Trim(recfromtrial!Code) & "00000"
                        !accountname = Acctname
                    ElseIf b = 5 Then
                        getaccounntnumber Trim(recfromtrial!Code) & "000", Acctname
                        !AccountCode = Trim(recfromtrial!Code) & "000"
                        !accountname = Acctname
                    Else
                    !AccountCode = reclev!AccountCode
                    !accountname = reclev!accountnameeng
                    reclev.MoveNext
                    End If
                    !amount = recfromtrial!endingbalance * -1
                    firsttotal3 = firsttotal3 + Val(recfromtrial!endingbalance) * -1
                    .Update
                    recfromtrial.MoveNext
                 If ProgressBar1.Value <> ProgressBar1.Max Then
                    ProgressBar1.Value = ProgressBar1.Value + 1
                Else
                    ProgressBar1.Value = 0
                End If
   lblper.caption = Format(((ProgressBar1.Value * 100) / ProgressBar1.Max), "###") & " %"
            lblprogress.caption = "Calculating Net Sales ..."
        DoEvents
        Wend
        
    .addnew
    !LogUser = cLogUser
    !details = "Gross Profit(Loss)"
    firsttotal3 = (firsttotal2 + firsttotal3)
    !lastbalance = firsttotal3
    .Update
    
    .addnew
    !LogUser = cLogUser
    !details = "Less:"
    .Update
    
    recfromtrial.close
    reclev.close
    
    'fourth
        recfromtrial.Open "SELECT " & a & "AS code, SUM(EndingDebit) - SUM(EndingCredit) AS endingbalance from TrialBalance WHERE loguser = '" & cLogUser & "' and (SUBSTRING(AccountCode, 1, 1) = '2') and (SUBSTRING(AccountCode, 1, 3) >= '223') and (SUBSTRING(AccountCode, 1, 3) <= '226') GROUP BY " & a & " order by " & a, con, adOpenKeyset, adLockOptimistic
        reclev.Open "Select *," & a & " as code from " & l & " where (SUBSTRING(AccountCode, 1, 3) >= '223') and (SUBSTRING(AccountCode, 1, 3) <= '226') order by " & a, con, adOpenKeyset, adLockOptimistic
        'MsgBox recfromtrial.RecordCount
        'MsgBox reclev.RecordCount

         If recfromtrial.BOF = False Then
                  recupdate.Open "Update TrialBalance set okprint = '1' WHERE loguser = '" & cLogUser & "' and (SUBSTRING(AccountCode, 1, 1) = '2') and (SUBSTRING(AccountCode, 1, 3) >= '223') and (SUBSTRING(AccountCode, 1, 3) <= '226')", con, adOpenKeyset, adLockOptimistic
            recfromtrial.MoveFirst
        End If
        If reclev.BOF = False Then
            reclev.MoveFirst
        End If
        
        While recfromtrial.EOF = False
                ProgressBar1.Max = recfromtrial.RecordCount
                 .addnew
                    !LogUser = cLogUser
                    If b = 4 Then
                        getaccounntnumber Trim(recfromtrial!Code) & "00000", Acctname
                        !AccountCode = Trim(recfromtrial!Code) & "00000"
                        !accountname = Acctname
                    ElseIf b = 5 Then
                        getaccounntnumber Trim(recfromtrial!Code) & "000", Acctname
                        !AccountCode = Trim(recfromtrial!Code) & "000"
                        !accountname = Acctname
                    Else
                    !AccountCode = reclev!AccountCode
                    !accountname = reclev!accountnameeng
                    reclev.MoveNext
                    End If
                    !amount = (recfromtrial!endingbalance) * -1
                    firsttotal4 = firsttotal4 + Val(recfromtrial!endingbalance) * -1
                    .Update
                    recfromtrial.MoveNext
                    If ProgressBar1.Value <> ProgressBar1.Max Then
                    ProgressBar1.Value = ProgressBar1.Value + 1
                Else
                    ProgressBar1.Value = 0
                End If
        lblper.caption = Format(((ProgressBar1.Value * 100) / ProgressBar1.Max), "###") & " %"
        lblprogress.caption = "Calculating Gross Profit(Loss) ..."
        DoEvents
        Wend
    .addnew
    !LogUser = cLogUser
    !details = "Profit(Loss) From Operation"
    firsttotal4 = (firsttotal3 + firsttotal4)
    !lastbalance = firsttotal4
    .Update
    
    .addnew
    !LogUser = cLogUser
    !details = "Less:"
    .Update
    recfromtrial.close
    reclev.close
    
    'fifth
        recfromtrial.Open "SELECT " & a & "AS code, SUM(EndingDebit) - SUM(EndingCredit) AS endingbalance from TrialBalance WHERE loguser = '" & cLogUser & "' and (SUBSTRING(AccountCode, 1, 1) = '2') and (SUBSTRING(AccountCode, 1, 3) >= '227') and (SUBSTRING(AccountCode, 1, 3) <= '227') GROUP BY " & a & " order by " & a, con, adOpenKeyset, adLockOptimistic
        reclev.Open "Select *," & a & " as code from " & l & " where (SUBSTRING(AccountCode, 1, 3) >= '227') and (SUBSTRING(AccountCode, 1, 3) <= '227') order by " & a, con, adOpenKeyset, adLockOptimistic
        'MsgBox recfromtrial.RecordCount
        'MsgBox reclev.RecordCount

        If recfromtrial.BOF = False Then
                 recupdate.Open "Update TrialBalance set okprint = '1' WHERE loguser = '" & cLogUser & "' and (SUBSTRING(AccountCode, 1, 1) = '2') and (SUBSTRING(AccountCode, 1, 3) >= '227') and (SUBSTRING(AccountCode, 1, 3) <= '227')", con, adOpenKeyset, adLockOptimistic

            recfromtrial.MoveFirst
        End If
        If reclev.BOF = False Then
            reclev.MoveFirst
        End If
        
        While recfromtrial.EOF = False
                ProgressBar1.Max = recfromtrial.RecordCount
                    .addnew
                    !LogUser = cLogUser
                    If b = 4 Then
                        getaccounntnumber Trim(recfromtrial!Code) & "00000", Acctname
                        !AccountCode = Trim(recfromtrial!Code) & "00000"
                        !accountname = Acctname
                    ElseIf b = 5 Then
                        getaccounntnumber Trim(recfromtrial!Code) & "000", Acctname
                        !AccountCode = Trim(recfromtrial!Code) & "000"
                        !accountname = Acctname
                    Else
                    !AccountCode = reclev!AccountCode
                    !accountname = reclev!accountnameeng
                    reclev.MoveNext
                    End If
                    !amount = (recfromtrial!endingbalance) * -1
                    firsttotal5 = firsttotal5 + Val(recfromtrial!endingbalance) * -1
                    .Update
                    recfromtrial.MoveNext
                    If ProgressBar1.Value <> ProgressBar1.Max Then
                    ProgressBar1.Value = ProgressBar1.Value + 1
                Else
                    ProgressBar1.Value = 0
                End If
        lblper.caption = Format(((ProgressBar1.Value * 100) / ProgressBar1.Max), "###") & " %"
        lblprogress.caption = "Calculating Profit(Loss) From Operation ..."
        DoEvents
        Wend
    
    .addnew
    !LogUser = cLogUser
    !details = "Add:"
    .Update
    recfromtrial.close
    reclev.close
    
    'sixth
        recfromtrial.Open "SELECT " & a & "AS code, SUM(EndingDebit) - SUM(EndingCredit) AS endingbalance from TrialBalance WHERE loguser = '" & cLogUser & "' and (SUBSTRING(AccountCode, 1, 1) = '2') and (SUBSTRING(AccountCode, 1, 3) >= '215') and (SUBSTRING(AccountCode, 1, 3) <= '215') GROUP BY " & a & " order by " & a, con, adOpenKeyset, adLockOptimistic
        reclev.Open "Select *," & a & " as code from " & l & " where (SUBSTRING(AccountCode, 1, 3) >= '215') and (SUBSTRING(AccountCode, 1, 3) <= '215') order by " & a, con, adOpenKeyset, adLockOptimistic
        'MsgBox recfromtrial.RecordCount
        'MsgBox reclev.RecordCount

        If recfromtrial.BOF = False Then
                 recupdate.Open "Update TrialBalance set okprint = '1' WHERE loguser = '" & cLogUser & "' and (SUBSTRING(AccountCode, 1, 1) = '2') and (SUBSTRING(AccountCode, 1, 3) >= '215') and (SUBSTRING(AccountCode, 1, 3) <= '215')", con, adOpenKeyset, adLockOptimistic

            recfromtrial.MoveFirst
        End If
        If reclev.BOF = False Then
            reclev.MoveFirst
        End If
        
        While recfromtrial.EOF = False
                ProgressBar1.Max = recfromtrial.RecordCount
                    .addnew
                    !LogUser = cLogUser
                    If b = 4 Then
                        getaccounntnumber Trim(recfromtrial!Code) & "00000", Acctname
                        !AccountCode = Trim(recfromtrial!Code) & "00000"
                        !accountname = Acctname
                    ElseIf b = 5 Then
                        getaccounntnumber Trim(recfromtrial!Code) & "000", Acctname
                        !AccountCode = Trim(recfromtrial!Code) & "000"
                        !accountname = Acctname
                    Else
                    !AccountCode = reclev!AccountCode
                    !accountname = reclev!accountnameeng
                    reclev.MoveNext
                    End If
                    !amount = (recfromtrial!endingbalance) * -1
                    firsttotal6 = firsttotal6 + Val(recfromtrial!endingbalance) * -1
                    .Update
                    recfromtrial.MoveNext
                If ProgressBar1.Value <> ProgressBar1.Max Then
                    ProgressBar1.Value = ProgressBar1.Value + 1
                Else
                    ProgressBar1.Value = 0
                End If
        lblper.caption = Format(((ProgressBar1.Value * 100) / ProgressBar1.Max), "###") & " %"
        lblprogress.caption = "Calculating Pre Tax Profit ..."
        DoEvents
        Wend
    .addnew
    !LogUser = cLogUser
    !details = "Pre-Tax Profit(Loss)"
    firsttotal6 = (firsttotal4 + firsttotal5 - firsttotal6)
    !lastbalance = firsttotal6
    .Update
    
    .addnew
    !LogUser = cLogUser
    !details = "Less:"
    .Update
    recfromtrial.close
    reclev.close
    
    'seven
        recfromtrial.Open "SELECT " & a & "AS code, SUM(EndingDebit) - SUM(EndingCredit) AS endingbalance from TrialBalance WHERE loguser = '" & cLogUser & "' and (SUBSTRING(AccountCode, 1, 1) = '2') and (SUBSTRING(AccountCode, 1, 3) >= '228') and (SUBSTRING(AccountCode, 1, 3) <= '228') GROUP BY " & a & " order by " & a, con, adOpenKeyset, adLockOptimistic
        reclev.Open "Select *," & a & " as code from " & l & " where (SUBSTRING(AccountCode, 1, 3) >= '228') and (SUBSTRING(AccountCode, 1, 3) <= '228') order by " & a, con, adOpenKeyset, adLockOptimistic
         If recfromtrial.BOF = False Then
                  recupdate.Open "Update TrialBalance set okprint = '1' WHERE loguser = '" & cLogUser & "' and (SUBSTRING(AccountCode, 1, 1) = '2') and (SUBSTRING(AccountCode, 1, 3) >= '228') and (SUBSTRING(AccountCode, 1, 3) <= '228')", con, adOpenKeyset, adLockOptimistic

            recfromtrial.MoveFirst
        End If
        If reclev.BOF = False Then
            reclev.MoveFirst
        End If
        While recfromtrial.EOF = False
                ProgressBar1.Max = recfromtrial.RecordCount
                    .addnew
                    !LogUser = cLogUser
                    If b = 4 Then
                        getaccounntnumber Trim(recfromtrial!Code) & "00000", Acctname
                        !AccountCode = Trim(recfromtrial!Code) & "00000"
                        !accountname = Acctname
                    ElseIf b = 5 Then
                        getaccounntnumber Trim(recfromtrial!Code) & "000", Acctname
                        !AccountCode = Trim(recfromtrial!Code) & "000"
                        !accountname = Acctname
                    Else
                    !AccountCode = reclev!AccountCode
                    !accountname = reclev!accountnameeng
                    reclev.MoveNext
                    End If
                    !amount = (recfromtrial!endingbalance) * -1
                    firsttotal7 = firsttotal7 + Val(recfromtrial!endingbalance) * -1
                    .Update
                    recfromtrial.MoveNext
                If ProgressBar1.Value <> ProgressBar1.Max Then
                    ProgressBar1.Value = ProgressBar1.Value + 1
                Else
                    ProgressBar1.Value = 0
                End If
            lblper.caption = Format(((ProgressBar1.Value * 100) / ProgressBar1.Max), "###") & " %"
            lblprogress.caption = "Calculating Net Activity Profit(Loss) ..."
        DoEvents
        Wend
    
    .addnew
    !LogUser = cLogUser
    !details = "Net Activity Profit(Loss)"
    firsttotal7 = (firsttotal6 + firsttotal7)
    !lastbalance = firsttotal7
    .Update
    
    
    .addnew
    !LogUser = cLogUser
    !details = "Less:"
    .Update
    
    recfromtrial.close
    reclev.close
    
    'eight
        recfromtrial.Open "SELECT " & a & "AS code, SUM(EndingDebit) - SUM(EndingCredit) AS endingbalance from TrialBalance WHERE loguser = '" & cLogUser & "' and (SUBSTRING(AccountCode, 1, 1) = '2') and (SUBSTRING(AccountCode, 1, 3) >= '229') and (SUBSTRING(AccountCode, 1, 3) <= '229') GROUP BY " & a & " order by " & a, con, adOpenKeyset, adLockOptimistic
        reclev.Open "Select *," & a & " as code from " & l & " where (SUBSTRING(AccountCode, 1, 3) >= '229') and (SUBSTRING(AccountCode, 1, 3) <= '229') order by " & a, con, adOpenKeyset, adLockOptimistic
        'MsgBox recfromtrial.RecordCount
        'MsgBox reclev.RecordCount
         If recfromtrial.BOF = False Then
                  recupdate.Open "Update TrialBalance set okprint = '1' WHERE loguser = '" & cLogUser & "' and (SUBSTRING(AccountCode, 1, 1) = '2') and (SUBSTRING(AccountCode, 1, 3) >= '229') and (SUBSTRING(AccountCode, 1, 3) <= '229')", con, adOpenKeyset, adLockOptimistic

            recfromtrial.MoveFirst
        End If
        If reclev.BOF = False Then
            reclev.MoveFirst
        End If
        While recfromtrial.EOF = False
        ProgressBar1.Max = recfromtrial.RecordCount
                    .addnew
                    !LogUser = cLogUser
                    If b = 4 Then
                        getaccounntnumber Trim(recfromtrial!Code) & "00000", Acctname
                        !AccountCode = Trim(recfromtrial!Code) & "00000"
                        !accountname = Acctname
                    ElseIf b = 5 Then
                        getaccounntnumber Trim(recfromtrial!Code) & "000", Acctname
                        !AccountCode = Trim(recfromtrial!Code) & "000"
                        !accountname = Acctname
                    Else
                    !AccountCode = reclev!AccountCode
                    !accountname = reclev!accountnameeng
                    reclev.MoveNext
                    End If
                    !amount = recfromtrial!endingbalance * -1
                    firsttotal8 = firsttotal8 + Val(recfromtrial!endingbalance) * -1
                    .Update
                    recfromtrial.MoveNext
                If ProgressBar1.Value <> ProgressBar1.Max Then
                    ProgressBar1.Value = ProgressBar1.Value + 1
                Else
                    ProgressBar1.Value = 0
                End If
        lblper.caption = Format(((ProgressBar1.Value * 100) / ProgressBar1.Max), "###") & " %"
         lblprogress.caption = "Calculating Total Period's Profit(Loss) ..."
        DoEvents
        Wend
    .addnew
    !LogUser = cLogUser
    !details = " "
    .Update
    
    .addnew
    !LogUser = cLogUser
    !lastbalance = "------------------------------"
    .Update
    
    .addnew
    !LogUser = cLogUser
    !details = "Total Period's Profit(Loss)"
    firsttotal8 = (firsttotal7 + firsttotal8)
    !lastbalance = firsttotal8
    .Update
    
    .addnew
    !LogUser = cLogUser
    !lastbalance = "=================="
    .Update
    
    '.AddNew
    '!Details = "Profit Per Share     No. of Shares : 1,000,000 "
    '!lastBalance = Format((Val(firsttotal8) / 1000000), "##0.##0")
    '.Update
    
    recfromtrial.close
    reclev.close

End With
ProgressBar1.Value = ProgressBar1.Max
ProgressBar1.Visible = False
lblprogress.Visible = False
lblper.Visible = False
cmdprint.Enabled = True
cmdclose.Enabled = True
comchoice.Enabled = True
recpro.close

On Error Resume Next
dataanu.rscomprofitandloss.close
On Error GoTo 0

dataanu.comprofitandloss cLogUser
re_Profitandloss_statement.Show 1

End Sub

Private Sub Label2_Click()

End Sub


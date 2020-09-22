VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmbalancesheet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Balance Sheet ÇáãíÒÇäíÉ ÇáÚãæãíÉ  "
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3795
   Icon            =   "frmbalancesheet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   3795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   3375
      Begin VB.CommandButton cmdprint 
         Caption         =   "Show"
         Height          =   375
         Left            =   2280
         TabIndex        =   3
         Top             =   720
         Width           =   855
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
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   150
         Left            =   120
         TabIndex        =   1
         Top             =   1320
         Visible         =   0   'False
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   265
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label Label4 
         Caption         =   "ßÔÝ ÈÃÑÕÏÉ ÇáãíÒÇäíÉ"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   15
         Left            =   0
         TabIndex        =   6
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmbalancesheet.frx":0442
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Balance sheet  Statemen "
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1830
      End
   End
End
Attribute VB_Name = "frmbalancesheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim recfromtrial As New ADODB.Recordset
Dim recfromtrial2 As New ADODB.Recordset
Dim reclev As New ADODB.Recordset
Dim recbal As New ADODB.Recordset
Dim con As New ADODB.Connection

Private Sub cmdPrint_Click()
'this is for balance sheet table
Dim recnis1 As New ADODB.Recordset

recbal.Open "Select * from BalanceSheet", con, adOpenKeyset, adLockOptimistic
If recbal.BOF = False Then
    Dim recdelete As New ADODB.Recordset
    recdelete.Open "Delete from balancesheet", con, adOpenKeyset, adLockOptimistic
End If

recbal.Requery
ProgressBar1.Visible = True
ProgressBar1.Max = 100
b = Trim(comchoice.Text)
a = Switch(b = "2", "SUBSTRING(AccountCode, 1, 3)", _
            b = "3", "SUBSTRING(AccountCode, 1, 5)", _
            b = "4", "SUBSTRING(AccountCode, 1, 7)", _
            b = "5", "SUBSTRING(AccountCode, 1, 9)", _
            b = "6" Or b = "All", "SUBSTRING(AccountCode, 1, 12)") ' from trial balance
            
v = Switch(b = "2", "'[1][0-9][0-9]000000000'", _
            b = "3", "'[1][0-9][0-9][0-9][0-9]0000000'", _
            b = "4", "'[1][0-9][0-9][0-9][0-9][0-9][0-9]00000'", _
            b = "5", "'[1][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9]000'", _
            b = "6" Or b = "All", "'[1][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9]'") ' from trial balance
            
    'this will add only the (selectedlevel-1) because selected level will be added to the sub level also
    
    anu = "SELECT accountcode,accountname,EndingDebit-EndingCredit AS balance from TrialBalance WHERE" & _
    "(SUBSTRING(AccountCode, 1, 1) = '1')and (SUBSTRING(AccountCode, 1, 2) <> '12') and accountcode like " & v & " Order By accountcode "
    
    nis = "SELECT accountcode,accountname,EndingDebit-EndingCredit AS balance from TrialBalance WHERE" & _
    "(SUBSTRING(AccountCode, 1, 1) = '1') and (SUBSTRING(AccountCode, 1, 2) = '12') and accountcode like " & v & " Order By accountcode "
        Dim recnis As New ADODB.Recordset
        
            recnis.Open nis, con, adOpenKeyset, adLockOptimistic ' this is for only contra asset accounts
            MsgBox recnis.RecordCount
            If recnis.BOF = False Then
                recnis.MoveFirst
            End If
           recfromtrial.Open anu, con, adOpenKeyset, adLockOptimistic
           MsgBox recfromtrial.RecordCount 'debug purpost
While recfromtrial.EOF = False
            
        X = Switch(b = "2", Mid(recfromtrial!AccountCode, 1, 3), _
            b = "3", Mid(recfromtrial!AccountCode, 1, 5), _
            b = "4", Mid(recfromtrial!AccountCode, 1, 7), _
            b = "5", Mid(recfromtrial!AccountCode, 1, 9), _
            b = "6" Or b = "All", recfromtrial!AccountCode) ' from trial balance
            
        If recnis.EOF = False Then
            X1 = Switch(b = "2", Mid(recnis!AccountCode, 1, 3), _
                b = "3", Mid(recnis!AccountCode, 1, 5), _
                b = "4", Mid(recnis!AccountCode, 1, 7), _
                b = "5", Mid(recnis!AccountCode, 1, 9), _
                b = "6" Or b = "All", recnis!AccountCode)
                
                nis1 = "SELECT " & a & " as accountcode, SUM(EndingDebit - EndingCredit) AS balance from TrialBalance WHERE " & _
                    a & "='" & X1 & "'" & " and accountcode >= " & "'" & Trim(recnis!AccountCode) & "'" & " GROUP BY " & a
                    
                    recnis1.Open nis1, con, adOpenKeyset, adLockOptimistic
                    MsgBox recnis1.RecordCount
                    recnis11 = 1
        End If

'recnis.MoveNext
'recnis1.Close
                    anu1 = "SELECT " & a & "as accountcode, SUM(EndingDebit - EndingCredit) AS balance from TrialBalance WHERE " & _
                    a & "='" & X & "'" & " and accountcode >= " & "'" & Trim(recfromtrial!AccountCode) & "'" & " GROUP BY " & a
                    
                    recfromtrial2.Open anu1, con, adOpenKeyset, adLockOptimistic
                           
                           MsgBox recfromtrial2.RecordCount  'debug purpost
            With recbal
                If recfromtrial2.BOF = False Then
                    If X = Trim(recfromtrial2!AccountCode) Then
                        .AddNew
                            !AccountCode = recfromtrial!AccountCode
                            !accountname = recfromtrial!accountname
                            Dim contrabalance As Currency
                            If recnis.EOF = True Then
                                contrabalance = 0
                            Else
                                contrabalance = (recnis1!Balance)
                            End If
                                
                            !Balance = Val(recfromtrial!Balance) + (recfromtrial2!Balance) - (contrabalance) ' check the last balance
                            If recfromtrial!AccountCode Like "###000000000" Then
                            !levelname = "2"
                            ElseIf recfromtrial!AccountCode Like "#####0000000" Then
                            !levelname = "3"
                            ElseIf recfromtrial!AccountCode Like "#######00000" Then
                            !levelname = "4"
                            ElseIf recfromtrial!AccountCode Like "########000" Then
                            !levelname = "5"
                            ElseIf recfromtrial!AccountCode Like "############" Then
                            !levelname = "6"
                            End If
                            t = t + Val(recfromtrial!Balance) + (recfromtrial2!Balance) ' this is for testing
                        .Update
                    Else
                        .AddNew
                            !AccountCode = recfromtrial!AccountCode
                            !accountname = recfromtrial!accountname
                            !Balance = recfromtrial!Balance
                            If recfromtrial!AccountCode Like "###000000000" Then
                            !levelname = "2"
                            ElseIf recfromtrial!AccountCode Like "#####0000000" Then
                            !levelname = "3"
                            ElseIf recfromtrial!AccountCode Like "#######00000" Then
                            !levelname = "4"
                            ElseIf recfromtrial!AccountCode Like "########000" Then
                            !levelname = "5"
                            ElseIf recfromtrial!AccountCode Like "############" Then
                            !levelname = "6"
                            End If
                            .Update
                    End If
                Else
                        .AddNew
                            !AccountCode = recfromtrial!AccountCode
                            !accountname = recfromtrial!accountname
                            If recnis.EOF = True Then
                                contrabalance = 0
                            Else
                                contrabalance = (recnis1!Balance)
                                MsgBox recnis1!AccountCode
                            End If
                            !Balance = (recfromtrial!Balance) - (contrabalance)
                            If recfromtrial!AccountCode Like "###000000000" Then
                            !levelname = "2"
                            ElseIf recfromtrial!AccountCode Like "#####0000000" Then
                            !levelname = "3"
                            ElseIf recfromtrial!AccountCode Like "#######00000" Then
                            !levelname = "4"
                            ElseIf recfromtrial!AccountCode Like "########000" Then
                            !levelname = "5"
                            ElseIf recfromtrial!AccountCode Like "############" Then
                            !levelname = "6"
                            End If
                        .Update

                End If
                                    

            End With
                        If recnis11 = 1 Then
                            recnis1.Close
                            recnis11 = 2
                        End If
                           recfromtrial2.Close
                        If ProgressBar1.Value <> ProgressBar1.Max Then
                            ProgressBar1.Value = ProgressBar1.Value + 1
                        Else
                            ProgressBar1.Value = 0
                        End If
                    recfromtrial.MoveNext
                    If recnis.EOF = False Then
                    recnis.MoveNext
                    End If
Wend
            recnis.Close
            recfromtrial.Close
            ProgressBar1.Value = ProgressBar1.Max
            ProgressBar1.Visible = False
recbal.Requery
recbal.Close
re_Balancesheet.Show 1
End Sub

Private Sub Form_Load()
con.Open "dsn=Finance;Uid=Sa;Pwd=;"
For i = 2 To 6
comchoice.AddItem i
Next
comchoice.AddItem "All"
comchoice.ListIndex = 0
ProgressBar1.Min = 0
ProgressBar1.Max = 100
End Sub

Private Sub Form_Unload(Cancel As Integer)
con.Close
End Sub

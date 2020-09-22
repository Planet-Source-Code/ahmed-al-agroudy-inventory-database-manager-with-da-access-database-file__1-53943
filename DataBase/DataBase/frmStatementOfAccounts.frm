VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmStatmentOfAccounts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Statement Of Accounts ßÔÝ ÍÓÇÈ "
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8040
   ControlBox      =   0   'False
   Icon            =   "frmStatementOfAccounts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   8040
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Left            =   1200
      Top             =   4680
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   150
      Left            =   120
      TabIndex        =   26
      Top             =   2280
      Visible         =   0   'False
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   265
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Grouping Printing ÇáØÈÇÚÉ ÈÇáãÌãæÚÇÊ"
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
      TabIndex        =   24
      Top             =   0
      Width           =   4095
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
      TabIndex        =   12
      Top             =   0
      Value           =   -1  'True
      Width           =   3495
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   2640
      Width           =   7815
      Begin VB.CommandButton cmdprint 
         Caption         =   "P&rint ØÈÇÚÉ"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5040
         TabIndex        =   37
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton cmdclose 
         Caption         =   "&Close ÛáÞ "
         Height          =   375
         Left            =   6360
         TabIndex        =   29
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton cmdpreview 
         Caption         =   "&Preview ÚÑÖ"
         Height          =   375
         Left            =   3700
         TabIndex        =   19
         Top             =   1440
         Width           =   1335
      End
      Begin VB.ComboBox comfromname 
         Height          =   315
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   480
         Width           =   4695
      End
      Begin VB.ComboBox comtoname 
         Height          =   315
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   840
         Width           =   4695
      End
      Begin VB.ComboBox comto 
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Top             =   840
         Width           =   1575
      End
      Begin VB.ComboBox comfrom 
         Height          =   315
         Left            =   1440
         TabIndex        =   1
         Top             =   480
         Width           =   1575
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
         Format          =   48955393
         CurrentDate     =   37578
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
         Format          =   48955393
         CurrentDate     =   37578
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Name  ÇÓã ÇáÍÓÇÈ"
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
         Left            =   3000
         TabIndex        =   39
         Top             =   240
         Width           =   2145
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Çáí "
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   1560
         Width           =   525
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From ãä "
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From ãä "
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Çáí "
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   525
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Numbers  ÑÞã ÇáÍÓÇÈ"
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
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1935
      Left            =   120
      TabIndex        =   13
      Top             =   300
      Visible         =   0   'False
      Width           =   7815
      Begin VB.CommandButton cmdprint1 
         Caption         =   "P&rint ØÈÇÚÉ"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5040
         TabIndex        =   36
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton cmdclose1 
         Caption         =   "&Close ÛáÞ "
         Height          =   375
         Left            =   6360
         TabIndex        =   28
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton cmdpreview1 
         Caption         =   "&Preview ÚÑÖ"
         Height          =   375
         Left            =   3700
         TabIndex        =   18
         Top             =   1440
         Width           =   1335
      End
      Begin VB.ComboBox commain 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   480
         Width           =   2415
      End
      Begin VB.ComboBox comsub 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   840
         Width           =   2415
      End
      Begin MSComCtl2.DTPicker dtpfrom1 
         Height          =   315
         Left            =   1440
         TabIndex        =   20
         Top             =   1200
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Format          =   48955393
         CurrentDate     =   37578
      End
      Begin MSComCtl2.DTPicker dtpto1 
         Height          =   315
         Left            =   1440
         TabIndex        =   21
         Top             =   1560
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Format          =   48955393
         CurrentDate     =   37578
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ÇáãÌãæÚÉ ÇáÑÆíÓíÉ "
         Height          =   195
         Index           =   2
         Left            =   3840
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   480
         Width           =   1260
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ÇáãÌãæÚÉ ÇáÝÑÚíÉ"
         Height          =   195
         Index           =   2
         Left            =   3945
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   840
         Width           =   1140
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Çáì"
         Height          =   195
         Index           =   2
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ãä "
         Height          =   195
         Index           =   2
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   1200
         Width           =   225
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Group Details ÊÝÇÕíá ÇáãÌãæÚÉ"
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
         TabIndex        =   25
         Top             =   240
         Width           =   2370
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From "
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   23
         Top             =   1200
         Width           =   390
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To "
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   22
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sub Group"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   17
         Top             =   840
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Main Group"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   16
         Top             =   480
         Width           =   825
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar2 
      Height          =   150
      Left            =   5160
      TabIndex        =   38
      Top             =   2280
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      TabIndex        =   31
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label Label51 
      BackStyle       =   0  'Transparent
      Caption         =   "checking time"
      Height          =   255
      Left            =   3840
      TabIndex        =   30
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label lp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0 %"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7320
      TabIndex        =   27
      Top             =   2250
      Visible         =   0   'False
      Width           =   315
   End
End
Attribute VB_Name = "frmStatmentOfAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim recfin As New ADODB.Recordset ' this is for the financemaster to add the combobox
Dim recacc As New ADODB.Recordset ' this is for the financemaster for search details
Dim recop As New ADODB.Recordset 'this is for openinng balance
Dim recglop As New ADODB.Recordset ' this is for the glbalance for opening
Dim recglde As New ADODB.Recordset 'this is for the glbalance for details
Dim con As New ADODB.Connection ' the connection objects
Dim recprint As New ADODB.Recordset 'statementofaccount table
Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdclose1_Click()
Unload Me
End Sub

Private Sub cmdpreview_Click()
Dim recdroptable As New ADODB.Recordset

If Trim(cLogUser) = "" Then
    cLogUser = "xx"
End If

'this is all to create the temptable
Dim reccreatetable As New ADODB.Recordset

On Error GoTo errcreatetable
    reccreatetable.Open "Select * into " & cLogUser & "openingbalance from openingbalance", con
    reccreatetable.Open "Select * into " & cLogUser & "financemaster from FinanceMaster", con
    reccreatetable.Open "Select * into " & cLogUser & "glmaster from glmaster", con
    
errcreatetable:
'MsgBox Err.Number & "  " & Err.Description
    If Err.Number <> 0 Then
    
        If MsgBox("Your User Name Has Been Using In Another Workstation;" & vbCrLf & "Please LogOff And Try Again or Do You Want To Continue... ?", vbInformation + vbYesNo, "Multi Login") = vbNo Then
            cmdpreview.Enabled = True
            cmdclose.Enabled = True
            cmdprint.Enabled = True
            On Error Resume Next
            con.close
            On Error GoTo 0
            Exit Sub
        Else
            recdroptable.Open "Drop table " & cLogUser & "openingbalance", con
            recdroptable.Open "Drop table " & cLogUser & "financemaster", con
            recdroptable.Open "Drop table " & cLogUser & "glmaster", con
            
            reccreatetable.Open "Select * into " & cLogUser & "openingbalance from openingbalance", con
            reccreatetable.Open "Select * into " & cLogUser & "financemaster from FinanceMaster", con
            reccreatetable.Open "Select * into " & cLogUser & "glmaster from glmaster", con

        End If
    End If
'end create the temptable
On Error GoTo 0


On Error Resume Next
    con.close
On Error GoTo 0
con.Open "Dsn=Finance;Uid=Sa;Pwd=;" ' this is for refresh the connection

   Dim recdelete As New ADODB.Recordset
   recdelete.Open "delete from statementofaccount where loguser ='" & cLogUser & "'", con, adOpenKeyset, adLockOptimistic
   recprint.Open "Select * from statementofaccount", con, adOpenKeyset, adLockOptimistic
   
If Val(comto.Text) < Val(comfrom.Text) Or Trim(comto.Text) = "" Or Trim(comfrom.Text) = "" Then
    MsgBox "please Enter Your AccountNumber Correctly", vbInformation, "Disorder Numbers"
    Exit Sub
End If

If dtpfrom.Value > dtpto.Value Then
    MsgBox "Please Enter Your Date Correctly", vbInformation, "Disorder Date"
    Exit Sub
End If
'now i open only the selected account numbers from combo box
cmdpreview.Enabled = False
cmdclose.Enabled = False
cmdprint.Enabled = False

recacc.Open "SELECT * from " & cLogUser & "financemaster where active <> '0' and  accountcode >=" & "'" & comfrom.Text & "'" & "and accountcode <=" & "'" & comto.Text & "'", con, adOpenKeyset, adLockOptimistic
If recacc.BOF = False Then
recacc.MoveFirst
Else
    MsgBox "You Have No Data to Print the Details", vbInformation, "No Data"
    recacc.close
    cmdpreview1.Enabled = True
    cmdclose1.Enabled = True
    cmdprint1.Enabled = True
    Exit Sub
End If
ProgressBar1.Visible = True
ProgressBar2.Visible = True
lp.Visible = True
ProgressBar1.Value = 0
ProgressBar2.Value = 0
ProgressBar1.Min = 0
ProgressBar1.Max = recacc.RecordCount
While recacc.EOF = False

Dim openingdebit As Currency
Dim openingcredit As Currency
Dim debitopenfromgl As Currency
Dim creditopenfromgl As Currency
Dim rose As Currency
Dim rosedebit As Currency
Dim rosecredit As Currency
Dim recgldecreditamount As Currency ' credit amount
Dim recgldedebitamount As Currency ' debit amount

'end take mother name

    ' now i open the opening balance for the particular account number
    recop.Open "select * from " & cLogUser & "openingbalance where accountcode=" & "'" & recacc!AccountCode & "'", con, adOpenKeyset, adLockOptimistic
    If recop.BOF = False Then
        recop.MoveFirst
        openingdebit = recop!beginningdebit ' opening debit
        openingcredit = recop!beginningcredit ' opening credit
    Else
        openingdebit = 0
        openingcredit = 0
    End If
    recop.close ' i close the account

    'now open the glmaster for the opening balance
    recglop.Open "SELECT SUM(DebitAmount) AS debitopenfromgl, SUM(CreditAmount) AS creditopenfromgl FROM  " & cLogUser & "GLMaster wHERE recorddate < " & "'" & Format(dtpfrom.Value, "mm/dd/yyyy") & "'" & " and accountcode=" & "'" & recacc!AccountCode & "'", con, adOpenKeyset, adLockOptimistic
    If recglop.BOF = False Then
        recglop.MoveFirst
         debitopenfromgl = IIf(IsNull(recglop!debitopenfromgl), 0, recglop!debitopenfromgl)
         creditopenfromgl = IIf(IsNull(recglop!creditopenfromgl), 0, recglop!creditopenfromgl)
        openingdebit = openingdebit + debitopenfromgl ' last opening debit
        openingcredit = openingcredit + creditopenfromgl ' last opening credit
    End If
    recglop.close
    
    openingdebit = openingdebit - openingcredit
    
    If openingdebit < 0 Then
        openingcredit = openingdebit * -1
        openingdebit = 0
    Else
        openingcredit = 0
    End If
    
    With recprint ' adding the opening balance Details
        .addnew
        !LogUser = cLogUser
        !AccountCode = recacc!AccountCode
        !accountname = recacc!accountnameeng
        !accountnamearab = recacc!accountnamearab
        'end take mother name
        'getallmothername Trim(recacc!AccountCode), c
        'some diclarations
        '!mothername = c
        !recorddate = dtpfrom.Value
        !JOurnalNo = "OPENING BALANCE "
        !autocode = " "
        !particulars = " "
        !Debit = openingdebit
        !Credit = openingcredit
        !Balance = Val(openingdebit) - Val(openingcredit)
        !directbalane = Val(openingdebit) - Val(openingcredit)
        .Update
    End With
    rose = Val(openingdebit) - Val(openingcredit)
    
    ' now open glmaster for the details
    recglde.Open "SELECT * FROM  " & cLogUser & "GLMaster wHERE recorddate >=" & "'" & Format(dtpfrom.Value, "mm/dd/yyyy") & "'" & "and recorddate <=" & "'" & Format(dtpto.Value, "mm/dd/yyyy") & "'" & "and accountcode=" & "'" & recacc!AccountCode & "' order by recorddate,autonumber", con, adOpenKeyset, adLockOptimistic
    If recglde.BOF = False Then
    recglde.MoveFirst
'recglde.Open "insert into glmaster select '" & _
'cLogUser & "','" & recacc!AccountCode & "','" & recacc!accountnameeng & "','" & recacc!accountnamearab & "','" & _
' c & "','" & recglde!recorddate & "','" & recglde!JOurnalNo & "','" & recglde!autoNumber & "','" & _
' recglde!particulars & "','" & recglde!Cashdetails & "','" & _
'                        " from   " & _
'cLogUser & "GLMaster wHERE recorddate >=" & "'" & Format(dtpfrom.Value, "mm/dd/yyyy") & _
'"'" & "and recorddate <=" & "'" & Format(dtpto.Value, "mm/dd/yyyy") & "'" & _
'"and accountcode=" & "'" & recacc!AccountCode & _
'"' order by recorddate,autonumber", con, adOpenKeyset, adLockOptimistic
        ProgressBar2.Value = 0
        ProgressBar2.Max = recglde.RecordCount
        While recglde.EOF = False
                With recprint ' adding the Details From general ledger
                    .addnew
                        !LogUser = cLogUser
                        !AccountCode = recacc!AccountCode
                        !accountname = recacc!accountnameeng
                        !accountnamearab = recacc!accountnamearab

                        'end take mother name
                        'getallmothername Trim(recacc!AccountCode), c
                        'some diclarations
                        '!mothername = c
                        !recorddate = recglde!recorddate
                        !JOurnalNo = recglde!JOurnalNo
                        !autocode = recglde!autoNumber
                        !particulars = recglde!particulars

                        recgldecreditamount = recglde!creditamount
                        recgldedebitamount = recglde!DebitAmount
                        !Debit = recgldedebitamount
                        !Credit = recgldecreditamount
                        rose = Val(rose) + recgldedebitamount - recgldecreditamount

                        !Balance = Val(rose)
                        !directbalane = recgldedebitamount - recgldecreditamount
                        rosedebit = Val(rosedebit) + Val(recgldedebitamount)
                        rosecredit = Val(rosecredit) + Val(recgldecreditamount)
                    .Update
                End With
                DoEvents
                ProgressBar2.Value = ProgressBar2.Value + 1

        recglde.MoveNext
        Wend
    End If
    recglde.close
recacc.MoveNext
If ProgressBar1.Value <> ProgressBar1.Max Then
    ProgressBar1.Value = ProgressBar1.Value + 1
Else
    ProgressBar1.Value = 1
End If
lp.caption = Format(((ProgressBar1.Value * 100) / ProgressBar1.Max), "###") & " %"
rose = 0
rosedebit = 0
rosecredit = 0
Wend
lp.Visible = False
ProgressBar1.Visible = False
ProgressBar2.Visible = False
recacc.close


If Trim(cLogUser) = "" Then
    cLogUser = "xx"
End If


'On Error GoTo ar
recdroptable.Open "Drop table " & cLogUser & "openingbalance", con
recdroptable.Open "Drop table " & cLogUser & "financemaster", con
recdroptable.Open "Drop table " & cLogUser & "glmaster", con
'ar:
'MsgBox Err.Description

On Error Resume Next
    con.close
On Error GoTo 0
con.Open "Dsn=Finance;Uid=Sa;Pwd=;" ' this is for refresh the connection

cmdPrint_Click

cmdpreview.Enabled = True
cmdclose.Enabled = True
cmdprint.Enabled = True
    
End Sub
Private Sub lbldate(a As RptLabel)
a.caption = "From Date :  " & Format(dtpfrom.Value, "dd/mm/yyyy") & "   To :  " & Format(dtpto.Value, "dd/mm/yyyy")
End Sub
Private Sub lbldateara(a As RptLabel)
Dim dfrom As Date
Dim dto As Date
allword = "  ãä  ÊÇÑíÎ   " & DatePart("yyyy", Format(dtpfrom.Value, "dd/mm/yyyy")) & "/" & DatePart("m", Format(dtpfrom.Value, "dd/mm/yyyy")) & "/" & DatePart("d", Format(dtpfrom.Value, "dd/mm/yyyy")) & "  áÛÇíÉ  " & DatePart("yyyy", Format(dtpto.Value, "dd/mm/yyyy")) & "/" & DatePart("m", Format(dtpto.Value, "dd/mm/yyyy")) & "/" & DatePart("d", Format(dtpto.Value, "dd/mm/yyyy"))
'allword = "  ãä  ÊÇÑíÎ   " & Format(dtpfrom.Value, "dd/mm/yyyy") & "  áÛÇíÉ  " & Format(dtpto.Value, "dd/mm/yyyy")
a.caption = allword
End Sub

Private Sub lblaccountnumber(a As RptLabel)
a.caption = "From Account Number :  " & Trim(comfrom.Text) & "   To :  " & Trim(comto.Text)
End Sub

Private Sub cmdpreview1_Click()

Dim recdroptable As New ADODB.Recordset

If Trim(cLogUser) = "" Then
    cLogUser = "xx"
End If

'this is all to create the temptable
Dim reccreatetable As New ADODB.Recordset
On Error GoTo errcreatetable
    reccreatetable.Open "Select * into " & cLogUser & "openingbalance from openingbalance", con
    reccreatetable.Open "Select * into " & cLogUser & "financemaster from FinanceMaster", con
    reccreatetable.Open "Select * into " & cLogUser & "glmaster from glmaster", con
    
errcreatetable:
    If Err.Number <> 0 Then
        If MsgBox("Your User Name Has Been Using In Another Workstation;" & vbCrLf & "Please LogOff And Try Again or Do You Want To Continue... ?", vbInformation + vbYesNo, "Multi Login") = vbNo Then
            cmdpreview1.Enabled = True
            cmdclose1.Enabled = True
            cmdprint1.Enabled = True
            On Error Resume Next
            con.close
            On Error GoTo 0
            Exit Sub
        Else
            recdroptable.Open "Drop table " & cLogUser & "openingbalance", con
            recdroptable.Open "Drop table " & cLogUser & "financemaster", con
            recdroptable.Open "Drop table " & cLogUser & "glmaster", con
            
            reccreatetable.Open "Select * into " & cLogUser & "openingbalance from openingbalance", con
            reccreatetable.Open "Select * into " & cLogUser & "financemaster from FinanceMaster", con
            reccreatetable.Open "Select * into " & cLogUser & "glmaster from glmaster", con
        End If
    End If

'end create the temptable

On Error Resume Next
    con.close
'On Error GoTo 0

con.Open "Dsn=Finance;Uid=Sa;Pwd=;" ' this is for refresh the connection

Dim recdelete As New ADODB.Recordset
recdelete.Open "delete from statementofaccount where loguser ='" & cLogUser & "'", con, adOpenKeyset, adLockOptimistic
recprint.Open "Select * from statementofaccount", con, adOpenKeyset, adLockOptimistic

If Trim(commain.Text) = "" Or Trim(comsub.Text) = "" Then
    MsgBox "please Choose the Main Code & Sub Code Correctly", vbInformation, "Invalid Codes or Names"
    Exit Sub
End If

If dtpfrom1.Value > dtpto1.Value Then
    MsgBox "Please Enter Your Date Correctly", vbInformation, "Disorder Date"
    Exit Sub
End If
DoEvents
cmdpreview1.Enabled = False
cmdclose1.Enabled = False
cmdprint1.Enabled = False
'now i open only the selected account numbers from combo box

On Error Resume Next
recacc.close ' if some time is open this will close that
On Error GoTo 0

If Trim(Mid(Trim(comsub.Text), 1, 7)) <> "0000" Then
    'if it is not zero then should opent only the specific account numbers
    recacc.Open "SELECT * from groupmember where groupcatcode = '" & Trim(Mid(Trim(commain.Text), 1, 4)) & "' and  groupnamecode = '" & Trim(Mid(Trim(comsub.Text), 1, 7)) & "'", con, adOpenKeyset, adLockOptimistic
Else
    'if it is zero it should open all codes
    recacc.Open "SELECT * from groupmember where groupcatcode = '" & Trim(Mid(Trim(commain.Text), 1, 4)) & "'", con, adOpenKeyset, adLockOptimistic
End If

'MsgBox recacc.RecordCount
If recacc.BOF = False Then
recacc.MoveFirst
Else
    MsgBox "You Have No Data to Print the Details", vbInformation, "No Data"
    recacc.close
    cmdpreview1.Enabled = True
    cmdclose1.Enabled = True
    cmdprint1.Enabled = True
    Exit Sub
End If
ProgressBar1.Visible = True
ProgressBar2.Visible = True
lp.Visible = True
ProgressBar1.Value = 0
ProgressBar2.Value = 0
ProgressBar1.Min = 0
ProgressBar1.Max = recacc.RecordCount
While recacc.EOF = False

Dim openingdebit As Currency
Dim openingcredit As Currency
Dim debitopenfromgl As Currency
Dim creditopenfromgl As Currency
Dim rose As Currency
Dim rosedebit As Currency
Dim rosecredit As Currency
Dim recgldecreditamount As Currency ' credit amount
Dim recgldedebitamount As Currency ' debit amount

'end take mother name
    ' now i open the opening balance for the particular account number
    recop.Open "select * from " & cLogUser & "openingbalance where accountcode=" & "'" & recacc!memberAcctCode & "'", con, adOpenKeyset, adLockOptimistic
    If recop.BOF = False Then
        recop.MoveFirst
        openingdebit = recop!beginningdebit ' opening debit
        openingcredit = recop!beginningcredit ' opening credit
    Else
        openingdebit = 0
        openingcredit = 0
    End If
    recop.close ' i close the account

    'now open the glmaster for the opening balance
    recglop.Open "SELECT SUM(DebitAmount) AS debitopenfromgl, SUM(CreditAmount) AS creditopenfromgl FROM  " & cLogUser & "GLMaster wHERE recorddate <" & "'" & Format(dtpfrom1.Value, "mm/dd/yyyy") & "'" & "and accountcode=" & "'" & recacc!memberAcctCode & "'", con, adOpenKeyset, adLockOptimistic
    If recglop.BOF = False Then
        recglop.MoveFirst
        On Error Resume Next
         debitopenfromgl = recglop!debitopenfromgl
         creditopenfromgl = recglop!creditopenfromgl
         On Error GoTo 0
        openingdebit = openingdebit + debitopenfromgl ' last opening debit
        openingcredit = openingcredit + creditopenfromgl ' last opening credit
    End If
    recglop.close
    
     If openingdebit < 0 Then
        openingcredit = openingdebit * -1
        openingdebit = 0
    Else
        openingcredit = 0
    End If
    
    With recprint ' adding the opening balance Details
        .addnew
        !LogUser = cLogUser
        !AccountCode = recacc!memberAcctCode
        'getallmothername Trim(recacc!memberAcctCode), c
        !accountname = recacc!memberNameEng
        !accountnamearab = recacc!memberNameArab
        '!mothername = c
        !recorddate = dtpfrom1.Value
        !JOurnalNo = "OPENING BALANCE "
        !autocode = " "
        !particulars = " "
        !Debit = openingdebit
        !Credit = openingcredit
        !Balance = Val(openingdebit) - Val(openingcredit)
        !directbalane = Val(openingdebit) - Val(openingcredit)
        .Update
    End With
    rose = Val(openingdebit) - Val(openingcredit)
    
    ' now open glmaster for the details
    recglde.Open "SELECT * FROM  " & cLogUser & "GLMaster wHERE recorddate >=" & "'" & Format(dtpfrom1.Value, "mm/dd/yyyy") & "'" & "and recorddate <=" & "'" & Format(dtpto1.Value, "mm/dd/yyyy") & "'" & "and accountcode=" & "'" & recacc!memberAcctCode & "' order by recorddate,autonumber", con, adOpenKeyset, adLockOptimistic
    If recglde.BOF = False Then
    recglde.MoveFirst
        ProgressBar2.Max = recglde.RecordCount
        ProgressBar2.Value = 0
        While recglde.EOF = False
                With recprint ' adding the Details From general ledger
                    .addnew
                        !LogUser = cLogUser
                        !AccountCode = recacc!memberAcctCode
                        !accountname = recacc!memberNameEng
                        !accountnamearab = recacc!memberNameArab
                        'getallmothername Trim(recacc!memberAcctCode), c ' for mother name
                        '!mothername = c
                        !recorddate = recglde!recorddate
                        !JOurnalNo = recglde!JOurnalNo
                        !autocode = recglde!autoNumber
                        !particulars = recglde!particulars
                        !araparticulars = recglde!araparticulars
                        
                        recgldecreditamount = recglde!creditamount
                        recgldedebitamount = recglde!DebitAmount
                        !Debit = recgldedebitamount
                        !Credit = recgldecreditamount
                        rose = Val(rose) + recgldedebitamount - recgldecreditamount
                        
                        !Balance = Val(rose)
                        !directbalane = recgldedebitamount - recgldecreditamount
                        rosedebit = Val(rosedebit) + Val(recgldedebitamount)
                        rosecredit = Val(rosecredit) + Val(recgldecreditamount)
                    .Update
                End With
        ProgressBar2.Value = ProgressBar2.Value + 1
        recglde.MoveNext
        Wend
    End If
    recglde.close
recacc.MoveNext
If ProgressBar1.Value <> ProgressBar1.Max Then
    ProgressBar1.Value = ProgressBar1.Value + 1
Else
    ProgressBar1.Value = 1
End If
lp.caption = Format(((ProgressBar1.Value * 100) / ProgressBar1.Max), "###") & " %"
rose = 0
rosedebit = 0
rosecredit = 0
DoEvents
Wend
lp.Visible = False
ProgressBar1.Visible = False
ProgressBar2.Visible = False
recacc.close

'On Error GoTo ar
If Trim(cLogUser) = "" Then
    cLogUser = "xx"
End If

recdroptable.Open "Drop table " & cLogUser & "openingbalance", con
recdroptable.Open "Drop table " & cLogUser & "financemaster", con
recdroptable.Open "Drop table " & cLogUser & "glmaster", con
'ar:
'MsgBox Err.Description

On Error Resume Next
con.close
On Error GoTo 0

con.Open "Dsn=Finance;Uid=Sa;Pwd=;" ' this is for refresh the connection
'this is drop table

cmdprint1_Click

cmdpreview1.Enabled = True
cmdclose1.Enabled = True
cmdprint1.Enabled = True
End Sub

Private Sub cmdPrint_Click()
On Error Resume Next
dataanu.rscom_statementofaccounts_Grouping.close
On Error GoTo 0

dataanu.com_statementofaccounts_Grouping cLogUser


If MsgBox("Select the Language to Print      ÇÎÊÇÑ áÛÉ ÇáØÈÇÚÉ  " & vbCrLf & "Yes :- For Arabic                                       ÈÇáÚÑÈí " & vbCrLf & "No  :- For English                                  ÈÇáÇäÌáíÒí", vbInformation + vbYesNo, "Language Option  ÇÎÊíÇÑ ÇááÛÉ") = vbYes Then
    lbldateara re_ara_statementofaccount.Sections(3).Controls("lbldate")
    re_ara_statementofaccount.Show 1
Else
    lbldate re_eng_statementofaccount.Sections(3).Controls("lbldate")
    re_eng_statementofaccount.Show 1
End If

End Sub

Private Sub cmdprint1_Click()

On Error Resume Next
dataanu.rscom_statementofaccounts_Grouping.close
On Error GoTo 0

dataanu.com_statementofaccounts_Grouping cLogUserÏ

If MsgBox("Select the Language to Print      ÇÎÊÇÑ áÛÉ ÇáØÈÇÚÉ  " & vbCrLf & "Yes :- For Arabic                                       ÈÇáÚÑÈí " & vbCrLf & "No  :- For English                                  ÈÇáÇäÌáíÒí", vbInformation + vbYesNo, "Language Option  ÇÎÊíÇÑ ÇááÛÉ") = vbYes Then
    lbldateara re_ara_statementofaccount.Sections(3).Controls("lbldate")
    re_ara_statementofaccount.Show 1
Else
    lbldate re_eng_statementofaccount.Sections(3).Controls("lbldate")
    re_eng_statementofaccount.Show 1
End If

End Sub

Private Sub comfrom_Click()
'this is for find the account name
Dim recfindaccount As New ADODB.Recordset
recfindaccount.Open "select * from financemaster where active <> '0' and  accountcode = " & "'" & Trim(comfrom.Text) & "'", con, adOpenKeyset, adLockOptimistic
If recfindaccount.BOF = False Then
    comfromname.Text = recfindaccount!accountnamearab & "\" & recfindaccount!accountnameeng
End If
recfindaccount.close
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
recfindaccount.Open "select * from financemaster where active <> '0' and  accountcode = " & "'" & Trim(comfrom.Text) & "'", con, adOpenKeyset, adLockOptimistic
If recfindaccount.BOF = False Then
    comfromname.Text = recfindaccount!accountnamearab & "\" & recfindaccount!accountnameeng
    comto.Text = Trim(comfrom.Text)
Else
    comfrom.Text = "  "
    comfrom.SetFocus
    Exit Sub
End If
recfindaccount.close
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
recfindacc.Open "Select * from financemaster where active <> '0' and  accountnameeng = " & "'" & namename & "'", con, adOpenKeyset, adLockOptimistic
If recfindacc.BOF = False Then
    comfrom.Text = recfindacc!AccountCode
End If
 recfindacc.close
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

recfindacc.Open "Select * from financemaster where active <> '0' and  accountnameeng = " & "'" & namename & "'", con, adOpenKeyset, adLockOptimistic
If recfindacc.BOF = False Then
    comfrom.Text = recfindacc!AccountCode
Else
    MsgBox "Please Choose Correct Account Name", vbInformation, "Invalid Name"
    comfromname.SetFocus
    Exit Sub
End If
 recfindacc.close
' end change the account number

End Sub

Private Sub commain_Click()
comsub.clear

Dim recsub As New ADODB.Recordset
recsub.Open "select * from groupname where groupcatcode = '" & Trim(Mid(Trim(commain.Text), 1, 4)) & "'", con, adOpenKeyset, adLockOptimistic
    If recsub.BOF = False Then
        If recsub.RecordCount > 1 Then
            comsub.AddItem "0000     All"
        End If
        While recsub.EOF = False
            comsub.AddItem recsub!Idno & "     " & recsub!GroupNameEng
            recsub.MoveNext
        Wend
    End If
recsub.close

End Sub

Private Sub Command1_Click()
End Sub

Private Sub comto_Click()
'this is for find the account name
Dim recfindaccount As New ADODB.Recordset
recfindaccount.Open "select * from financemaster where active <> '0' and  accountcode = " & "'" & Trim(comto.Text) & "'", con, adOpenKeyset, adLockOptimistic
If recfindaccount.BOF = False Then
    comtoname.Text = recfindaccount!accountnamearab & "\" & recfindaccount!accountnameeng
End If
recfindaccount.close
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
recfindaccount.Open "select * from financemaster where active <> '0' and  accountcode = " & "'" & Trim(comto.Text) & "'", con, adOpenKeyset, adLockOptimistic
If recfindaccount.BOF = False Then
    comtoname.Text = recfindaccount!accountnamearab & "\" & recfindaccount!accountnameeng
Else
comto.Text = " "
Exit Sub
End If
recfindaccount.close
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
recfindacc.Open "Select * from financemaster where active <> '0' and  accountnameeng = " & "'" & namename & "'", con, adOpenKeyset, adLockOptimistic
If recfindacc.BOF = False Then
    comto.Text = recfindacc!AccountCode
End If
 recfindacc.close
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
    cmdpreview.SetFocus
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
recfindacc.Open "Select * from financemaster where active <> '0' and  accountnameeng = " & "'" & namename & "'", con, adOpenKeyset, adLockOptimistic
If recfindacc.BOF = False Then
    comto.Text = recfindacc!AccountCode
Else
    MsgBox "Please Choose Correct Account Name", vbInformation, "Invalid Name"
    comtoname.SetFocus
    Exit Sub
End If
 recfindacc.close
' end change the account number


End Sub

Private Sub dtpfrom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    dtpto.SetFocus
End If

End Sub

Private Sub dtpto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdpreview.SetFocus
End If

End Sub

Private Sub Form_Load()
Timer2.Interval = 1
Label33.caption = Time
Mainform.sbStatusBar.Panels(1).Text = "Status : General Ledger ..."

con.Open "Dsn=finance;Uid=sa;Pwd=;"

Dim constring As String
Dim myclass As New HabitatClass
Dim sqltable As Boolean
Dim xtable As String
Dim conrecfin As New ADODB.Connection

constring = "Dsn=finance;Uid=sa;Pwd=;"
xtable = "SELECT * from financemaster where active <> '0' order by accountcode"
sqltable = True
myclass.GetTables recfin, conrecfin, xtable, constring, sqltable
'recfin.Open "SELECT * from financemaster where active <> '0' order by accountcode", con, adOpenKeyset, adLockOptimistic, adCmdText

comfrom.clear
comto.clear
comfromname.clear
comtoname.clear

recfin.MoveFirst
While recfin.EOF = False
    comfrom.AddItem recfin!AccountCode
    comto.AddItem recfin!AccountCode

    anu = recfin!accountnamearab & "\" & recfin!accountnameeng

    comfromname.AddItem anu
    comtoname.AddItem anu
    recfin.MoveNext
Wend


recprint.Open "Select * from statementofaccount", con, adOpenKeyset, adLockOptimistic
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
recmain.Open "groupcat", con, adOpenKeyset, adLockOptimistic
    If recmain.BOF = False Then
        While recmain.EOF = False
            commain.AddItem Trim(recmain!Code) & "     " & Trim(recmain!NameEng)
            recmain.MoveNext
        Wend
    End If
recmain.close
Frame1.Top = Frame2.Top
Frame1.Left = Frame2.Left
Me.Height = 2985
Timer2.Interval = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
con.close
On Error GoTo 0
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

Private Sub optbranch_Click()
'If optbranch.Value = True Then
'    combranch.Enabled = True
'Else
'    combranch.Enabled = False
'End If
End Sub

Private Sub optbranch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    comfrom.SetFocus
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

Private Sub Timer2_Timer()
Label51.caption = Time
End Sub

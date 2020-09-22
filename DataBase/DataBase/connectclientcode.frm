VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form connectclientcode 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Client Code Linking"
   ClientHeight    =   1245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5445
   ControlBox      =   0   'False
   Icon            =   "connectclientcode.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   5445
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      Begin VB.CommandButton Command1 
         Caption         =   "S&how /ÚÑÖ "
         Height          =   495
         Left            =   3960
         TabIndex        =   4
         Top             =   120
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2640
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdclose 
         Caption         =   "&Close/ ÇÛáÇÞ"
         Height          =   495
         Left            =   3960
         TabIndex        =   1
         Top             =   600
         Width           =   1095
      End
      Begin MSComctlLib.ProgressBar p 
         Height          =   135
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   238
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.Label Label3 
         Caption         =   "ÇÏÎá ÑÞã ÇáÚãíá "
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   15
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Type The Client number ."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   2295
      End
   End
End
Attribute VB_Name = "connectclientcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Trim(Combo1.Text) <> "" Then
    Command1_Click
End If
End Sub

Private Sub Command1_Click()

Dim Conf As New ADODB.Connection
Dim cons As New ADODB.Connection
Dim recfin As New ADODB.Recordset
Dim recadd As New ADODB.Recordset
Dim recmar As New ADODB.Recordset

Command1.Enabled = False
cmdclose.Enabled = False

Conf.Open "Dsn=anufoxpro;uid=sa;pwd=;"
cons.Open "Dsn=finance;Uid=Sa;Pwd=;"

recadd.Open "Delete from waelclient where loguser ='" & cLogUser & "'", cons, adOpenKeyset, adLockOptimistic
recadd.Open "select * from waelclient", cons, adOpenKeyset, adLockOptimistic

If UCase(Trim(Combo1.Text)) = "ALL" Then
'this is to get all the marcusfl data table
    recmar.Open "Select * from marcusfl where left(cust_code,1) = 'O' order by cust_code", Conf, adOpenKeyset, adLockOptimistic
Else
    recmar.Open "Select * from marcusfl where cust_code = '" & Trim(Combo1.Text) & "'", Conf, adOpenKeyset, adLockOptimistic
    If recmar.BOF = True Then
        MsgBox "This Client Code is Not Existing Please Check Or Contact EDP", vbInformation, "Invalid Code"
        recmar.Close
        recadd.Close
        Conf.Close
        cons.Close
        
        Command1.Enabled = True
        cmdclose.Enabled = True
        Exit Sub
    End If
End If


p.Min = 0
p.Max = recmar.RecordCount
p.Visible = True
If recmar.BOF = False Then
    recmar.MoveFirst
    While recmar.EOF = False
    
            'open the appropriate account code from financemaster
            recfin.Open "Select * from financemaster where accountcode = '" & Trim(recmar!acctNo) & "'", cons, adOpenKeyset, adLockOptimistic
                With recadd
                    .AddNew
                        anuraja = recmar!acctNo
                        !LogUser = cLogUser
                        !cust_code = recmar!cust_code
                        !mcustcode = recmar!mcustcode
                        !category = recmar!Grp
                        If Trim(recmar!arabicname) <> "" Then
                            name1 = Trim(recmar!arabicname) & " \ "
                        End If

                        If Trim(recmar!first_name) <> "" Then
                            name1 = name1 & Trim(recmar!first_name) & " "
                        End If
                        If Trim(recmar!mid_name) <> "" Then
                            name1 = name1 & Trim(recmar!mid_name) & " "
                        End If
                        If Trim(recmar!last_name) <> "" Then
                            name1 = name1 & Trim(recmar!last_name)
                        End If
                        
                        !custname = name1
                        name1 = ""
                        If recfin.BOF = False Then
                            !AccountCode = recfin!AccountCode
                            !accountnameeng = recfin!accountnameeng
                            !accountnamearab = recfin!accountnamearab
                        End If
                    .Update
                End With
            recfin.Close
        p.Value = p.Value + 1
        rajj = Int(Val(p.Value) * 100 / Val(p.Max)) & "  % "
        Label1.caption = "Please Wait .... " & rajj
        DoEvents
        recmar.MoveNext
    Wend
End If
Label1.caption = "Type The Client Code."

On Error Resume Next
dataanu.rscom_clientcodelinking.Close
On Error GoTo 0
dataanu.com_clientcodelinking cLogUser

re_clientcodelinking.Show 1
p.Visible = False
recadd.Clone
Conf.Close
cons.Close

Command1.Enabled = True
cmdclose.Enabled = True
End Sub



Private Sub Form_Load()
Combo1.AddItem "ALL"
Combo1.Text = "ALL"
End Sub

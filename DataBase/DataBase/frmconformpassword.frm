VERSION 5.00
Begin VB.Form frmconformpassword 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Enter Password  ÇÏÎá ßáãÉ ÇáãÑæÑ"
   ClientHeight    =   1275
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4050
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   753.313
   ScaleMode       =   0  'User
   ScaleWidth      =   3802.731
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK  ãæÇÝÞ"
         Default         =   -1  'True
         Height          =   350
         Left            =   720
         TabIndex        =   4
         Top             =   720
         Width           =   1140
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel ÇáÛÇÁ"
         Height          =   350
         Left            =   2160
         TabIndex        =   3
         Top             =   720
         Width           =   1140
      End
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   240
         Width           =   2085
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ßáãÉ ÇáãÑæÑ "
         Height          =   195
         Left            =   3120
         TabIndex        =   5
         Top             =   240
         Width           =   795
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Password:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmconformpassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim countanu As Integer
Dim conpass As New ADODB.Connection
Dim recpass As New ADODB.Recordset

Private Sub cmdcancel_Click()
frmrecieptvou.stopstop1 = 1
frmpaymentvou.stopstop1 = 1
Unload Me
End Sub

Private Sub cmdok_Click()
If Trim(txtPassword.Text) <> "" Then
    conpass.Open "Dsn=Finance;Uid=Sa;Pwd=;"
    'If frmrecieptvou.fromwho = "c" Or frmpaymentvou.fromwho = "c" Then
        recpass.Open "Select * from users where role ='Cashier' and  pwd = " & "'" & Trim(txtPassword.Text) & "'", conpass, adOpenKeyset, adLockOptimistic
            If recpass.BOF = True Then
                recpass.close
                conpass.close
                countanu = countanu + 1
    
                        MsgBox "Invalid Password or Password Expired Contact Your System Administrator", , "Wrong Password"
                        txtPassword.SetFocus
                        SendKeys "{Home}+{End}"
            Else
                
                recpass.close
                conpass.close
                Unload Me
            End If
'    ElseIf frmrecieptvou.fromwho = "a" Or frmpaymentvou.fromwho = "a" Then
'                recpass.Open "Select * from users where pwd = " & "'" & Trim(txtPassword.Text) & "'", conpass, adOpenKeyset, adLockOptimistic
'                    If recpass.BOF = True Then
'                        recpass.Close
'                        conpass.Close
'                        countanu = countanu + 1
'
'                                MsgBox "Invalid Password or Password Expired Contact Your System Administrator", , "Wrong Password"
'                                txtPassword.SetFocus
'                                SendKeys "{Home}+{End}"
'                    Else
'                        recpass.Close
'                        conpass.Close
'                        Unload Me
'                    End If
'    End If
Else
    countanu = countanu + 1
    MsgBox "Invalid Password or Password Expired Contact Your System Administrator", , "Wrong Password"
    txtPassword.SetFocus
    SendKeys "{Home}+{End}"

End If
    If countanu = 3 Then
        MsgBox "This Program is going to terminate contact Your System Administrator", vbInformation, "Program Termination"
        frmrecieptvou.cmdCancelclearclear_Click
        End
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
countanu = 0
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdok_Click
End If
End Sub

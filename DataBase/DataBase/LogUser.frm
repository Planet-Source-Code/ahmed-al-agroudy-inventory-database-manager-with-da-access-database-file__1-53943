VERSION 5.00
Begin VB.Form LogUser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enter User Name and Password  ÇÏÎÇá ÇÓã ÇáãÓÊÎÏã æÇÓã ÇáÓÑí"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5190
   Icon            =   "LogUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   350
      Left            =   3840
      TabIndex        =   5
      Top             =   885
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   350
      Left            =   3840
      TabIndex        =   4
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   350
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   840
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1920
      Style           =   1  'Simple Combo
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Type your  UserName and Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label Label2 
      Caption         =   "Password ÇÓã ÇáÓÑí "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   885
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "User Name ÇÓã ÇáãÓÊÎÏã "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   540
      Width           =   1815
   End
End
Attribute VB_Name = "LogUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Times As Integer
Dim rstUser As New ADODB.Recordset

Private Sub Combo1_GotFocus()
Me.Command1.Default = False
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.Text1.SetFocus
End If
End Sub

Private Sub Command1_Click()
Dim FoundUser As Boolean
If Trim(Me.Combo1) = "" Then
    mess = MsgBox("Enter User Name ÇÏÎÇá ÇÓã ÇãÓÊÎÏã  ", vbInformation + vbOKOnly, "Message ÑÓÇáÉ")
    Me.Combo1.SetFocus
    Exit Sub
End If
On Error Resume Next
rstUser.close
rstUser.Open "Select * from Users where userid = " & "'" & Trim(Me.Combo1) & "'", constring, adOpenKeyset, adLockPessimistic, cmdtable
If rstUser.EOF = True Then
    mess = MsgBox("Unauthorized user áÇíÑÎÕ ááãÓÊÎÏã  ", vbInformation + vbOKOnly, "Message ÑÓÇáÉ")
    Me.Combo1.SetFocus
    rstUser.close
    Exit Sub
  Else
    If Format(rstUser!ExpireDate, "mm/dd/yyyy") < Format(Date, "mm/dd/yyyy") And Trim(rstUser!Expires) <> "Never" Then
        msg = MsgBox("Sorry, UserID already expired", vbCritical + vbOKOnly, "Log-in Failure")
        Exit Sub
    End If
    If Trim(rstUser!Pwd) = Trim(Me.Text1) Then
       Mainform.sbStatusBar.Panels(4).Text = "User:" & Trim(Me.Combo1)
       rstUser!logintime = Time
       rstUser!logged = "Yes"
       rstUser.Update
       xRole = Trim(rstUser!role)
       If UCase(xRole) = UCase("Admin") Then
          Mainform.xuser.Enabled = True
         Else
         Mainform.xuser.Enabled = False
       End If
       LogSucess = True
       Dim cUser As String
       Dim cROle As String
       cUser = Trim(Me.Text1)
       cROle = xRole
       UserRole = xRole
       cLogUser = Trim(Me.Combo1)
       Call EnableMenu(cUser, cROle)
       
       Unload Me
      Else
      mess = MsgBox("Invalid Password ÑÞã ÇáÓÑí ÎØÇð", vbInformation + vbOKOnly, "Message ÑÓÇáÉ")
      rstUser.close
      Me.Text1.SetFocus
    End If
End If

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Text1_GotFocus()
Me.Command1.Default = True
End Sub

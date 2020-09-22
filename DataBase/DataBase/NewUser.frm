VERSION 5.00
Begin VB.Form NewUser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Register New User ÇÓã ÇáãÓÊÎÏã "
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7515
   Icon            =   "NewUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5280
      Style           =   1  'Simple Combo
      TabIndex        =   12
      Top             =   1240
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close ÛáÞ "
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
      Left            =   4920
      TabIndex        =   10
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok ãæÇÝÞ "
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
      Left            =   3360
      TabIndex        =   9
      Top             =   1800
      Width           =   1335
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2880
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1230
      Width           =   1935
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Never E&xpired  ÇÈÏÇÁ áÇíäÊåÇÁ "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   6
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox Text2 
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
      IMEMode         =   3  'DISABLE
      Left            =   2880
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   850
      Width           =   1935
   End
   Begin VB.TextBox Text1 
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
      IMEMode         =   3  'DISABLE
      Left            =   2880
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   480
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2880
      Style           =   1  'Simple Combo
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "&Duration in Days ÈÞÇÁ Ýí Çáíæã "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   11
      Top             =   780
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "&User Role æÙíÝÉ ÇáãÓÊÎÏã "
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
      Left            =   240
      TabIndex        =   7
      Top             =   1305
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "&Confirm Password ÇßÏ ÇáÇÓã ÇáÓÑí "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   945
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "&Password ÇáÓÑí "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "User &Name ÇÓã ÇáÓÑí "
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
      Left            =   240
      TabIndex        =   0
      Top             =   195
      Width           =   1815
   End
End
Attribute VB_Name = "NewUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rstUser As New ADODB.Recordset
Dim rstRoles As New ADODB.Recordset

Private Sub Check1_Click()
If Me.Check1.Value = 1 Then
    Me.Combo3.Enabled = False
    Me.Label5.Enabled = False
   Else
     Me.Combo3.Enabled = True
     Me.Label5.Enabled = True
End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.Text1.SetFocus
End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If Me.Combo3.Enabled = True Then
        Me.Combo3.SetFocus
      Else
       Me.Command1.SetFocus
 End If
End If
End Sub

Private Sub Command1_Click()
rstUser.Open "Users", constring, adOpenKeyset, adLockPessimistic, adCmdTable
X = rstUser.RecordCount
rstUser.Close
If Trim(Me.Combo1) = "" Then
    msg = MsgBox("User Name must not be blank ÇÓã ÇáãÓÊÎÏã áÇ íÚãá ", vbExclamation + vbOKOnly, "Message")
    Me.Combo1.SetFocus
    Exit Sub
End If
If Trim(Me.Text1.Text) <> Trim(Me.Text2.Text) Then
    msg = MsgBox("Confirmed Password does not match against the Password" & vbCrLf & _
                 "ÇßÏ ÇáÇÓã ÇáÓÑí æáÇ ÊÚíÏ ÇáãÍæáÉ  ", vbExclamation + vbOKOnly, "Message ÑÓÇáÉ ")
    Me.Text2.SetFocus
    Exit Sub
End If
If Me.Check1.Value = 0 And Trim(Me.Combo3.Text) = "" Then
    msg = MsgBox("Enter Duration in DaysÇÏÎÇá ÇáÈÞÇÁ Ýí ÇáÇíÇã ", vbExclamation + vbOKOnly, "Message")
    Me.Combo3.SetFocus
    Exit Sub
End If
If Trim(Me.Combo2) = "" Then
    msg = MsgBox("Select what User's Role ãä ÝÖáß ãÇåí æÙíÝÉ ÇáãÓÊÎÏã ", vbExclamation + vbOKOnly, "Message")
    Me.Combo2.SetFocus
    Exit Sub
End If

mss = MsgBox("Do you want to Register now this new User?åá ÊÑíÏ ÇáÇÔÊÑÇß Ýí åÐÇ ÇáãÓÊÎÏã ÇáÌÏíÏ  ", vbQuestion + vbYesNoCancel, "Please confirm ãä ÝÖáß ÇáÊÇßíÏ ")
If mss = 6 Then
    rstUser.Open "Select * from Users where Userid = " & "'" & Trim(Me.Combo1) & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
    If rstUser.EOF = False Then
        msg = MsgBox("User already exist!ÇáãÓÊÎÏã íßæä ÍÝÙ" & vbCrLf & _
                      "Do you want to edit this user?", vbInformation + vbYesNo, "Message")
        If msg = vbYes Then
           With rstUser
                !Userid = Me.Combo1
                !Pwd = Me.Text1
                !logintime = ""
                !logouttime = ""
                !Lastlogin = ""
                !logged = ""
                If Me.Check1.Value = 1 Then
                    !Expires = "Never"
                    !ExpireDate = ""
                  Else
                  !Expires = "Expire"
                  !ExpireDate = Date + Val(Me.Combo3)
                End If
                !role = Me.Combo2
                .Update
                msg = MsgBox("User Modified")
                Me.Combo1 = ""
                Me.Text1 = ""
                Me.Text2 = ""
                Me.Combo3 = ""
                rstUser.Close
                Exit Sub
              End With
            Else
             rstUser.Close
             Exit Sub
          End If
        Me.Combo1.SetFocus
        Exit Sub
    End If
    On Error Resume Next
    'rstUser.MoveLast
    
    On Error GoTo 0
       With rstUser
        .AddNew
        !Code = "0" & LTrim(Str(X + 1))
        !Userid = Me.Combo1
        !Pwd = Me.Text1
        !logintime = ""
        !logouttime = ""
        !Lastlogin = ""
        !logged = ""
        If Me.Check1.Value = 1 Then
            !Expires = "Never"
            !ExpireDate = ""
          Else
          !Expires = "Expire"
          !ExpireDate = Date + Val(Me.Combo3)
        End If
        !role = Me.Combo2
        .Update
        msg = MsgBox("New user added")
        rstUser.Close
        Me.Combo1 = ""
        Me.Text1 = ""
        Me.Text2 = ""
        Me.Combo3 = ""

    End With
ElseIf mss = 2 Then
    Unload Me
End If
 


End Sub

Private Sub Command2_Click()
Unload Me
rstRoles.Close
End Sub

Private Sub Form_Load()
On Error Resume Next
rstRoles.Open "UserRoles", constring, adOpenKeyset, adLockPessimistic, adCmdTable
Do Until rstRoles.EOF = True
    Me.Combo2.AddItem rstRoles!UserRoles
    rstRoles.MoveNext
Loop
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.Text2.SetFocus
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.Combo2.SetFocus
End If

End Sub

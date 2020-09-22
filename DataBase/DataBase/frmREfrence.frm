VERSION 5.00
Begin VB.Form frmREfrence 
   Appearance      =   0  'Flat
   Caption         =   "Refernce"
   ClientHeight    =   1350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   3990
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   840
      Top             =   1440
   End
   Begin VB.Frame Frame1 
      Caption         =   "Serch for Journal #"
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   3855
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2280
         TabIndex        =   4
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdSerch 
         Caption         =   "&Search"
         Height          =   375
         Left            =   3000
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Write Your Serial No Here"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.TextBox txtReference 
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "The No You Looked for"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmREfrence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SerchForSEri
Dim MyTime

Private Sub cmdSerch_Click()
MyTime = 1

Dim myno
If Text1.Text = "" Then
Exit Sub
End If
myno = Text1.Text
Dim serchSERIES2 As New ADODB.Recordset
serchSERIES2.Open "Select * from Payjournal where serno = " & "'" & myno & "'" & "", constring, adOpenDynamic, adLockOptimistic
If serchSERIES2.EOF = True Then
MsgBox "Nothing found"
Exit Sub
End If

Me.txtReference.Text = serchSERIES2!SerialNo

End Sub

Private Sub Form_Activate()
Me.cmdSerch.Default = False

SerchForSEri = FrmPayableSetup.txtForFrmRef.Text
Dim serchSERIES As New ADODB.Recordset
serchSERIES.Open "Select * from Payjournal where serialno = " & "'" & SerchForSEri & "'" & "", constring, adOpenDynamic, adLockOptimistic

If serchSERIES.EOF = True Then
Exit Sub
End If

Me.txtReference.Text = serchSERIES!serno


MyTime = 1


End Sub

Private Sub Form_Click()
MyTime = 1

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
MyTime = 1

End Sub

Private Sub Text1_Click()
MyTime = 1

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
MyTime = 1


If KeyAscii = 13 Then
Me.cmdSerch.Default = True
End If
End Sub

Private Sub Timer1_Timer()
MyTime = MyTime + 1

If MyTime > 5000 Then
Unload Me
End If
End Sub

Private Sub txtReference_Change()
MyTime = 1

End Sub

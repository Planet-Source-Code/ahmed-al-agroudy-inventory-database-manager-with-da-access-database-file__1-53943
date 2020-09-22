VERSION 5.00
Begin VB.Form frmpaymentforcredit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enter Payment Amount."
   ClientHeight    =   780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3270
   Icon            =   "frmpaymentforcredit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   780
   ScaleWidth      =   3270
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3255
      Begin VB.TextBox txtpayment 
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   1680
         TabIndex        =   1
         Top             =   240
         Width           =   1485
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Payment Amount"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   285
         Width           =   1200
      End
   End
End
Attribute VB_Name = "frmpaymentforcredit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim xdecimal As Integer

Private Sub txtpayment_GotFocus()
xdecimal = 0

End Sub

Private Sub txtpayment_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then
   xdecimal = xdecimal + 1
 End If
If xdecimal = 2 Then
  Beep
  xdecimal = 1
 txtpayment.SetFocus
 SendKeys "{Left}+{End}"
 SendKeys "{Delete}"
End If
If KeyAscii = 13 Or KeyAscii = 46 Or KeyAscii >= 48 And KeyAscii <= 57 Then
 c = KeyAscii
 
 
 Else
   If KeyAscii = 46 And xdecimal > 1 Then
       xdecimal = 0
   End If
 If txtpayment.Text <> " " Then
        xdecimal = 0

  SendKeys "{End}+{Home}"
  SendKeys "{Delete}"
  txtpayment.SetFocus

 End If
End If
  
If KeyAscii = 13 Or KeyAscii = 46 Or KeyAscii >= 48 And KeyAscii <= 57 Then
 strcheck = KeyAscii
 Else
  
  SendKeys "{End}+{Home}"
  SendKeys "{Delete}"
 txtpayment.Text = ""
  Beep
End If

If KeyAscii = 13 And Trim(txtpayment.Text) <> "" Then
frmcreditnote.transferamount = Trim(txtpayment.Text)
Unload Me
Exit Sub
End If

If KeyAscii = 13 And Trim(txtpayment.Text) = "" Then
frminvoice.transferamount = 0
Unload Me
End If

End Sub

Private Sub txtpayment_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
frmcreditnote.transferamount = 0

    Unload Me
End If
End Sub

Private Sub txtpayment_LostFocus()
xdecimal = 0

End Sub

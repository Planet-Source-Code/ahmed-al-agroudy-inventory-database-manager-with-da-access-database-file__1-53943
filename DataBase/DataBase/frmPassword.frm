VERSION 5.00
Begin VB.Form frmPassword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enter Your Password"
   ClientHeight    =   930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4785
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   930
   ScaleWidth      =   4785
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   325
      Left            =   3600
      TabIndex        =   2
      Top             =   120
      Width           =   1000
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Left            =   3600
      TabIndex        =   3
      ToolTipText     =   "Add  New Entry"
      Top             =   480
      Width           =   1000
   End
   Begin VB.TextBox txtPrepBy 
      Height          =   285
      Left            =   1200
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtBuffer 
      Height          =   285
      Left            =   2760
      TabIndex        =   6
      Text            =   "Text3"
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtPassword 
      Height          =   325
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox txtUserId 
      Height          =   325
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User ID"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CON1 As ADODB.Connection
Dim rstpay As ADODB.Recordset
Dim rstLog  As ADODB.Recordset

Private Sub cmdNew_Click()
Unload Me
End Sub

Private Sub Command1_Click()
txtPassword_LostFocus1

End Sub

Private Sub Command1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtPassword_LostFocus1
End If

End Sub

Private Sub Form_Load()
Dim CON1 As ADODB.Connection
Set CON1 = New ADODB.Connection
Set rstLog = New ADODB.Recordset
Set rstpay = New ADODB.Recordset

'txtPassword.SetFocus
conStr = "Provider=MSDASQL;DSN=Finance;UID=sa; PWD=;"
CON1.Open conStr
rstpay.Open "Select * from payablesetup", CON1, adOpenDynamic, adLockOptimistic
rstLog.Open "Select * from newlog", CON1, adOpenDynamic, adLockOptimistic
End Sub

Private Sub txtPassword_GotFocus()
Me.Command1.Default = True
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command1.SetFocus
End If
End Sub

Private Sub txtPassword_LostFocus1()

'This is to Delete the Records from the ListView1-------------------------------------------------------------------
If Me.txtBuffer = "Delete" Then
Unload Me
xyes = MsgBox("Are You Sure You Want to Delete", vbYesNo + vbQuestion, "Deleting record")
If xyes = vbYes Then

'Delete items in the ListV
        If FrmPayableSetup.ListView1.ListItems.Count = 0 Then
        frmMenu.del.Enabled = False
        Exit Sub
        End If
  xindex = FrmPayableSetup.ListView1.SelectedItem.Index
   xItem = FrmPayableSetup.ListView1.SelectedItem
     
        FrmPayableSetup.ListView1.ListItems.Remove xindex

'Delete From the File Permanently(Put the Deletemark LATER)
    
 Dim Delx As New ADODB.Recordset
 Delx.Open "select * from payablesetup where serialno = " & "'" & xItem & "'" & "", constring, adOpenDynamic, adLockOptimistic
 While Delx.EOF = False
  Delx!deletemark = 1
  Delx!DeleUser = FrmPayableSetup.CmbPrepBy
  MsgBox "Records Deleted Successfully", vbInformation, "Conformation"
 Delx.MoveNext
  Wend
 
    
    
    
'If rstpay.EOF = False Then
'rstpay.MoveFirst
'End If
'
'        While rstpay.EOF = False
'
'        If xItem = (rstpay!serialno) Then
'        'rstpay.Delete
'        rstpay!DeleteMark = 1
'        rstpay!Deleuser = FrmPayableSetup.CmbPrepBy
'        MsgBox "Records deleted"
'
'        End If
'
'  rstpay.MoveNext
'  Wend
' Unload Me
End If
Exit Sub
End If 'This is the end for Deletion
'-------------------------------------------------------------------------------------------------

Dim varBuf As String
varBuf = txtBuffer.Text


On Error Resume Next
If Trim(varBuf) = "Save" Or Trim(varBuf) = "Update" Or Trim(varBuf) = "UpdateList3" Or "UpdateList6" Or Trim(varBuf) = "Payable cancelation" Or Trim(varBuf) = "UpdateItemList2" Or Trim(varBuf) = "DeletePurchase" Then 'This is for SAVE and UPDATE
'This is to Defend Users When SAVE records & Delete
'It will Cross Check with the combo "Prepaired By" and this is the User id also and this will check the password at the NewLog File

Dim UI
Dim Pw
Dim found
Dim xnext
found = 0
xnext = 0
UI = txtUserId.Text
Pw = txtPassword.Text

rstLog.MoveFirst


While rstLog.EOF = False

If rstLog!Userid = Me.txtPrepBy Then   'check the user id
       xnext = "ok"

        If rstLog!Password = Pw Then
        found = "ok"
        End If
End If

'Exit Sub
rstLog.MoveNext
Wend
End If

If xnext = 0 Then 'User Id is not there
MsgBox "You are not the Authorised Person ", vbCritical, "Password"
Exit Sub
End If

        
        If found = 0 Then 'Password is incorrect
        MsgBox "Incorrect Password Try Again", vbExclamation, "Password"
        txtUserId.SetFocus
        Exit Sub
        End If
        
If txtBuffer.Text = "Save" Then 'This is for SAVE
FrmPayableSetup.saveme
Unload Me

ElseIf txtBuffer.Text = "SaveCr" Then
FrmPaymentAnalysis.saveme2
Unload Me

ElseIf txtBuffer.Text = "Update" Then 'This is for UPDATE Payable SetUp
FrmPayableSetup.UpdateMe

ElseIf txtBuffer.Text = "UpdateList3" Then 'This is for UPDATE XPayable(List3)
FrmPaymentAnalysis.UpdateList3Xpay

ElseIf txtBuffer.Text = "UpdateList6" Then 'This is for UPDATE XReceipt(List6)
FrmPaymentAnalysis.UpdateList6XRecei

ElseIf txtBuffer.Text = "UpdateItemList2" Then 'This is for UPDATE XReceipt(List6)
frmPurchaseSetup.UpdateList2Item

ElseIf txtBuffer.Text = "DeletePurchase" Then 'This is for UPDATE XReceipt(List6)
frmPurchaseSetup.DeleteMe

ElseIf txtBuffer.Text = "Payable cancelation" Then 'This is for PAYABLE CANCELLATAION
FrmPayableSetup.PayableCancellation
End If
Unload Me
End Sub

Private Sub txtUserId_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtPassword.SetFocus
End If
End Sub

VERSION 5.00
Begin VB.Form RenArabName 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rename  ÇÚÇÏÉ ÇáÊÓãíÉ "
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6405
   Icon            =   "RenameArabName.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   6405
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel ÇáÛÇÁ"
      Height          =   350
      Left            =   4080
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok ãæÇÝÞ"
      Height          =   350
      Left            =   2760
      TabIndex        =   6
      Top             =   1320
      Width           =   1215
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1920
      RightToLeft     =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   5
      Top             =   810
      Width           =   2775
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1920
      Locked          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   3
      Top             =   470
      Width           =   2775
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1920
      Locked          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label6 
      Caption         =   "ÇÓã ÇáÍÓÇÈ ÈÇáÚÑÈí "
      Height          =   255
      Left            =   4800
      TabIndex        =   10
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "ÇÓã ÇáÍÓÇÈ ÈÇäÌáíÒí "
      Height          =   255
      Left            =   4800
      TabIndex        =   9
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "ÑÞã  ÇáÍÓÇÈ "
      Height          =   255
      Left            =   4200
      TabIndex        =   8
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "AccountName Arab"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "AccountName Eng"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Account No."
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "RenArabName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Newstring As String
Dim oldstring As String

Private Sub Check1_Click()
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.Command1.SetFocus
End If
End Sub

Private Sub Combo3_LostFocus()
Newstring = Me.Combo3
End Sub

Private Sub Command1_Click()
Newstring = Me.Combo3
AcctCode = Me.Combo1
Dim rstEdit As New ADODB.Recordset
Dim rstEdit1 As New ADODB.Recordset
Dim xClass As New HabitatClass
Dim xtable As String
Dim sqltable As Boolean
Dim CON1 As New ADODB.Connection

If oldstring <> Newstring Then
  mess = MsgBox("Do you want to save changes? åá ÊÑíÏ ÍÝÙ ÇáÊÛíÑÇÊ ", vbOKCancel + vbQuestion, "Please confirm ãä ÝÖáß ÇáÊÇßíÏ ")
   If mess = vbOK Then
    On Error Resume Next
    rstEdit.Close
    If WhatColumnclick = 1 Then
     rstEdit.Open "Update FinanceMaster Set accountnamearab=" & "'" & Newstring & "'" & " Where AccountCode = " & " '" & AcctCode & "'", constring, adOpenKeyset, adLockOptimistic, adCmdText
    End If
    'On Error GoTo EditCountry
    'rstEdit.Open "Update FinanceMaster Set accountnamearab=" & "'" & Newstring & "'" & " Where AccountCode = " & " '" & AcctCode & "'", conString, adOpenKeyset, adLockOptimistic, adCmdText
   If FormNo - 1 = -0 Then
              xtable = "Country" ', conString, adOpenKeyset, adLockOptimistic, adCmdTable
              ElseIf FormNo - 1 = 0 Then
                 xtable = "TopLevel" ', conString, adOpenKeyset, adLockOptimistic, adCmdTable
              ElseIf FormNo - 1 = 1 Then
                 xtable = "Level1" ', conString, adOpenKeyset, adLockOptimistic, adCmdTable
              ElseIf FormNo - 1 = 2 Then
                 xtable = "Level2" ', conString, adOpenKeyset, adLockOptimistic, adCmdTable
              ElseIf FormNo - 1 = 3 Then
                  xtable = "Level3" ', conString, adOpenKeyset, adLockOptimistic, adCmdTable
              ElseIf FormNo - 1 = 4 Then
                 xtable = "Level4" ', conString, adOpenKeyset, adLockOptimistic, adCmdTable
              ElseIf FormNo - 1 = 5 Then
                  xtable = "Level5" ', conString, adOpenKeyset, adLockOptimistic, adCmdTable
              ElseIf FormNo - 1 = 6 Then
                 xtable = "Level6" ', conString, adOpenKeyset, adLockOptimistic, adCmdTable
    End If
    xClass.GetTables rstEdit1, CON1, xtable, constring, sqltable
    While rstEdit1.EOF = False
              If Trim(rstEdit1!AccountCode) = AcctCode Then
                  If WhatColumnclick = 1 Then
                    rstEdit1!accountnamearab = Newstring
                   Else
                    'if then user attemp to move the item from Main Acct to sub-cat or Versa
                    Dim rsCheckTrn As New ADODB.Recordset
                    rsCheckTrn.Open "select * from financemaster where accountCode=" & "'" & AcctCode & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
                    If rsCheckTrn.EOF = False Then
                       If rsCheckTrn!TotalTrans <> 0 Then
                          mess = MsgBox("AccountCode has already a transaction(s),You cannot able to change this into Sub-Category", vbExclamation + vbOKOnly, "Message")
                          rsCheckTrn.Close
                          Exit Sub
                        Else
                         rstEdit1!remarks = Newstring
                       End If
                        
                      'If not existing
'                       Dim rstITem As New ADODB.Recordset
'                       If xTAble = "Level2" Then
'                           AC=Left(Acctno,)
                     End If
                  End If
                   rstEdit1.Update
             End If
             rstEdit1.MoveNext
    Wend
    rstEdit1.Close
    CON1.Close
    
    xtable = "Select * from FinanceMaster order by AccountNameEng"
    sqltable = True
    xClass.GetTables rstEdit1, CON1, xtable, constring, sqltable
    While rstEdit1.EOF = False
          If Trim(rstEdit1!AccountCode) = AcctCode Then
             If WhatColumnclick = 1 Then
                rstEdit1!accountnamearab = Newstring
               Else
                 rstEdit1.Delete
             End If
             rstEdit1.Update
          End If
          rstEdit1.MoveNext
    Wend
    rstEdit1.Close
    CON1.Close
    On Error Resume Next
    Newform.ListView1.SelectedItem.SubItems(WhatColumnclick) = Newstring
    Unload Me
    If oldstring <> Newstring Then
      If WhatColumnclick = 1 Then
         mess = MsgBox("Successfully rename ÇÚÇÏÉ ÇáÊÓãíÉ äÇÌÍÉ ", vbExclamation + vbOKOnly, "Message")
        Else
         mess = MsgBox("Successfully moved to as a Sub-Category", vbExclamation + vbOKOnly, "Message")
         Call Refresh
      End If
      Exit Sub
    End If
   Else
    Exit Sub
   End If
End If


'EditCountry:
'X = Err.Description
'If X <> 0 Then
'    On Error Resume Next
'    'rstEdit.Close
'    rstEdit.Open "Update FinanceMaster Set accountnamearab=" & "'" & Newstring & "'" & " Where AccountCode = " & " '" & AcctCode & "'", conString, adOpenKeyset, adLockOptimistic, adCmdText
'End If
'If FormNo - 1 = -0 Then
'
'       xTAble = "Country" ', conString, adOpenKeyset, adLockOptimistic, adCmdTable
'              ElseIf FormNo - 1 = 0 Then
'                 xTAble = "TopLevel" ', conString, adOpenKeyset, adLockOptimistic, adCmdTable
'              ElseIf FormNo - 1 = 1 Then
'                 xTAble = "Level1" ', conString, adOpenKeyset, adLockOptimistic, adCmdTable
'              ElseIf FormNo - 1 = 2 Then
'                 xTAble = "Level2" ', conString, adOpenKeyset, adLockOptimistic, adCmdTable
'              ElseIf FormNo - 1 = 3 Then
'                  xTAble = "Level3" ', conString, adOpenKeyset, adLockOptimistic, adCmdTable
'              ElseIf FormNo - 1 = 4 Then
'                 xTAble = "Level4" ', conString, adOpenKeyset, adLockOptimistic, adCmdTable
'              ElseIf FormNo - 1 = 5 Then
'                  xTAble = "Level5" ', conString, adOpenKeyset, adLockOptimistic, adCmdTable
'              ElseIf FormNo - 1 = 6 Then
'                 xTAble = "Level6" ', conString, adOpenKeyset, adLockOptimistic, adCmdTable
'    End If
'    xClass.GetTables rstEdit1, con1, xTAble, conString, sqltable
'    While rstEdit1.EOF = False
'              If Trim(rstEdit1!AccountCode) = AcctCode Then
'                  If WhatColumnclick = 1 Then
'                    rstEdit1!accountnamearab = Newstring
'                   Else
'                    xx = rstEdit1.Source
'                    rstEdit1!remarks = Newstring
'                  End If
'                   rstEdit1.Update
'             End If
'             rstEdit1.MoveNext
'    Wend
'    rstEdit1.Close
'    con1.Close
'    'rstEdit.Open "FinanceMaster", conString, adOpenKeyset, adLockOptimistic, adCmdTable
'    xTAble = "Select * from FinanceMaster order by AccountNameEng"
'    sqltable = True
'    xClass.GetTables rstEdit1, con1, xTAble, conString, sqltable
'    While rstEdit1.EOF = False
'          If Trim(rstEdit1!AccountCode) = AcctCode Then
'             If WhatColumnclick = 1 Then
'                rstEdit1!accountnamearab = Newstring
'               Else
'                 rstEdit1!remarks = Newstring
'             End If
'             rstEdit1.Update
'          End If
'          rstEdit1.MoveNext
'    Wend
'    rstEdit1.Close
'    con1.Close
'    On Error Resume Next
'    Newform.ListView1.SelectedItem.SubItems(WhatColumnclick) = Newstring
'    Unload Me
'If oldstring <> Newstring Then
'   mess = MsgBox("Successfully rename ÇÚÇÏÉ ÇáÊÓãíÉ äÇÌÍÉ ", vbExclamation + vbOKOnly, "Message")
'End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Activate()
'Me.Combo3.SetFocus
oldstring = Me.Combo3
End Sub


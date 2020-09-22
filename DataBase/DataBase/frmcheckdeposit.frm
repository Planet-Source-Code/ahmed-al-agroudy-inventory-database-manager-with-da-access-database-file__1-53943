VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmcheckdeposit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Check Deposit  ÇáÔíßÇÊ ÇáãæÏÚÉ "
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10935
   Icon            =   "frmcheckdeposit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   10935
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   3600
      Top             =   4080
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "C&ancel "
      Height          =   375
      Left            =   8640
      TabIndex        =   1
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "&Save "
      Height          =   375
      Left            =   9840
      TabIndex        =   0
      Top             =   4680
      Width           =   1095
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4695
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   8281
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial (Arabic)"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Cheque No."
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Due Date"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Amount"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Receipt No"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Customer Code & Customer Name"
         Object.Width           =   9349
      EndProperty
   End
   Begin VB.Label lblchecks 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   3360
      TabIndex        =   6
      Top             =   4710
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Selected AmountÇÌãÇáí ÇáãÈáÛ "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3720
      TabIndex        =   5
      Top             =   4800
      Visible         =   0   'False
      Width           =   2760
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   7920
      TabIndex        =   4
      Top             =   4800
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Selected Chesks Are : - ÇáÔíßÇÊ ÇáãÎÊÇÑÉ  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   3
      Top             =   4800
      Visible         =   0   'False
      Width           =   3345
   End
End
Attribute VB_Name = "frmcheckdeposit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public constring As String
Dim myclass As New HabitatClass
Dim sqltable As Boolean
Dim xtable As String
Dim CON1 As New ADODB.Connection
Dim recvou As New ADODB.Recordset
Public showclick As Integer
Dim closeclose As Integer

Private Sub cmdcancel_Click()
If cmdcancel.caption = "Close" Then
    Unload Me
    Exit Sub
End If

If MsgBox("Are You Sure You Want to exit With Out Save", vbYesNo, "Discard") = vbYes Then
    xtable = "delete from checkdeposittemp where paymentno = '" & Trim(frmpaymentvou.txtreceiptnumber.Text) & "'"
    sqltable = True
    Dim recfind As New ADODB.Recordset
    myclass.GetTables recfind, CON1, xtable, constring, sqltable
    frmpaymentvou.cmbpaymenttype.ListIndex = 0
    frmpaymentvou.comsetmode.ListIndex = 0
    CON1.Close
    checknumberanddate = ""
    Unload Me
End If
End Sub


Private Sub cmdsave_Click()

i = 1
For i = 1 To ListView1.ListItems.Count
    If ListView1.ListItems(i).Checked = True Then
        countcountonce = 1
    End If
Next

If countcountonce <> 1 Then
    MsgBox "Please Choose The Checks to Save", vbInformation, "Empty Data"
    countcountonce = 0
    Exit Sub
End If
countcountonce = 0

Dim recdep As New ADODB.Recordset
'this is to clear the checkdeposittemp table
sqltable = True
xtable = "delete checkdeposittemp where paymentno = '" & Trim(frmpaymentvou.txtreceiptnumber.Text) & "'"
myclass.GetTables recdep, CON1, xtable, constring, sqltable
CON1.Close
'end clear

sqltable = True
xtable = "select * from checkdeposittemp"
myclass.GetTables recdep, CON1, xtable, constring, sqltable
i = 1
If ListView1.ListItems.Count > 0 Then
    For i = 1 To ListView1.ListItems.Count
            recdep.AddNew
            recdep!CheckNo = ListView1.ListItems(i).Text
            If Trim(ListView1.ListItems(i).ListSubItems(1).Text) = "No Date" Then
            Else
               recdep!duedate = Trim(ListView1.ListItems(i).ListSubItems(1).Text)
            End If
            recdep!amount = Trim(ListView1.ListItems(i).ListSubItems(2).Text)
            recdep!receiptno = Trim(ListView1.ListItems(i).ListSubItems(3).Text)
            recdep!custno = Trim(ListView1.ListItems(i).ListSubItems(4).Text)
            recdep!paymentno = Trim(frmpaymentvou.txtreceiptnumber.Text) ' this is the payment no
            recdep!TRansDate = Format(Date, "mm/dd/yyyy")
            If ListView1.ListItems(i).Checked = True Then
                recdep!Selected = "1"
            End If
            recdep.Update
    Next
End If
recdep.Close
CON1.Close
'end save to the checkdeposittemp table


'sqltable = True
Dim total As Currency
i = 1
For i = 1 To ListView1.ListItems.Count
    If ListView1.ListItems(i).Checked = True Then
        total = total + Trim(ListView1.ListItems(i).ListSubItems(2).Text)
        If Trim(checknumberanddate) = "" Then
            checknumberanddate = Trim(ListView1.ListItems(i).Text) & " - " & Format(Trim(ListView1.ListItems(i).ListSubItems(1).Text), "dd/mm/yyyy") & ", "
        Else
            checknumberanddate = checknumberanddate & Trim(ListView1.ListItems(i).Text) & " - " & Format(Trim(ListView1.ListItems(i).ListSubItems(1).Text), "dd/mm/yyyy") & ", "
        End If
    End If
Next
frmpaymentvou.txtdebitamt.Text = ""
frmpaymentvou.txtdebitamt.Text = Format(total, "###############0.#0")
frmpaymentvou.txtdebitamt.Enabled = False
frmpaymentvou.checknumberanddate = checknumberanddate
frmpaymentvou.Changeamount
Unload Me
End Sub


Private Sub Form_Activate()
cmdcancel.SetFocus
End Sub

Private Sub Form_Load()
constring = "Dsn=finance;uid=sa;pwd="
sqltable = True
'to take the value from checkdeposittemp table
Dim recdep As New ADODB.Recordset
xtable = "select * from checkdeposittemp where paymentno = '" & Trim(frmpaymentvou.txtreceiptnumber.Text) & "'"
myclass.GetTables recdep, CON1, xtable, constring, sqltable
i = 1
    If recdep.BOF = False Then
            While i <= recdep.RecordCount
                ListView1.ListItems.Add , , Trim(recdep!CheckNo)
                If IsNull(recdep!duedate) = True Then
                    ListView1.ListItems(i).ListSubItems.Add , , "No Date"
                Else
                    ListView1.ListItems(i).ListSubItems.Add , , recdep!duedate
                End If
                ListView1.ListItems(i).ListSubItems.Add , , Format(recdep!amount, "##############0.#0")
                ListView1.ListItems(i).ListSubItems.Add , , recdep!receiptno
                ListView1.ListItems(i).ListSubItems.Add , , recdep!custno
               ' ListView1.ListItems(i).ListSubItems.Add , , recdep!Paymode
                If Trim(recdep!Selected) = "1" Then
                    ListView1.ListItems(i).Checked = True
                End If
                i = i + 1
                addsuccess = 1
                recdep.MoveNext
            Wend
    End If
        recdep.Close
        CON1.Close
    If showclick = 1 Then
        ListView1.Enabled = False
        showclick = 0
        cmdcancel.caption = "Close"
        cmdcancel.Left = 9960
        cmdsave.Visible = False
        Exit Sub ' this will stop here
    End If
    
    If addsuccess = 1 Then
        addsuccess = 0
        Exit Sub
    End If
    
'if checkdeposittemp dont have any detalils this runs
On Error Resume Next
recdep.Close
CON1.Close
On Error GoTo 0
sqltable = True

If Val(Mid(Trim(frmpaymentvou.cmbpaymenttype.Text), 1, 2)) = 3 Then
    xtable = "select * from vouchers where (left(paymode,2) = '03' or left(paymode,2) = '10') and " & _
    "deleted = '0' and deposit = '0' and " & _
    "svoucher = 'Collections' order by chkdue "
End If

myclass.GetTables recvou, CON1, xtable, constring, sqltable

i = 1
If recvou.BOF = False Then
While i <= recvou.RecordCount
        ListView1.ListItems.Add , , Trim(recvou!moderef)
        If IsNull(recvou!chkdue) = True Then
            ListView1.ListItems(i).ListSubItems.Add , , "No Date"
        Else
            ListView1.ListItems(i).ListSubItems.Add , , recvou!chkdue
        End If
        On Error GoTo 0
        ListView1.ListItems(i).ListSubItems.Add , , Format(recvou!receiptamount, "##############0.#0")
        ListView1.ListItems(i).ListSubItems.Add , , recvou!receiptno
        ListView1.ListItems(i).ListSubItems.Add , , recvou!custno & "  " & recvou!custname
        'ListView1.ListItems(i).ListSubItems.Add , , recvou!Paymode
    i = i + 1
    recvou.MoveNext
Wend
End If
recvou.Close
CON1.Close

End Sub

Private Sub lblchecks_Click()

End Sub

Private Sub ListView1_Click()
Timer1.Interval = 10
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
ListView1.SortKey = ColumnHeader.Index - 1
ListView1.Sorted = True

End Sub

Private Sub ListView1_DblClick()
Timer1.Interval = 10
End Sub


Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Timer1.Interval = 10
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
Timer1.Interval = 10
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
Timer1.Interval = 10
cmdsave.SetFocus
End Sub

Private Sub Timer1_Timer()
Dim total As Currency
lblchecks.caption = 0
i = 1

For i = 1 To ListView1.ListItems.Count
    If ListView1.ListItems(i).Checked = True Then
        If Val(lblchecks.caption) + 1 > 1 Then
            ListView1.ListItems(i).Checked = False
            MsgBox "You Can Select Only One Check For One Voucher" & vbCrLf & "If You Want; You Can Make Another Voucher", vbInformation, "Exceeded"
            Exit Sub
        End If
       total = total + Trim(ListView1.ListItems(i).ListSubItems(2).Text)
       lblchecks.caption = Val(lblchecks.caption) + 1
    End If
Next

Label2.caption = "  " & FormatNumber(total, 2) & "  Pounds.  "
Timer1.Interval = 0
End Sub

VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frminvoice 
   Caption         =   "Csutomers unpaid  Invoices"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9960
   Icon            =   "frminvoice.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6540
   ScaleWidth      =   9960
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Left            =   0
      TabIndex        =   10
      Top             =   6000
      Width           =   9855
      Begin VB.CommandButton cmdcancel 
         Caption         =   "C&ancel"
         Height          =   375
         Left            =   7920
         TabIndex        =   13
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdclose 
         Caption         =   "&Save"
         Height          =   375
         Left            =   8880
         TabIndex        =   12
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdautodistribute 
         Caption         =   "Distribute"
         Height          =   375
         Left            =   1680
         TabIndex        =   11
         Top             =   120
         Width           =   855
      End
      Begin VB.Label lblbalance 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5520
         TabIndex        =   16
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Unpaid Amount ÇáÑÕíÏ ÇáãÊÈÞí "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3840
         TabIndex        =   15
         Top             =   120
         Width           =   1845
      End
      Begin VB.Label l8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Auto Distribute ÇáÓÏÇÏ ÇáÊáÞÇÆí "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   240
         TabIndex        =   14
         Top             =   0
         Width           =   1440
      End
   End
   Begin VB.TextBox Text2 
      DataField       =   "cusnumber"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1680
      TabIndex        =   9
      Text            =   "Text2"
      Top             =   2400
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   480
      Top             =   3120
      Visible         =   0   'False
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=FINANCE"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "FINANCE"
      OtherAttributes =   ""
      UserName        =   "sa"
      Password        =   ""
      RecordSource    =   "tempinvoice"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   600
      Left            =   7320
      Top             =   2760
   End
   Begin VB.TextBox Text1 
      DataField       =   "INVC_NO"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   4560
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   2640
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtcustcode 
      Height          =   375
      Left            =   7680
      TabIndex        =   0
      Text            =   "O000019"
      Top             =   1920
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5415
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   9551
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arabic Transparent"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Invoice No"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Record Date          Due Date"
         Object.Width           =   4586
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Amount"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Unpaid"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Applied"
         Object.Width           =   3175
      EndProperty
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      Height          =   255
      Left            =   7920
      TabIndex        =   7
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total amount   ÇÌãÇáí ÇáãÈáÛ ÇáãÏÝæÚ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5280
      TabIndex        =   6
      Top             =   240
      Width           =   3075
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Un Applied Amount  ÇáÑÕíÏ ÇáÛíÑ ãÓÊÛá "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   0
      Width           =   7095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name   ÇÓã ÇáÚãíá "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   2415
   End
   Begin VB.Menu xfile 
      Caption         =   "file"
      Visible         =   0   'False
      Begin VB.Menu edit 
         Caption         =   "Edit Amount   ÝÇÖÉ ÇáãÌãæÚ "
      End
      Begin VB.Menu dash 
         Caption         =   "-"
      End
      Begin VB.Menu delete 
         Caption         =   "Delete Amount  ÇáÛÇÇÁ ÇáãÌãæÚ "
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "frminvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public anucustcode As String
Public transferamount As Currency
Public xx As Integer
Public activeform As Integer
Public validatedebitamount As Currency
Public totalvalidatedebitamount As Currency
Public selectclick As Integer
Public receiptno As String
Dim con4 As New ADODB.Connection
Dim rectemp22 As New ADODB.Recordset
Dim recforunapplied As New ADODB.Recordset

Private Sub cmdautodistribute_Click()
ii = ListView1.ListItems.Count
i = 1
Dim aamount As Currency
aamount = Val(Label5.caption)
Do While i <= ii

    If ListView1.ListItems(i).ListSubItems(3).Text <= Val(aamount) Then
       transferamount = Val(ListView1.ListItems(i).ListSubItems(3).Text)
       aamount = aamount - transferamount
        ' totalvalidatedebitamount is to calculate the total debit when you press the enter
        totalvalidatedebitamount = totalvalidatedebitamount + transferamount
        If transferamount > 0 Then 'this amount directly from the payment table
            ListView1.ListItems(i).ListSubItems(4).Text = Format(Val(transferamount) + Val(ListView1.ListItems(i).ListSubItems(4).Text), "###########0.#0")
            ListView1.ListItems(i).ListSubItems(3).Text = Format(Val(ListView1.ListItems(i).ListSubItems(3).Text) - Val(transferamount), "###########0.#0")
            'ListView1.ListItems(selectclick).ListSubItems(5).Text = Format(Val(ListView1.ListItems(selectclick).ListSubItems(3).Text), "###########0.#0")
        End If
    Else
        transferamount = aamount
        ' totalvalidatedebitamount is to calculate the total debit when you press the enter
        totalvalidatedebitamount = totalvalidatedebitamount + transferamount
        If transferamount > 0 Then 'this amount directly from the payment table
            ListView1.ListItems(i).ListSubItems(4).Text = Format(Val(transferamount) + Val(ListView1.ListItems(i).ListSubItems(4).Text), "###########0.#0")
            ListView1.ListItems(i).ListSubItems(3).Text = Format(Val(ListView1.ListItems(i).ListSubItems(3).Text) - Val(transferamount), "###########0.#0")
            'ListView1.ListItems(selectclick).ListSubItems(5).Text = Format(Val(ListView1.ListItems(selectclick).ListSubItems(3).Text), "###########0.#0")
            Exit Do
        End If
    End If
     i = i + 1
Loop


End Sub

Private Sub cmdcancel_Click()
If MsgBox("Are You Sure You Want to Cancel" & vbCrLf & " ¿  åá ÇäÊ ãÊÇßÏ ááÇáÛÇÁ", vbYesNo, "Conformation") = vbYes Then
    Timer1.Interval = 0
    Unload Me
End If
End Sub

Private Sub cmdclose_Click()
If cmdclose.caption = "Close" Then
    Unload Me
    Exit Sub
End If

If Adodc1.Recordset.BOF = False Then
Adodc1.Recordset.MoveFirst
   While Adodc1.Recordset.EOF = False
        Adodc1.Recordset.MoveFirst
        Adodc1.Recordset.Delete
        Adodc1.Recordset.Requery
    Wend
End If

For savedata = 1 To ListView1.ListItems.Count
    With Adodc1.Recordset
        .AddNew
        !num = savedata
        !cusnumber = Trim(Mid(Trim(frmrecieptvou.comreceivedfrom.Text), 1, 10))
        !InvoiceNumber = Trim(ListView1.ListItems(savedata).Text)
        !InvDate = Trim(ListView1.ListItems(savedata).ListSubItems(1).Text) 'first 9 characters receiptdate and others duedates
        !amount = Val(Trim(ListView1.ListItems(savedata).ListSubItems(2).Text))
        !unpaid = Val(Trim(ListView1.ListItems(savedata).ListSubItems(3).Text))
        !Applied = Val(Trim(ListView1.ListItems(savedata).ListSubItems(4).Text)) ' = "", 0, Trim(ListView1.ListItems(savedata).ListSubItems(4).Text))
        !unappliedbalance = Trim(Label5.caption)
        !display = Right(Trim(ListView1.ListItems(savedata).ListSubItems(1).Text), 2)
        .Update
    End With
Next
frmrecieptvou.txtallinvoice.Text = ""

For alli = 1 To ListView1.ListItems.Count
    If Val(Trim(ListView1.ListItems(alli).ListSubItems(4).Text)) > 0 Then
        frmrecieptvou.txtallinvoice.Text = frmrecieptvou.txtallinvoice.Text & ", " & Trim(ListView1.ListItems(alli).Text)
    End If
Next
frmrecieptvou.txtdebitamt.Enabled = False
Unload Me
End Sub

Private Sub delete_Click()
totalvalidatedebitamount = totalvalidatedebitamount - Val(Trim(ListView1.ListItems(selectclick).ListSubItems(4).Text))
ListView1.ListItems(selectclick).ListSubItems(3).Text = Format(Val(ListView1.ListItems(selectclick).ListSubItems(3).Text) + Val(Trim(ListView1.ListItems(selectclick).ListSubItems(4).Text)))
ListView1.ListItems(selectclick).ListSubItems(4).Text = " "
'ListView1.ListItems(selectclick).ListSubItems(5).Text = Format(Val(ListView1.ListItems(selectclick).ListSubItems(3).Text), "###########0.#0")

End Sub

Private Sub edit_Click()
ListView1_DblClick
End Sub

Private Sub Form_Activate()

If activeform = 1 Then
recforunapplied.Open "Select * from tempagaintsinvoice where invoiceno = 'Un Applied Amount' and receiptno=" & "'" & Trim(frmrecieptvou.txtreceiptnumber.Text) & "'", con4, adOpenKeyset, adLockOptimistic

Dim invcheck As New ADODB.Recordset
invcheck.Open "Select * from tempagaintsinvoice where invoiceno = 'Invoice SubTotal' and receiptno= " & "'" & Trim(frmrecieptvou.txtreceiptnumber.Text) & "'" & " and Applied = " & Trim(frmrecieptvou.txtdebitamt.Text), con4, adOpenKeyset, adLockOptimistic

If invcheck.BOF = False And recforunapplied.BOF = True Then
    invcheck.Close
    recforunapplied.Close
    frmrecieptvou.xxxyyy
    Unload Me
    MsgBox "All Payments Settled Successfully  for this Customer", vbInformation, "No Amount to Apply"
    Exit Sub
End If

If recforunapplied.BOF = False Then
    validatedebitamount = Val(recforunapplied!Applied)
    Label5.caption = validatedebitamount
Else
    validatedebitamount = Val(frmrecieptvou.txtdebitamt.Text) ' for check the total applied
End If

Label4.caption = Format(Val(frmrecieptvou.txtdebitamt.Text), "###,###,###,##0.#0")
recforunapplied.Close
Label2.caption = frmrecieptvou.comreceivedfrom.Text

    If Adodc1.Recordset.BOF = False Then ' if you enter the list already this will load the details from that
        Exit Sub
    End If

'if this your first time then this will executed
Dim recdelfin As New ADODB.Recordset ' this is from foxpro data
Dim decon As New ADODB.Connection ' this is to connect the foxpro dsn
decon.Mode = adModeShareDenyNone
decon.Open "Dsn=anufoxpro;uid=sa;pwd=;"
recdelfin.Open "select * from SJMASTER where cust_code = " & "'" & anucustcode & "'" & " and unpaidamt > 0 order by delv_date,invc_date", decon, adOpenKeyset, adLockOptimistic
If recdelfin.BOF = False Then
    recdelfin.MoveFirst
End If
xx = 0
While recdelfin.EOF = False
    'If anucustcode = Trim(recdelfin!cust_code) Then
        xx = xx + 1
        ListView1.ListItems.Add , , Trim(recdelfin!invc_no)
        Dim invc_date As Date, delv_date
        invc_date = Trim(recdelfin!invc_date)
        delv_date = Trim(recdelfin!delv_date)
        If recdelfin!LDIsp = False Then
            idisp = "F"
        Else
            idisp = "T"
        End If
        ListView1.ListItems(xx).ListSubItems.Add , , Format(invc_date, "dd/mm/yyyy") & "          " & Format(delv_date, "dd/mm/yyyy") & "             " & idisp
        ListView1.ListItems(xx).ListSubItems.Add , , Format(Trim(recdelfin!tot_amt), "###########0.#0")
        ListView1.ListItems(xx).ListSubItems.Add , , Format(Trim(recdelfin!unpaidamt), "###########0.#0")
        If frmrecieptvou.comreceivedfrom.Enabled = True Then
        ListView1.ListItems(xx).ListSubItems.Add , , " " ' this is for applied value
        Else
        ListView1.ListItems(xx).ListSubItems.Add , , Format(Trim(recdelfin!Applied), "############0.#0")
        End If
    recdelfin.MoveNext
    Wend
recdelfin.Close
decon.Close




'this is for the debit note
decon.Open "Dsn=anufoxpro;uid=sa;pwd=;"
recdelfin.Open "Select * from credmain where left(invc_no,3) = 'ODN' and Cust_code = '" & anucustcode & "'", decon, adOpenKeyset, adLockOptimistic
If recdelfin.BOF = False Then
    recdelfin.MoveFirst
End If
rr = 0
While recdelfin.EOF = False
        rr = rr + 1
        ListView1.ListItems.Add , , Trim(recdelfin!invc_no)
        invc_date = Trim(recdelfin!invc_date)
        delv_date = Trim(recdelfin!delv_date)
        ListView1.ListItems(rr).ListSubItems.Add , , Format(invc_date, "dd/mm/yyyy") & "          " & Format(delv_date, "dd/mm/yyyy") & "             " & idisp
        ListView1.ListItems(rr).ListSubItems.Add , , Format(Trim(recdelfin!tot_amt), "###########0.#0")
        ListView1.ListItems(rr).ListSubItems.Add , , Format(Trim(recdelfin!tot_amt - recdelfin!paidamt), "###########0.#0")
        ListView1.ListItems(rr).ListSubItems.Add , , Format(Trim("0.00"), "############0.#0")
    recdelfin.MoveNext
    Wend
recdelfin.Close
decon.Close
'end debit note




If xx = 0 And rr = 0 Then
    Unload Me
    MsgBox "This Customer Have No Any Unpaid Invoice" & vbCrLf & "  ßá ÇáÝæÇÊíÑ ãÏÝæÚå", vbInformation, "No Invoice"
End If
activeform = 2
End If
End Sub

Private Sub Form_Load()
con4.Open "dsN=fINANCE;UID=SA;PWD=;"

ListView1.ListItems.Clear
totalvalidatedebitamount = 0 'just assigning value
Label2.caption = frmrecieptvou.comreceivedfrom.Text
Label4.caption = Format(Val(frmrecieptvou.txtdebitamt.Text), "###,###,###,##0.#0")
validatedebitamount = Val(frmrecieptvou.txtdebitamt.Text) ' for check the total applied
If activeform <> 0 Then ' when you in edit mode it will show you only the applied and saved
    If Adodc1.Recordset.BOF = False Then ' if you enter the list already this will load the details from that
        ListView1.ListItems.Clear
        Adodc1.Recordset.MoveFirst
        xx = 0
        While Adodc1.Recordset.EOF = False

            With Adodc1.Recordset
                xx = xx + 1
                    ListView1.ListItems.Add , , Trim(!InvoiceNumber)
                    ListView1.ListItems(xx).ListSubItems.Add , , Trim(!InvDate)
                    ListView1.ListItems(xx).ListSubItems.Add , , Format(Trim(!amount), "###########0.#0")
                    ListView1.ListItems(xx).ListSubItems.Add , , Format(Trim(!unpaid), "###########0.#0")
                    ListView1.ListItems(xx).ListSubItems.Add , , Format(IIf(Trim(!Applied) <= 0, " ", Trim(!Applied)), "###########0.#0")
                .MoveNext
            End With

        Wend
        Exit Sub
    End If
    'con2.Close
End If
If activeform = 0 Then
rectemp22.Open "select * from tempinvoice2 where receiptno=" & "'" & receiptno & "'" & " and custid = " & "'" & anucustcode & "'", con4, adOpenKeyset, adLockOptimistic

    If rectemp22.BOF = False Then

        rectemp22.MoveFirst
        xx = 0
        While rectemp22.EOF = False
            'If receiptno = Trim(rectemp22!receiptno) And anucustcode = Trim(rectemp22!custid) Then
                    xx = xx + 1
                    With rectemp22
                    ListView1.ListItems.Add , , Trim(!invoiceno)
                    ListView1.ListItems(xx).ListSubItems.Add , , Trim(!receiptdate)
                    ListView1.ListItems(xx).ListSubItems.Add , , Format(Trim(!amount), "###########0.#0")
                    ListView1.ListItems(xx).ListSubItems.Add , , Format(Trim(!unpaid), "###########0.#0")
                    ListView1.ListItems(xx).ListSubItems.Add , , Format(Trim(!Applied), "############0.#0")
                    activeform = 2
                    End With
           'End If
        rectemp22.MoveNext
        Wend
    End If
rectemp22.Close
End If

End Sub


Private Sub Form_Resize()
On Error Resume Next
ListView1.Width = Me.Width - 350
ListView1.Height = Me.Height - 1800
Frame1.Top = ListView1.Height + 700
Frame1.Width = ListView1.Width
cmdclose.Left = Frame1.Width - 975
cmdcancel.Left = cmdclose.Left - 975
XC = Val(ListView1.Width) / 5

For i = 1 To 5
    ListView1.ColumnHeaders(i).Width = XC - 72
Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
con4.Close
End Sub

Private Sub Label7_Click()

End Sub

Private Sub ListView1_Click()
If ListView1.ListItems.Count > 0 Then
selectclick = ListView1.SelectedItem.Index
End If
End Sub

Private Sub ListView1_DblClick()
If ListView1.ListItems.Count > 0 Then
frmpayment.Show 1
If transferamount > Val(ListView1.ListItems(selectclick).ListSubItems(3).Text) Then
    MsgBox "Your Payment Amount Can not More Than the Unpaid Amount" & vbCrLf & " áÇíãßä ÇáÏÝÚ ÇßËÑ ãä ÇáãÈáÛ ÇáÛíÑ ÇáãÏÝæÚ", vbInformation, "Invalid Amount"
    Exit Sub
End If

' totalvalidatedebitamount is to calculate the total debit when you press the enter
totalvalidatedebitamount = totalvalidatedebitamount + transferamount
'validatedebitamount is the total amount from cashier
If totalvalidatedebitamount > validatedebitamount Then ' validatedebitamount  is the first amount
    MsgBox "Please check Your Amount is More than Applicable Amount" & vbCrLf & "  ÊÇßÏ ãä ÇáãÏíæäíå áÇäåÇ ÇßËÑ ãä ÇáãÏÝæÚ", vbInformation, "Amount Over flow"
    totalvalidatedebitamount = totalvalidatedebitamount - transferamount
    Exit Sub
End If
If transferamount > 0 Then 'this amount directly from the payment table
    ListView1.ListItems(selectclick).ListSubItems(4).Text = Format(Val(transferamount) + Val(ListView1.ListItems(selectclick).ListSubItems(4).Text), "###########0.#0")
    ListView1.ListItems(selectclick).ListSubItems(3).Text = Format(Val(ListView1.ListItems(selectclick).ListSubItems(3).Text) - Val(transferamount), "###########0.#0")
    'ListView1.ListItems(selectclick).ListSubItems(5).Text = Format(Val(ListView1.ListItems(selectclick).ListSubItems(3).Text), "###########0.#0")
End If
End If
End Sub

Private Sub ListView1_KeyUp(KeyCode As Integer, Shift As Integer)
ListView1_Click

If KeyCode = 45 Then
    ListView1_DblClick
End If
If KeyCode = 46 And Val(ListView1.ListItems(selectclick).ListSubItems(4).Text) > 0 Then
    delete_Click
End If
End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = vbRightButton Then
    ListView1_Click
    If Val(ListView1.ListItems(selectclick).ListSubItems(4).Text) > 0 Then
        Delete.Enabled = True
    End If
    PopupMenu xfile, vbAlignRight, (ListView1.SelectedItem.Left + 6000), (ListView1.SelectedItem.Top + ListView1.Top + 250)
End If
End Sub

Private Sub Timer1_Timer()
totallistcount = 0
totallistcount = ListView1.ListItems.Count
totalvalidatedebitamount = 0
For ListCount = 1 To totallistcount
    totalbalance = totalbalance + Val(Trim(ListView1.ListItems(ListCount).ListSubItems(3).Text))
    totalvalidatedebitamount = totalvalidatedebitamount + Val(Trim(ListView1.ListItems(ListCount).ListSubItems(4).Text))
Next
lblbalance.caption = Format(totalbalance, "###,###,###,##0.#0")


Label5.caption = Format(Val(validatedebitamount - totalvalidatedebitamount), "###########0.#0")
totalbalance = 0
'Timer1.Interval = 0
End Sub


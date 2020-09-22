VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPettyCash 
   Caption         =   "Petty Cash"
   ClientHeight    =   8025
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   12330
   Icon            =   "frmPettyCashReq.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8025
   ScaleWidth      =   12330
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdSearch3 
      Caption         =   "Go"
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
      Left            =   5880
      TabIndex        =   50
      ToolTipText     =   "Enter Journal No to Search"
      Top             =   7470
      Width           =   405
   End
   Begin VB.TextBox txtSearch2 
      Height          =   325
      Left            =   4560
      TabIndex        =   49
      Top             =   7470
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7215
      Left            =   0
      TabIndex        =   4
      Top             =   120
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   12726
      _Version        =   393216
      Tab             =   1
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Cash Holder"
      TabPicture(0)   =   "frmPettyCashReq.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Petty CashList"
      TabPicture(1)   =   "frmPettyCashReq.frx":045E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame5"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Journal"
      TabPicture(2)   =   "frmPettyCashReq.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "LvwPCJournal"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin MSComctlLib.ListView LvwPCJournal 
         Height          =   6735
         Left            =   -74880
         TabIndex        =   39
         Top             =   360
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   11880
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Frame Frame5 
         Caption         =   "UnConfirmed List"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   120
         TabIndex        =   36
         Top             =   360
         Width           =   12045
         Begin VB.TextBox txttotUc 
            Height          =   325
            Left            =   10560
            TabIndex        =   37
            Top             =   2760
            Width           =   1335
         End
         Begin MSComctlLib.ListView LvwPCashUnConf 
            Height          =   2535
            Left            =   120
            TabIndex        =   38
            Top             =   240
            Width           =   11775
            _ExtentX        =   20770
            _ExtentY        =   4471
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Confirmed List"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3615
         Left            =   120
         TabIndex        =   33
         Top             =   3480
         Width           =   12045
         Begin VB.TextBox Text1 
            Height          =   325
            Left            =   10560
            TabIndex        =   34
            Top             =   3120
            Width           =   1335
         End
         Begin MSComctlLib.ListView LvwPCashConf 
            Height          =   2775
            Left            =   120
            TabIndex        =   35
            Top             =   360
            Width           =   11775
            _ExtentX        =   20770
            _ExtentY        =   4895
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Petty Cash Holder"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6615
         Left            =   -74880
         TabIndex        =   5
         Top             =   480
         Width           =   12015
         Begin VB.TextBox Text2 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#,##0.00;(#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3073
               SubFormatType   =   0
            EndProperty
            Height          =   315
            Left            =   8520
            TabIndex        =   48
            Top             =   240
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.TextBox txtCreditAmount 
            Height          =   315
            Left            =   9600
            TabIndex        =   47
            Top             =   600
            Width           =   1215
         End
         Begin VB.Frame Frame3 
            Caption         =   " List"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2655
            Left            =   120
            TabIndex        =   30
            Top             =   3840
            Width           =   11800
            Begin VB.TextBox txtTotCr 
               Height          =   325
               Left            =   10560
               TabIndex        =   42
               Top             =   2200
               Width           =   1095
            End
            Begin VB.TextBox txtTotal 
               Height          =   325
               Left            =   9360
               TabIndex        =   32
               Top             =   2200
               Width           =   1095
            End
            Begin MSComctlLib.ListView LVWPettyCash 
               Height          =   1815
               Left            =   50
               TabIndex        =   31
               Top             =   360
               Width           =   11655
               _ExtentX        =   20558
               _ExtentY        =   3201
               View            =   3
               LabelWrap       =   0   'False
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   0
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Particulars"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2415
            Left            =   120
            TabIndex        =   9
            Top             =   1200
            Width           =   11775
            Begin VB.ComboBox cmbEnryType 
               Height          =   315
               Left            =   1320
               TabIndex        =   45
               Top             =   480
               Width           =   3015
            End
            Begin VB.TextBox txtPoNo 
               Height          =   325
               Left            =   7080
               TabIndex        =   40
               Top             =   1200
               Width           =   1935
            End
            Begin VB.TextBox txtExpl 
               Height          =   325
               Left            =   1320
               TabIndex        =   17
               Top             =   1560
               Width           =   3975
            End
            Begin VB.TextBox txtInvNo 
               Height          =   325
               Left            =   7080
               TabIndex        =   16
               Top             =   840
               Width           =   1935
            End
            Begin VB.TextBox txtAmt 
               BeginProperty DataFormat 
                  Type            =   0
                  Format          =   "#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3073
                  SubFormatType   =   0
               EndProperty
               Height          =   325
               Left            =   7080
               TabIndex        =   15
               Top             =   480
               Width           =   1335
            End
            Begin VB.TextBox txtPartic 
               Height          =   325
               Left            =   1320
               TabIndex        =   14
               Top             =   1200
               Width           =   3975
            End
            Begin VB.TextBox txtPaidto 
               Height          =   325
               Left            =   1320
               TabIndex        =   13
               Top             =   840
               Width           =   3975
            End
            Begin VB.ComboBox cmbAccNo2 
               Height          =   315
               Left            =   1320
               TabIndex        =   12
               Top             =   1920
               Width           =   3015
            End
            Begin VB.ComboBox cmbAccName2 
               Height          =   315
               Left            =   7080
               TabIndex        =   11
               Top             =   1920
               Width           =   3495
            End
            Begin VB.CommandButton cmdAdd 
               Caption         =   "&Add"
               Height          =   350
               Left            =   10680
               TabIndex        =   10
               Top             =   1920
               Width           =   975
            End
            Begin MSMask.MaskEdBox mskPOdate 
               Height          =   330
               Left            =   7080
               TabIndex        =   43
               Top             =   1560
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   582
               _Version        =   393216
               MaxLength       =   10
               Format          =   "dd/mm/yyyyy"
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin VB.Label Label4 
               Caption         =   "Entry Type"
               Height          =   255
               Left            =   120
               TabIndex        =   46
               Top             =   480
               Width           =   1215
            End
            Begin VB.Label eee 
               Caption         =   "S E  Date"
               Height          =   255
               Left            =   5880
               TabIndex        =   44
               Top             =   1560
               Width           =   1215
            End
            Begin VB.Label Label3 
               Caption         =   "S E  No"
               Height          =   255
               Left            =   5880
               TabIndex        =   41
               Top             =   1200
               Width           =   1215
            End
            Begin VB.Label Label12 
               Caption         =   "Explanations"
               Height          =   255
               Left            =   120
               TabIndex        =   24
               Top             =   1560
               Width           =   1215
            End
            Begin VB.Label Label10 
               Caption         =   "Invoice No"
               Height          =   255
               Left            =   5880
               TabIndex        =   23
               Top             =   840
               Width           =   1215
            End
            Begin VB.Label Label7 
               Caption         =   "Paid To"
               Height          =   255
               Left            =   120
               TabIndex        =   22
               Top             =   840
               Width           =   1215
            End
            Begin VB.Label Label14 
               Caption         =   "Amount"
               Height          =   255
               Left            =   5880
               TabIndex        =   21
               Top             =   480
               Width           =   1215
            End
            Begin VB.Label Label9 
               Caption         =   "Particulars"
               Height          =   255
               Left            =   120
               TabIndex        =   20
               Top             =   1200
               Width           =   1215
            End
            Begin VB.Label Label2 
               Caption         =   "Account Name"
               Height          =   255
               Left            =   5880
               TabIndex        =   19
               Top             =   1920
               Width           =   1215
            End
            Begin VB.Label Label1 
               Caption         =   "Account No"
               Height          =   255
               Left            =   120
               TabIndex        =   18
               Top             =   1920
               Width           =   1215
            End
         End
         Begin VB.ComboBox CMBAccName 
            Height          =   315
            Left            =   1440
            TabIndex        =   8
            Top             =   240
            Width           =   3975
         End
         Begin VB.ComboBox CMBAccNo 
            Height          =   315
            Left            =   1440
            TabIndex        =   7
            Top             =   600
            Width           =   3015
         End
         Begin VB.ComboBox txtSerialNo 
            Height          =   315
            Left            =   8520
            TabIndex        =   6
            Top             =   240
            Width           =   1815
         End
         Begin MSMask.MaskEdBox mskDate 
            Height          =   330
            Left            =   8520
            TabIndex        =   25
            Top             =   600
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   582
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mm/yyyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label13 
            Caption         =   "Account No"
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label11 
            Caption         =   "Account Name"
            Height          =   255
            Left            =   240
            TabIndex        =   28
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label6 
            Caption         =   "Serial No"
            Height          =   255
            Left            =   7680
            TabIndex        =   27
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label5 
            Caption         =   "Date"
            Height          =   255
            Left            =   7680
            TabIndex        =   26
            Top             =   600
            Width           =   1215
         End
      End
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   11280
      TabIndex        =   3
      Top             =   7440
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   7440
      Width           =   975
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   7440
      Width           =   975
   End
   Begin VB.CommandButton CmdNew 
      Caption         =   "&New"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Click to make New Transaction"
      Top             =   7440
      Width           =   975
   End
   Begin VB.Frame Frame6 
      Height          =   550
      Left            =   3720
      TabIndex        =   51
      Top             =   7320
      Width           =   2655
      Begin VB.Label Label8 
         Caption         =   "Find"
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   240
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmPettyCash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstChartx As ADODB.Recordset
Dim rstCharty As ADODB.Recordset
Dim conn1 As ADODB.Connection
Dim Constg As ADODB.Connection
Public PrvTotlLVWPC
Public PrvTotCRlLVWPC

Dim WhatColumnNo As Integer 'nelson

Public varName5555
Dim catName As String
Dim Prevcap As String
Dim acctNo As String
 Dim Categry
 
Private Sub CMBAccName_Click()

Dim varName
X = 0
For X = 1 To Len(CMBAccName)
    If Mid(Trim(CMBAccName), X, 1) = "\" Then Exit For
     varName = varName & Mid(CMBAccName, X, 1)
Next



Dim varFindcode
Dim rstChartZ As New ADODB.Recordset


varFindName = CMBAccName.Text
rstChartZ.Open "Select * from financemaster where Accountnameeng  = " & "'" & Trim(varName) & "'" & " order by accountcode", constring, adOpenDynamic, adLockOptimistic

If rstChartZ.EOF = True Then
MsgBox "Sorry You Entered the Wrong Name", vbInformation, "Please Check the Account Name"
Exit Sub
End If


Me.CMBAccNo.Text = rstChartZ!AccountCode

End Sub

Private Sub CMBAccName_GotFocus()
Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(CMBAccName.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)

End Sub

Private Sub CMBAccName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
CMBAccName_Click
cmbEnryType.SetFocus
End If
End Sub

Private Sub cmbAccName2_Click()

Dim varName2
X = 0
For X = 1 To Len(cmbAccName2)
    If Mid(Trim(cmbAccName2), X, 1) = "\" Then Exit For
     varName2 = varName2 & Mid(cmbAccName2, X, 1)
Next


Dim varFindcode2
Dim rstChartZ2 As New ADODB.Recordset

If cmbAccName2.Text = "" Then
MsgBox "Combo Account Name is Empty", vbInformation
Exit Sub
End If
On Error Resume Next

varFindName2 = cmbAccName2.Text
rstChartZ2.Open "Select * from financemaster where Accountnameeng = " & "'" & varName2 & "'" & " order by accountcode", constring, adOpenDynamic, adLockOptimistic


Me.cmbAccNo2.Text = rstChartZ2!AccountCode
On Error GoTo 0
End Sub

Private Sub cmbAccName2_GotFocus()
Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(cmbAccName2.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)

End Sub

Private Sub cmbAccName2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEnd Then
cmdAdd_Click
cmbEnryType.SetFocus
End If
End Sub

Private Sub cmbAccName2_KeyPress(KeyAscii As Integer)


If KeyAscii = 13 Then
cmbAccName2_Click
End If

If cmdadd.caption = "Update" Then
cmdAdd_Click
End If

End Sub

Private Sub cmbAccName2_LostFocus()

Dim varName4
X = 0
For X = 1 To Len(cmbAccName2)
    If Mid(Trim(cmbAccName2), X, 1) = "\" Then Exit For
     varName4 = varName4 & Mid(cmbAccName2, X, 1)
Next



Dim rstCharZZ2 As New ADODB.Recordset
varFindName2 = cmbAccName2.Text
rstCharZZ2.Open "Select * from financemaster where Accountnameeng = " & "'" & varName4 & "'" & " order by accountcode", constring, adOpenDynamic, adLockOptimistic

If rstCharZZ2.EOF = True Then
MsgBox "You Entered the Wrong Name", vbInformation
End If
End Sub

Private Sub CMBAccNo_Click()
Dim varFindName
Dim rstCharty As New ADODB.Recordset


varFindName = CMBAccNo.Text
rstCharty.Open "Select * from financemaster where Accountcode = " & "'" & varFindName & "'" & " order by accountcode", constring, adOpenDynamic, adLockOptimistic

If rstCharty.EOF = True Then
MsgBox "Sorry You Entered the Wrong No", vbInformation, "Please Check the Account No"
Exit Sub
End If


Me.CMBAccName.Text = rstCharty!accountnameeng & "\" & rstCharty!accountnamearab

Dim catName As String
Dim Prevcap As String
Dim acctNo As String
acctNo = Trim(CMBAccNo)
Prevcap = Trim(Me.caption)
Call DisplayCats(Prevcap, acctNo, catName)
Me.caption = "Petty Cash" & "//" & catName
DrCat = catName


End Sub


Private Sub CMBAccNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
CMBAccNo_Click
txtserialNo.SetFocus
End If

End Sub

Private Sub CMBAccNo_LostFocus()
If cmbAccNo2.Text = "" Then
MsgBox "Combo Account No is Empty", vbInformation
Exit Sub
End If

End Sub

Private Sub cmbAccNo2_Click()



Dim varFindName2
Dim rstChartY2 As New ADODB.Recordset
varFindName2 = cmbAccNo2.Text
rstChartY2.Open "Select * from financemaster where Accountcode = " & "'" & Trim(varFindName2) & "'" & " order by accountcode", constring, adOpenDynamic, adLockOptimistic



'If rstChartY2.EOF = True Then
'MsgBox "Sorry You Entered the Wrong No", vbInformation, "Please Check the Account No"
'Exit Sub
'End If


On Error Resume Next
Me.cmbAccName2.Text = rstChartY2!accountnameeng & "\" & rstChartY2!accountnamearab




Dim catName As String
Dim Prevcap As String
Dim acctNo As String
acctNo = Trim(cmbAccNo2)
Prevcap = Trim(Me.caption)
Call DisplayCats(Prevcap, acctNo, catName)
Me.caption = "Petty Cash" & "//" & catName
DrCat = catName

On Error GoTo 0
End Sub

Private Sub cmbAccNo2_GotFocus()
Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(cmbAccNo2.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cmbAccNo2_KeyPress(KeyAscii As Integer)
cmbAccNo2_Click

If KeyAscii = 13 Then
'cmbAccName2.SetFocus




txtAmt.SetFocus

End If

End Sub

Private Sub cmbAccNo2_LostFocus()
Dim varFindName2
Dim rstChartY3 As New ADODB.Recordset
varFindName3 = cmbAccNo2.Text
rstChartY3.Open "Select * from financemaster where Accountcode = " & "'" & Trim(varFindName3) & "'" & " order by accountcode", constring, adOpenDynamic, adLockOptimistic



If rstChartY3.EOF = True Then
MsgBox "Sorry You Entered the Wrong No or There is No Account No", vbInformation, "Please Check the Account No"
Exit Sub
End If


'On Error Resume Next
'Me.cmbAccName2.Text = rstChartY2!accountnameeng & "\" & rstChartY2!accountnamearab


End Sub

Private Sub cmbEnryType_GotFocus()
Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(cmbEnryType.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)

End Sub

Private Sub cmbEnryType_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtPaidto.SetFocus
End If
End Sub
'
Private Sub cmdAdd_Click()
If cmbEnryType.Text = "" Then
MsgBox "Select the Combo 'Entry Type'", vbInformation, "Combo is Empty"
Exit Sub
End If

If txtPaidto.Text = "" Then
MsgBox "Select the Combo 'Pay To'", vbInformation, "Combo is Empty"
Exit Sub
End If

If txtPartic.Text = "" Then
MsgBox "Fill the TextBox 'Peticulars'", vbInformation, "TextBox is Empty"
Exit Sub
End If

If txtExpl.Text = "" Then
MsgBox "Fill the TextBox 'Explanations'", vbInformation, "TextBox is Empty"
Exit Sub
End If

If txtAmt.Text = "" Then
MsgBox "You did not Put the Amount", vbInformation, "TextBox is Empty"
Exit Sub
End If

If cmbAccNo2.Text = "" Then
MsgBox "Select the Combo 'Account No' for Particulars", vbInformation, "Combo is Empty"
Exit Sub
End If

If cmbAccName2.Text = "" Then
MsgBox "Select the Combo 'Account Name' for Particulars", vbInformation, "Combo is Empty"
Exit Sub
End If


If Me.cmdadd.caption = "&Add" And cmbEnryType.Text = "Debit Entry" Then

     Set MItem = Me.LVWPettyCash.ListItems.Add(, , Trim(txtPaidto.Text))
     MItem.SubItems(1) = Trim(txtPartic.Text)
     MItem.SubItems(2) = Trim(txtExpl.Text)
     MItem.SubItems(3) = Trim(txtInvNo.Text)
     MItem.SubItems(4) = Trim(txtPoNo.Text)
     MItem.SubItems(5) = Trim(mskPOdate.Text)

     MItem.SubItems(6) = Trim(cmbAccNo2.Text)
     MItem.SubItems(7) = Trim(cmbAccName2.Text)
     MItem.SubItems(8) = Format(Trim(txtAmt.Text), "############.#0")

 txtTotal.Text = Val(Format(txtTotal.Text, "##########.0#")) + Val(Format(txtAmt.Text, "############.#0"))

    txtPaidto.Text = ""
     txtPartic.Text = ""
     txtExpl.Text = ""
     txtAmt.Text = ""
     cmbAccName2.Text = ""
     txtInvNo.Text = ""
     'cmbAccNo2.Text = ""
txtTotal.Text = Format(txtTotal.Text, "###,###,###,###.#0")



ElseIf Me.cmdadd.caption = "&Add" And cmbEnryType.Text = "Other Credit Entry" Then

     Set MItem = Me.LVWPettyCash.ListItems.Add(, , Trim(txtPaidto.Text))
     MItem.SubItems(1) = Trim(txtPartic.Text)
     MItem.SubItems(2) = Trim(txtExpl.Text)
     MItem.SubItems(3) = Trim(txtInvNo.Text)
     MItem.SubItems(4) = Trim(txtPoNo.Text)
     MItem.SubItems(5) = Trim(mskPOdate.Text)

     MItem.SubItems(6) = Trim(cmbAccNo2.Text)
     MItem.SubItems(7) = Trim(cmbAccName2.Text)
    ' mitem.SubItems(8) = Format(Trim(txtAmt.Text), "############.#0")
     MItem.SubItems(9) = Format(Trim(txtAmt.Text), "############.#0")


'txtCreditAmount.Text = Val(Format(txtTotal.Text, "##########.0#")) - Val(Format(txtAmt.Text, "############.#0"))
 txtTotCr.Text = Val(Format(txtTotCr.Text, "##########.0#")) + Val(Format(txtAmt.Text, "############.#0"))
    txtPaidto.Text = ""
     txtPartic.Text = ""
     txtExpl.Text = ""
     txtAmt.Text = ""
     cmbAccName2.Text = ""
     txtInvNo.Text = ""
     'cmbAccNo2.Text = ""

txtTotCr.Text = Format(txtTotCr.Text, "###,###,###,###.#0")
'--------------------------------
ElseIf Me.cmdadd.caption = "Update" And cmbEnryType.Text = "Debit Entry" Then
Dim ListUpda As New ADODB.Recordset
Dim DelLVWPC, LVWPCIndex, CounDif

Mymsgx = MsgBox("Do you want to Update the ListView", vbInformation + vbYesNo, "Please Confirm")
If Mymsgx = vbYes Then
LVWPCIndex = Me.LVWPettyCash.SelectedItem.Index

Me.LVWPettyCash.ListItems.Remove (LVWPCIndex)

'CounDif = Val(frmPettyCash.txtAmt.Text) - Val(PrvTotlLVWPC)

            If Trim(Format(frmPettyCash.txtAmt.Text, "#############.#0")) < Trim(Format(PrvTotlLVWPC, "############.#0")) Then
                CounDif = Val(Format(PrvTotlLVWPC, "#############.#0")) - Val(Format(frmPettyCash.txtAmt.Text, "#############.#0"))
                txtTotal.Text = Val(Format(txtTotal.Text, "#############.#0")) - Val(CounDif)
            ElseIf Trim(Format(frmPettyCash.txtAmt.Text, "#############.#0")) > Trim(Format(PrvTotlLVWPC, "#############.#0")) Then
                CounDif = Val(Format(frmPettyCash.txtAmt.Text, "#############.#0")) - Val(Format(PrvTotlLVWPC, "#############.#0"))
                txtTotal.Text = Val(Format(txtTotal.Text, "#############.#0")) + Val(CounDif)
            End If
     
     Set MItem = Me.LVWPettyCash.ListItems.Add(, , Trim(txtPaidto.Text))
     MItem.SubItems(1) = Trim(txtPartic.Text)
     MItem.SubItems(2) = Trim(txtExpl.Text)
     MItem.SubItems(3) = Trim(txtInvNo.Text)
     MItem.SubItems(4) = Trim(txtPoNo.Text)
     MItem.SubItems(5) = Trim(mskPOdate.Text)

     MItem.SubItems(6) = Trim(cmbAccNo2.Text)
     MItem.SubItems(7) = Trim(cmbAccName2.Text)
     MItem.SubItems(8) = Format(Trim(txtAmt.Text), "############.#0")


     txtPaidto.Text = ""
     txtPartic.Text = ""
     txtExpl.Text = ""
     txtAmt.Text = ""
     cmbAccName2.Text = ""
     txtInvNo.Text = ""
     cmbAccNo2.Text = ""

Me.cmdadd.caption = "&Add"
End If
'------------------------------------
    ElseIf Me.cmdadd.caption = "Update" And cmbEnryType.Text = "Other Credit Entry" Then
    
    Mymsgx = MsgBox("Do you want to Update the ListView", vbInformation + vbYesNo, "Please Confirm")
    If Mymsgx = vbYes Then
    LVWPCIndex2 = Me.LVWPettyCash.SelectedItem.Index
    
    Me.LVWPettyCash.ListItems.Remove (LVWPCIndex2)
    
    'CounDif = Val(frmPettyCash.txtAmt.Text) - Val(PrvTotlLVWPC)
    Dim CounDifx
            If Trim(Format(frmPettyCash.txtAmt.Text, "#############.#0")) < Trim(Format(PrvTotlLVWPC, "############.#0")) Then
                CounDifx = Val(Format(PrvTotlLVWPC, "#############.#0")) - Val(Format(frmPettyCash.txtAmt.Text, "#############.#0"))
                txtTotCr.Text = Val(Format(txtTotCr.Text, "#############.#0")) - Val(CounDifx)
            ElseIf Trim(Format(frmPettyCash.txtAmt.Text, "#############.#0")) > Trim(Format(PrvTotlLVWPC, "#############.#0")) Then
                CounDifx = Val(Format(frmPettyCash.txtAmt.Text, "#############.#0")) - Val(Format(PrvTotlLVWPC, "#############.#0"))
                txtTotCr.Text = Val(Format(txtTotCr.Text, "#############.#0")) + Val(CounDifx)
            End If
         Set MItem = Me.LVWPettyCash.ListItems.Add(, , Trim(txtPaidto.Text))
         MItem.SubItems(1) = Trim(txtPartic.Text)
         MItem.SubItems(2) = Trim(txtExpl.Text)
         MItem.SubItems(3) = Trim(txtInvNo.Text)
         MItem.SubItems(4) = Trim(txtPoNo.Text)
         MItem.SubItems(5) = Trim(mskPOdate.Text)
    
         MItem.SubItems(6) = Trim(cmbAccNo2.Text)
         MItem.SubItems(7) = Trim(cmbAccName2.Text)
         MItem.SubItems(9) = Format(Trim(txtAmt.Text), "############.#0")
    
    
         txtPaidto.Text = ""
         txtPartic.Text = ""
         txtExpl.Text = ""
         txtAmt.Text = ""
         cmbAccName2.Text = ""
         txtInvNo.Text = ""
         cmbAccNo2.Text = ""
    
    Me.cmdadd.caption = "&Add"
    End If
End If
End Sub

Private Sub cmdclose_Click()
If cmdclose.caption = "&Close" Then
Unload Me
ElseIf cmdclose.caption = "&Cancel" Then
cmdNew.caption = "&New"
txtserialNo.Text = ""
 

cmdclose.caption = "&Close"
End If
End Sub

Private Sub CMDEDIT_Click()
If CmdEdit.caption = "&Edit" Then
    Dim rsEdit As New ADODB.Recordset
    rsEdit.Open "Select distinct SerialNo from pettycashHeld", constring, adOpenDynamic, adLockOptimistic
    
    txtPaidto.Enabled = True
    txtPartic.Enabled = True
    txtExpl.Enabled = True
    txtInvNo.Enabled = True
    CMBAccName.Enabled = True
    CMBAccNo.Enabled = True
    cmbAccName2.Enabled = True
    cmbAccNo2.Enabled = True
    txtAmt.Enabled = True
    txtserialNo.Enabled = True
    MskDate.Enabled = True
     
  
    
    If rsEdit.EOF <> True Then
    rsEdit.MoveFirst
    End If
    While rsEdit.EOF = False
    txtserialNo.AddItem rsEdit!SerialNo
    rsEdit.MoveNext
    Wend
    
    CmdEdit.caption = "&Update"
    
    
    
ElseIf CmdEdit.caption = "&Update" Then

Dim DelLVWPC As New ADODB.Recordset
Dim SErx
SErx = Me.txtserialNo.Text
DelLVWPC.Open "delete  from Pettycashheld where serialno = " & "'" & SErx & "'" & "", constring, adOpenDynamic, adLockOptimistic

Call SavePetty
CmdEdit.caption = "&Edit"
End If
End Sub

Private Sub cmdNew_Click()
cmdNew.ToolTipText = "Save the Transaction"



txtPaidto.Enabled = True
txtPartic.Enabled = True
txtExpl.Enabled = True
txtInvNo.Enabled = True
CMBAccName.Enabled = True
CMBAccNo.Enabled = True
cmbAccName2.Enabled = True
cmbAccNo2.Enabled = True
txtAmt.Enabled = True
txtserialNo.Enabled = True
MskDate.Enabled = True


MskDate.Text = Format(Date, "dd/mm/yyyy")

If cmdNew.caption = "&New" Then
CMBAccName.SetFocus
End If




If cmdNew.caption = "&Save" Then


Call SavePetty
Exit Sub
End If



Dim rsForSerial As New ADODB.Recordset
rsForSerial.Open "Select Mynom from Pettycashheld order by mynom", constring, adOpenDynamic, adLockOptimistic
If rsForSerial.EOF = False Then
rsForSerial.MoveFirst
rsForSerial.MoveLast

Me.txtserialNo.Text = Val(rsForSerial!Mynom) + 1
End If
rsForSerial.close




cmdNew.caption = "&Save"
cmdclose.caption = "&Cancel"
CmdEdit.caption = "&Edit"
cmdadd.caption = "&Add"
Me.LVWPettyCash.ListItems.clear

    txtPaidto.Text = ""
    txtPartic.Text = ""
    txtExpl.Text = ""
    txtAmt.Text = ""
    cmbAccName2.Text = ""
    txtInvNo.Text = ""
    'cmbAccNo2.Text = ""
   ' txtSerialNo.Text = ""

End Sub
Private Sub SavePetty()



If CMBAccNo.Text = "" Then
MsgBox "Select the Combo 'Account No' for Petty Cash Held", vbInformation, "Combo is Empty"
Exit Sub
End If

If CMBAccName.Text = "" Then
MsgBox "Select the Combo 'Account Name' for Petty Cash Held", vbInformation, "Combo is Empty"
Exit Sub
End If

'If Val(Format(txtTotal.Text, "##########.0#")) <> Val(Format(txtTotCr.Text, "##########.0#")) Then
'MsgBox "Total Debit and Total Credit Balances are not Equal", vbInformation, "Check the Amount"
'Exit Sub
'End If

If cmdNew.caption = "&Save" Then
Dim rsForSerial As New ADODB.Recordset
rsForSerial.Open "Select Mynom from Pettycashheld order by mynom", constring, adOpenDynamic, adLockOptimistic
If rsForSerial.EOF = False Then
rsForSerial.MoveFirst
rsForSerial.MoveLast

Me.txtserialNo.Text = Val(rsForSerial!Mynom) + 1
End If
rsForSerial.close
End If





If Me.LVWPettyCash.ListItems.Count = 0 Then
MsgBox "ListView is Empty and You can not save it now", vbInformation
Exit Sub
End If

er = MsgBox("Do you want to save the entries", vbInformation + vbYesNo, "Please Select")
If er = vbNo Then
Exit Sub
End If

Dim RsPetty As New ADODB.Recordset
RsPetty.Open "Select* from PettyCashHeld", constring, adOpenDynamic, adLockOptimistic


Dim AccNo2, AccName2, Pitem, PaidTo, InvNo, Parti, Expl, Amt, APp, POn, POd, Amt2

n = 0
For n = 1 To Me.LVWPettyCash.ListItems.Count
    PaidTo = Me.LVWPettyCash.ListItems.Item(n)
    Parti = Me.LVWPettyCash.ListItems.Item(n).SubItems(1)
    Expl = Me.LVWPettyCash.ListItems.Item(n).SubItems(2)
    InvNo = Me.LVWPettyCash.ListItems.Item(n).SubItems(3)
    POn = Me.LVWPettyCash.ListItems.Item(n).SubItems(4)
    POd = Me.LVWPettyCash.ListItems.Item(n).SubItems(5)
    
    AccNo2 = Me.LVWPettyCash.ListItems.Item(n).SubItems(6)
    AccName2 = Me.LVWPettyCash.ListItems.Item(n).SubItems(7)
    Amt = Me.LVWPettyCash.ListItems.Item(n).SubItems(8)
    Amt2 = Me.LVWPettyCash.ListItems.Item(n).SubItems(9)
    

With RsPetty
.addnew
!SerialNo = txtserialNo.Text
If Text2.Text <> "" Then
!NoOfTimeEdited = Right(Text2, 1)
End If


'Every time when it repeats within the loop it is saved but the same value
!Datex = MskDate
!AccountNo = CMBAccNo.Text  'CredIT
!accountname = CMBAccName.Text  'Credit or

'-------------------------------------------------------------------------

!AccountNo2 = AccNo2


'Here is to get the Classifi


                            acctNo = Trim(AccNo2)
                            Prevcap = Trim(Me.caption)
                            Call DisplayCats(Prevcap, acctNo, catName)
                           ' Categry = catName


!Classification = catName


!accountname2 = AccName2
'TotalPtyCash
'Item
!PaidTo = PaidTo
!InvNo = InvNo
!PoNumber = POn
!PODate = POd
!Partic = Parti
!Explan = Expl
On Error Resume Next
!amount = Amt
!creditamount = Amt2
On Error GoTo 0
!motham = Val(Format(txtTotal.Text)) - Val(Format(txtTotCr.Text)) 'This is the Total Creidt Balance                                                         'Trim(txtTotal.Text)
!Prepby = cLogUser
!EntryType = cmbEnryType.Text
'Appr
'TotalDb
.Update
gotdar = 1
End With
Next
If gotdar = 1 Then
MsgBox "Records Saved Succusfully", vbInformation, Confirmation
End If
gptdar = ""
cmdNew.caption = "&New"
cmdclose.caption = "&Close"

'This is to call the Class for the Listview UN confirm
Dim xcls222 As New HabitatClass
Dim rs222 As New ADODB.Recordset
'xcls222.ProcLvwPCashUnCon rs222


End Sub

Private Sub cmdSearch3_Click()
 Dim itmFound As ListItem   ' FoundItem variable.
 strFindMe = txtSearch2.Text
 Dim intSelectedOption As String
 If WhatColumnNo = 0 Then
   intSelectedOption = lvwText
  Else
  intSelectedOption = lvwSubItem
 End If
    Set itmFound = LvwPCashUnConf.Finditem(strFindMe, intSelectedOption, , lvwPartial)
    If itmFound Is Nothing Then  ' If no match, inform user and exit.
       MsgBox "No match found"
    Exit Sub
    Else
    itmFound.EnsureVisible ' Scroll ListView to show found ListItem.
    itmFound.Selected = True   ' Select the ListItem.
    LvwPCashUnConf.SetFocus
    End If
    LvwPCashUnConf.MultiSelect = True




End Sub

Private Sub Form_Load()
WhatColumnNo = 0 'nelson
txtPaidto.Enabled = False
txtPartic.Enabled = False
txtExpl.Enabled = False
txtInvNo.Enabled = False
CMBAccName.Enabled = False
CMBAccNo.Enabled = False
cmbAccName2.Enabled = False
cmbAccNo2.Enabled = False
txtAmt.Enabled = False
txtserialNo.Enabled = False
MskDate.Enabled = False

cmdNew.ToolTipText = "Click to make New Transaction"
'Set Listx = Me.LVWPettyCash.ColumnHeaders.Add(, , "Ticket")
Set Listx = Me.LVWPettyCash.ColumnHeaders.Add(, , "Paid To")
Set Listx = Me.LVWPettyCash.ColumnHeaders.Add(, , "Particulars", 2000)
Set Listx = Me.LVWPettyCash.ColumnHeaders.Add(, , "Explanations", 1600)
Set Listx = Me.LVWPettyCash.ColumnHeaders.Add(, , "Invoice No")
Set Listx = Me.LVWPettyCash.ColumnHeaders.Add(, , "S.E No")
Set Listx = Me.LVWPettyCash.ColumnHeaders.Add(, , "S.E Date")
Set Listx = Me.LVWPettyCash.ColumnHeaders.Add(, , "Account No", 1600)
Set Listx = Me.LVWPettyCash.ColumnHeaders.Add(, , "Account Name", 2500)
Set Listx = Me.LVWPettyCash.ColumnHeaders.Add(, , "Amount", , lvwColumnRight)
Set Listx = Me.LVWPettyCash.ColumnHeaders.Add(, , "Other Cr Amt", , lvwColumnRight)


Set ListUnConfirm = Me.LvwPCashUnConf.ColumnHeaders.Add(, , "No")
Set ListUnConfirm = Me.LvwPCashUnConf.ColumnHeaders.Add(, , "Date Entered")
Set ListUnConfirm = Me.LvwPCashUnConf.ColumnHeaders.Add(, , "Petty Cash Holder", 5720)
Set ListUnConfirm = Me.LvwPCashUnConf.ColumnHeaders.Add(, , "Printed", 1600)
Set ListUnConfirm = Me.LvwPCashUnConf.ColumnHeaders.Add(, , "Credit Amount", , lvwColumnRight)


Set ListConfirm = Me.LvwPCashConf.ColumnHeaders.Add(, , "No")
Set ListConfirm = Me.LvwPCashConf.ColumnHeaders.Add(, , "Date Entered")
Set ListConfirm = Me.LvwPCashConf.ColumnHeaders.Add(, , "Petty Cash Holder", 5720)
Set ListConfirm = Me.LvwPCashConf.ColumnHeaders.Add(, , "Printed", 1600)
Set ListConfirm = Me.LvwPCashConf.ColumnHeaders.Add(, , "Credit Amount", , lvwColumnRight)


Set ListJour = Me.LvwPCJournal.ColumnHeaders.Add(, , "Journal No")
Set ListJour = Me.LvwPCJournal.ColumnHeaders.Add(, , "No.")
Set ListJour = Me.LvwPCJournal.ColumnHeaders.Add(, , "Date")
Set ListJour = Me.LvwPCJournal.ColumnHeaders.Add(, , "Particulars", 2000)
Set ListJour = Me.LvwPCJournal.ColumnHeaders.Add(, , "Explanations", 1600)
Set ListJour = Me.LvwPCJournal.ColumnHeaders.Add(, , "Account No", 1600)
Set ListJour = Me.LvwPCJournal.ColumnHeaders.Add(, , "Account Name", 2500)
Set ListJour = Me.LvwPCJournal.ColumnHeaders.Add(, , "Db Amount", , lvwColumnRight)
Set ListJour = Me.LvwPCJournal.ColumnHeaders.Add(, , "CR Amount", , lvwColumnRight)

'Dim RsListUc As New ADODB.Recordset
'RsListUc.Open "Select Distinct SErialno,Accountno,Accountname,motham,DateX,PrintMark from pettycashheld where confMark is null", conString, adOpenDynamic, adLockOptimistic
'
'
'frmPettyCash.LvwPCashUnConf.ListItems.clear
'
'While RsListUc.EOF = False
'Set mitem = frmPettyCash.LvwPCashUnConf.ListItems.Add(, , Trim(RsListUc!serialno))
'mitem.SubItems(1) = Format(RsListUc!DateX, "dd/mm/yyyy")
''mitem.SubItems(2) = Trim(RsListUc!Accountno)
'mitem.SubItems(2) = Trim(RsListUc!AccountName)
'mitem.SubItems(3) = IIf(IsNull(RsListUc!printmark), "", (RsListUc!printmark))
'mitem.SubItems(4) = IIf(IsNull(RsListUc!motham), "", (RsListUc!motham))
'txttotUc.Text = Val(txttotUc) + Val(RsListUc!motham)
'RsListUc.MoveNext
'Wend

cmbEnryType.AddItem "Debit Entry"
cmbEnryType.AddItem "Other Credit Entry"


'Dim rstChartx As New ADODB.Recordset
'rstChartx.Open "Select * from financemaster order by Accountcode", conString, adOpenDynamic, adLockOptimistic
 Dim xClass As New HabitatClass
 Dim xtable As String
 Dim xtable2 As String

 Dim sqltable As Boolean
 Set rstChartx = New ADODB.Recordset

 sqltable = True
 Set conn1 = New ADODB.Connection
 xtable = "Select * from FinanceMaster order by AccountNameEng"

 xClass.GetTables rstChartx, conn1, xtable, constring, sqltable

 While rstChartx.EOF = False
      If rstChartx!Active = 1 Then
        If Mid(Trim(rstChartx!AccountCode), 1, 7) = "1111304" Then
        'Me.CMBAccNo.AddItem rstChartx!AccountCode
        Me.CMBAccName.AddItem rstChartx!accountnameeng & "\" & rstChartx!accountnamearab

        End If
        
        
        
        
        If Mid(Trim(rstChartx!AccountCode), 1, 7) = "1111304" Or Mid(Trim(rstChartx!AccountCode), 1, 6) = "111160" Or Mid(Trim(rstChartx!AccountCode), 1, 5) = "13105" Or Mid(Trim(rstChartx!AccountCode), 1, 7) = "1310402" Or Mid(Trim(rstChartx!AccountCode), 1, 7) = "1111304" Or Mid(Trim(rstChartx!AccountCode), 1, 5) = "13107" Or Mid(Trim(rstChartx!AccountCode), 1, 1) = "2" And Mid(Trim(rstChartx!AccountCode), 1, 3) <> "242" Then
        'Me.cmbAccNo2.AddItem rstChartx!AccountCode
        Me.cmbAccName2.AddItem rstChartx!accountnameeng & "\" & rstChartx!accountnamearab

        End If
         
       End If
      
    rstChartx.MoveNext
 Wend
rstChartx.close
 
 
 
 
  xtable2 = "Select * from FinanceMaster order by Accountcode"
 Set rstCharty = New ADODB.Recordset
 Set Constg = New ADODB.Connection
 xClass.GetTables rstCharty, Constg, xtable2, constring, sqltable

 While rstCharty.EOF = False
      If rstCharty!Active = 1 Then
        If Mid(Trim(rstCharty!AccountCode), 1, 7) = "1111304" Then
        Me.CMBAccNo.AddItem rstCharty!AccountCode
        'Me.CMBAccName.AddItem rstChartx!accountnameeng & "\" & rstChartx!accountnamearab

        End If
        
        If Mid(Trim(rstCharty!AccountCode), 1, 7) = "1111304" Or Mid(Trim(rstCharty!AccountCode), 1, 6) = "111160" Or Mid(Trim(rstCharty!AccountCode), 1, 5) = "13105" Or Mid(Trim(rstCharty!AccountCode), 1, 7) = "1310402" Or Mid(Trim(rstCharty!AccountCode), 1, 7) = "1111304" Or Mid(Trim(rstCharty!AccountCode), 1, 5) = "13107" Or Mid(Trim(rstCharty!AccountCode), 1, 1) = "2" And Mid(Trim(rstCharty!AccountCode), 1, 3) <> "242" Then
        Me.cmbAccNo2.AddItem rstCharty!AccountCode
       ' Me.cmbAccName2.AddItem rstChartx!accountnameeng & "\" & rstCharty!accountnamearab

        End If
         
       End If
      
    rstCharty.MoveNext
 Wend
rstCharty.close
 
 
 
 
 
 
 'This is to call the Class for the Listview confirmed
Dim xcls As New HabitatClass
Dim rs As New ADODB.Recordset
'xcls.ProcLvwPCashConfirmed rs

'This is to call the Class for the Listview UN confirm
Dim rs2 As New ADODB.Recordset
'xcls.ProcLvwPCashUnCon rs2

'This is to call the Class for the Listview UN confirm
Dim rs22 As New ADODB.Recordset
'xcls.ProcLvwPCJournal rs22

End Sub

Private Sub LvwPCashConf_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  LvwPCashConf.SortKey = ColumnHeader.Index - 1
 ' Set Sorted to True to sort the list.
   LvwPCashConf.Sorted = True

'Combo1.Text = LvwPCashConf.ColumnHeaders(ColumnHeader.Index)

End Sub

Private Sub LvwPCashConf_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = 2 Then
  PopupMenu frmMenu.PettCashCon
End If

End Sub

Private Sub LvwPCashUnConf_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  LvwPCashUnConf.SortKey = ColumnHeader.Index - 1
 ' Set Sorted to True to sort the list.
   LvwPCashUnConf.Sorted = True


'Combo1.Text = LvwPCashConf.ColumnHeaders(ColumnHeader.Index)

WhatColumnNo = ColumnHeader.Index - 1
End Sub

Private Sub LvwPCashUnConf_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = 2 Then
  PopupMenu frmMenu.PettCashUnCon
End If

End Sub

Private Sub LvwPCJournal_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  LvwPCJournal.SortKey = ColumnHeader.Index - 1
 ' Set Sorted to True to sort the list.
   LvwPCJournal.Sorted = True

'Combo1.Text = LvwPCJournal.ColumnHeaders(ColumnHeader.Index)

End Sub

Private Sub LvwPCJournal_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = 2 Then
 PopupMenu frmMenu.PettyJournal
End If
End Sub

Private Sub LVWPettyCash_Click()
  If Me.CmdEdit.caption = "&Update" And Me.LVWPettyCash.ListItems.Count <> 0 Then
        'frmPettyCash.txtFlag.Text = frmPettyCash.LVWPettyCash.SelectedItem
        frmPettyCash.txtPaidto.Text = frmPettyCash.LVWPettyCash.SelectedItem
        frmPettyCash.txtPartic.Text = frmPettyCash.LVWPettyCash.SelectedItem.SubItems(1)
        frmPettyCash.txtExpl.Text = frmPettyCash.LVWPettyCash.SelectedItem.SubItems(2)
        frmPettyCash.txtInvNo.Text = frmPettyCash.LVWPettyCash.SelectedItem.SubItems(3)
        frmPettyCash.txtPoNo.Text = frmPettyCash.LVWPettyCash.SelectedItem.SubItems(4)
        On Error Resume Next
        frmPettyCash.mskPOdate.Text = frmPettyCash.LVWPettyCash.SelectedItem.SubItems(5)
        On Error GoTo 0
        frmPettyCash.cmbAccNo2.Text = frmPettyCash.LVWPettyCash.SelectedItem.SubItems(6)
        frmPettyCash.cmbAccName2.Text = frmPettyCash.LVWPettyCash.SelectedItem.SubItems(7)
        frmPettyCash.txtAmt.Text = frmPettyCash.LVWPettyCash.SelectedItem.SubItems(8)
        
        If frmPettyCash.LVWPettyCash.SelectedItem.SubItems(8) = "" Then
        cmbEnryType.Text = "Other Credit Entry"
        Else
        cmbEnryType.Text = "Debit Entry"
        End If
        


        PrvTotlLVWPC = frmPettyCash.txtAmt.Text
        'PrvTotCRlLVWPC = frmPettyCash.txtAmt.Text
        
       ' frmMenu.PetCash.Caption = "Update"
        frmMenu.PetCEdit.caption = "Update"
        Me.cmdadd.caption = "Update"
 End If
End Sub

Private Sub LVWPettyCash_KeyDown(KeyCode As Integer, Shift As Integer)
LVWPettyCash_Click

End Sub

Private Sub LVWPettyCash_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = 2 Then
  PopupMenu frmMenu.PetCash
End If

End Sub

Private Sub mskPOdate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbAccName2.SetFocus
End If

End Sub

Private Sub txtAmt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtInvNo.SetFocus
End If

End Sub

Private Sub txtAmt_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode <> 13 Then
 If KeyCode <> 110 Then
 If KeyCode < 96 Or KeyCode > 105 Then

 If KeyCode < 48 Or KeyCode > 57 Then
    MsgBox "The Key you entered is not the Digit", vbExclamation, "Try Again"
    Me.txtAmt = ""
    Exit Sub
End If
End If
End If
End If

End Sub

Private Sub txtAmt_LostFocus()

txtAmt.Text = Format(txtAmt.Text, "###,###,###,###.#0")

End Sub

Private Sub txtExpl_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbAccNo2.SetFocus
End If

End Sub

Private Sub txtinvNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtPoNo.SetFocus
End If

End Sub

Private Sub txtPaidto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtPartic.SetFocus
End If

End Sub

Private Sub txtPartic_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtExpl.SetFocus
End If

End Sub

Private Sub txtPoNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
mskPOdate.SetFocus
End If
End Sub

Private Sub txtSearch2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmdSearch3_Click
End If
End Sub

Private Sub txtSerialNo_Click()





Dim VarUI, VarrsTake2LIst, totCRforLVWPC
Dim rsTake2LIst As New ADODB.Recordset

VarUI = Me.txtserialNo.Text



VarrsTake2LIst = "Select * from pettycashHeld where SErialno = " & "'" & VarUI & "'" & ""

rsTake2LIst.Open VarrsTake2LIst, constring, adOpenDynamic, adLockOptimistic


If rsTake2LIst.EOF = False Then
 rsTake2LIst.MoveFirst
End If

Me.CMBAccName.Text = rsTake2LIst!accountname
Me.CMBAccNo.Text = rsTake2LIst!AccountNo
Me.MskDate.Text = rsTake2LIst!Datex
Me.LVWPettyCash.ListItems.clear
  While rsTake2LIst.EOF = False
    ' Set mitem = Me.LVWPettyCash.ListItems.Add(, , Format(rsTake2LIst!serialno))
     
     Set MItem = Me.LVWPettyCash.ListItems.Add(, , Trim(rsTake2LIst!PaidTo))
    ' mitem.SubItems(1) = Trim(rsTake2LIst!PaidTo)
     MItem.SubItems(1) = Trim(rsTake2LIst!Partic)
     MItem.SubItems(2) = Trim(rsTake2LIst!Explan)
     MItem.SubItems(3) = IIf(IsNull(rsTake2LIst!InvNo), "", (rsTake2LIst!InvNo))
     MItem.SubItems(4) = IIf(IsNull(rsTake2LIst!PoNumber), "", (rsTake2LIst!PoNumber))
     MItem.SubItems(5) = IIf(IsNull(rsTake2LIst!PODate), "", (rsTake2LIst!PODate))
     
     
     MItem.SubItems(6) = IIf(IsNull(rsTake2LIst!AccountNo2), "", (rsTake2LIst!AccountNo2))
     MItem.SubItems(7) = Trim(rsTake2LIst!accountname2)
     MItem.SubItems(8) = Format(IIf(IsNull(rsTake2LIst!amount), "", (rsTake2LIst!amount)), "#############.#0")
     totCRforLVWPC = Val(totCRforLVWPC) + Val(rsTake2LIst!amount)
         
     rsTake2LIst.MoveNext
     Wend
 Me.txtTotal.Text = totCRforLVWPC
 
 'rsNoOfTimeEdited.Close
 Dim rsNoOfTimeEdited As New ADODB.Recordset
rsNoOfTimeEdited.Open "Select Distinct Serialno,NoOfTimeEdited from Pettycashheld where serialno = " & "'" & VarUI & "'" & "", constring, adOpenDynamic, adLockOptimistic

If rsNoOfTimeEdited!NoOfTimeEdited <> 0 Then
Text2.Visible = True
Text2.Text = rsNoOfTimeEdited!SerialNo & "-" & Val(rsNoOfTimeEdited!NoOfTimeEdited) + 1
End If

End Sub

Private Sub txtserialNo_KeyPress(KeyAscii As Integer)
If CmdEdit.caption = "&Edit" And KeyAscii = 13 Then
MskDate.SetFocus
ElseIf CmdEdit.caption = "&Update" Then

End If


End Sub

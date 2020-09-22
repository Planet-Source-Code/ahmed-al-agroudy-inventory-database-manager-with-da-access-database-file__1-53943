VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmassigningcheck 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Assigning Check  ÇáÔíßÇÊ ÇáãÓÌáÉ"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7335
   FillColor       =   &H00008000&
   Icon            =   "frmassigningcheck.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   7335
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3855
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   7335
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   2160
         Top             =   2160
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmassigningcheck.frx":0442
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmassigningcheck.frx":0894
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   350
         Left            =   3750
         TabIndex        =   14
         Top             =   540
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   609
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox txtnumber 
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   14
         Format          =   "##############"
         Mask            =   "##############"
         PromptChar      =   " "
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2415
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   4260
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Bank Name"
            Object.Width           =   8114
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Check Number"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.ComboBox combankname 
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
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   120
         Width           =   3855
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "&Ok "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdcancel 
         Caption         =   "&Close  "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   600
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   165
         Left            =   120
         TabIndex        =   12
         Top             =   3600
         Visible         =   0   'False
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   291
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.Label txttotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   3380
         TabIndex        =   15
         Top             =   540
         Width           =   375
      End
      Begin VB.Label lp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "100 %"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6720
         TabIndex        =   13
         Top             =   3600
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No of Check ÚÏÏ ÇáÔíßÇÊ "
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4080
         TabIndex        =   10
         Top             =   600
         Width           =   1905
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Check Number ÑÞã ÇáÔíß "
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1800
         TabIndex        =   9
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Assigned checks"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   1470
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   7
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Name ÇÓã ÇáÈäß "
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4320
         TabIndex        =   6
         Top             =   120
         Width           =   1650
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ÇÌãÇáí ÇáÔíßÇÊ ÇáÌÇÑí ÊÓÌíáåÇ"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   5280
         TabIndex        =   5
         Top             =   960
         Width           =   1920
      End
   End
End
Attribute VB_Name = "frmassigningcheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim recche As New ADODB.Recordset
Dim recbank As New ADODB.Recordset
Dim con2 As New ADODB.Connection

Private Sub cmdcancel_Click()
Unload Me
End Sub

Private Sub cmdok_Click()
recche.Requery
ListView1.ListItems.Clear

    If Val(txtnumber.Text) <= 0 Then
        MsgBox "Please Enter the Check number", vbInformation, "Missing Number"
        txtnumber.SetFocus
        Exit Sub
    End If
    
ProgressBar1.Min = 0
ProgressBar1.Max = Val(Trim(txttotal.caption)) + 1
'this is for checking for the check numbers
    Dim recchecknumber As New ADODB.Recordset
    i = Val(Trim(txtnumber.Text))
    X = 1 'this is for listview
    Do While i <= Val(Val(Trim(txtnumber.Text)) + Val(Trim(txttotal.caption)) - 1)
         recchecknumber.Open "select * from checkregister where cancel <> '1' and bank = '" & Trim(combankname.Text) & "' and checknumber = '" & i & "'", con2, adOpenKeyset, adLockOptimistic
         'MsgBox recchecknumber.RecordCount
            If recchecknumber.BOF = False Then
                ListView1.ListItems.Add , , "This is Already Registered.  " & recchecknumber!receiptno & "   " & Format(Trim(recchecknumber!valuedate), "dd/mm/yyyy"), , 1
                ListView1.ListItems(X).ListSubItems.Add , , i
                ListView1.ListItems.Item(X).ForeColor = &HC0&

            Else
                With recche
                    .AddNew
                    !bank = Trim(combankname.Text)
                    !checknumber = i
                    !singned = 1
                    .Update
                End With
                ProgressBar1.Visible = True
                lp.Visible = True
                On Error Resume Next
                DoEvents
                ProgressBar1.Value = ProgressBar1.Value + 1
                lp.caption = Format(((ProgressBar1.Value * 100) / ProgressBar1.Max), "###") & " %"
                ListView1.ListItems.Add , , "This is Registered Successfully.", , 2
                ListView1.ListItems(X).ListSubItems.Add , , i
                ListView1.ListItems.Item(X).ForeColor = &H8000&

            End If
            recchecknumber.Close
        i = i + 1
        X = X + 1
    Loop
'end checking the check numbers

    ProgressBar1.Value = ProgressBar1.Max
    ProgressBar1.Visible = False
    lp.Visible = False
    cmdcancel.SetFocus
    MsgBox "Your Check Number Registerd Successfully", vbInformation, "Registerd"
    Frame1.Enabled = True
    ListView1.Enabled = True
End Sub

Private Sub combankname_Click()
ListView1.ListItems.Clear
txtnumber.Text = ""
txttotal.caption = 1
End Sub

Private Sub combankname_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyPageDown
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(combankname.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Select
End Sub

Private Sub combankname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Trim(combankname.Text) <> "" Then
    txtnumber.SetFocus
End If
End Sub

Private Sub combankname_LostFocus()
If Trim(combankname.Text) = "" Then
    MsgBox "Please Choose the Bank Name ", vbInformation, "No Bank Name"
    combankname.SetFocus
    Exit Sub
End If

If Trim(combankname.Text) <> "" Then
    On Error Resume Next
    recche.Close
    On Error GoTo 0
    recche.Open "Select * from checkregister where bank=" & "'" & Trim(combankname.Text) & "'" & " order by autonumber", con2, adOpenKeyset, adLockOptimistic
    txttotal.caption = "1"
    txtnumber.SetFocus
End If
End Sub

Private Sub Form_Load()
con2.Open "dsN=fINANCE;UID=SA;PWD=;"
'this is for open the bank table
recbank.Open "Select * from  financemaster where substring(accountcode,1,7) = '1110201'", con2, adOpenKeysetm, adLockOptimistic

If recbank.BOF = False Then
    recbank.MoveFirst
    While recbank.EOF = False
        combankname.AddItem Trim(recbank!accountnameeng)
        recbank.MoveNext
    Wend
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
con2.Close
End Sub

Private Sub txtnumber_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Val(Trim(txtnumber.Text)) > 0 Then
    UpDown1.SetFocus
End If
End Sub

Private Sub txtnumber2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Val(Trim(txtnumber.Text)) > 0 Then
    cmdOK.SetFocus
End If
End Sub


Private Sub UpDown1_DownClick()
If Val(txttotal.caption) > 1 Then
txttotal.caption = Val(txttotal.caption) - 1
End If
End Sub

Private Sub UpDown1_UpClick()
If Val(txttotal.caption) < 50 Then
txttotal.caption = Val(txttotal.caption) + 1
End If
End Sub

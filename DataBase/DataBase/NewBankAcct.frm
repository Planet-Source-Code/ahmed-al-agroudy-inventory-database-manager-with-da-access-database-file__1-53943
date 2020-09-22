VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form NewBankAcct 
   Caption         =   "New Bank Account ÍÓÇÈÇÊ ÇáÈäß ÇáÌÏíÏÉ "
   ClientHeight    =   6660
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   5805
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   6255
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   11033
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Data Entry ÇÏÎÇá ÇáÈíÇäÇÊ "
      TabPicture(0)   =   "NewBankAcct.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label7"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label8"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label9"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label10"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label11"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label12"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label13"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Combo1"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Combo2"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Combo3"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Combo4"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Combo5"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Combo6"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Combo7"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Combo8"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Combo9"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Combo10"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Combo11"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Combo12"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Combo13"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Command2"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).ControlCount=   27
      TabCaption(1)   =   "&List ÞÇÆãÉ "
      TabPicture(1)   =   "NewBankAcct.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "ListView1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin MSComctlLib.ListView ListView1 
         Height          =   5775
         Left            =   60
         TabIndex        =   28
         Top             =   360
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   10186
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   13
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "a"
            Text            =   "AccountNumber"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "b"
            Text            =   "AccountName"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Key             =   "c"
            Text            =   "BankCode"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Key             =   "d"
            Text            =   "BankNameEng"
            Object.Width           =   6245
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "BankNameArab"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Key             =   "e"
            Text            =   "BankAddress"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "ChipNo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "SwiftNo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Key             =   "f"
            Text            =   "Telephone#"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Fax No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Key             =   "g"
            Text            =   "Mobile"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Key             =   "h"
            Text            =   "Telex"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Key             =   "i"
            Text            =   "Contact Person"
            Object.Width           =   3246
         EndProperty
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Ok"
         Enabled         =   0   'False
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
         Left            =   -72480
         TabIndex        =   27
         Top             =   5520
         Width           =   1095
      End
      Begin VB.ComboBox Combo13 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -73200
         Style           =   1  'Simple Combo
         TabIndex        =   26
         Top             =   4920
         Width           =   3855
      End
      Begin VB.ComboBox Combo12 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -73200
         Style           =   1  'Simple Combo
         TabIndex        =   24
         Top             =   4560
         Width           =   3855
      End
      Begin VB.ComboBox Combo11 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -73200
         Style           =   1  'Simple Combo
         TabIndex        =   22
         Top             =   4200
         Width           =   1935
      End
      Begin VB.ComboBox Combo10 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -73200
         Style           =   1  'Simple Combo
         TabIndex        =   20
         Top             =   3840
         Width           =   1935
      End
      Begin VB.ComboBox Combo9 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -73200
         Style           =   1  'Simple Combo
         TabIndex        =   18
         Top             =   3480
         Width           =   1935
      End
      Begin VB.ComboBox Combo8 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -73200
         Style           =   1  'Simple Combo
         TabIndex        =   16
         Top             =   3120
         Width           =   1935
      End
      Begin VB.ComboBox Combo7 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -73200
         Style           =   1  'Simple Combo
         TabIndex        =   14
         Top             =   2760
         Width           =   1935
      End
      Begin VB.ComboBox Combo6 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -73200
         Style           =   1  'Simple Combo
         TabIndex        =   12
         Top             =   2400
         Width           =   2535
      End
      Begin VB.ComboBox Combo5 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -73200
         Style           =   1  'Simple Combo
         TabIndex        =   10
         Top             =   2040
         Width           =   3855
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -73200
         Style           =   1  'Simple Combo
         TabIndex        =   8
         Top             =   1680
         Width           =   3135
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -73200
         Style           =   1  'Simple Combo
         TabIndex        =   6
         Top             =   1320
         Width           =   3135
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -73200
         Style           =   1  'Simple Combo
         TabIndex        =   4
         Top             =   960
         Width           =   3135
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -73200
         Style           =   1  'Simple Combo
         TabIndex        =   2
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label13 
         Caption         =   "Contact Person"
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
         Left            =   -74640
         TabIndex        =   25
         Top             =   4920
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "Telex"
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
         Left            =   -74640
         TabIndex        =   23
         Top             =   4560
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Mobile"
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
         Left            =   -74640
         TabIndex        =   21
         Top             =   4200
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "Fax Number"
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
         Left            =   -74640
         TabIndex        =   19
         Top             =   3840
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Telephone No."
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
         Left            =   -74640
         TabIndex        =   17
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Swift Number"
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
         Left            =   -74640
         TabIndex        =   15
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Chips Number"
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
         Left            =   -74640
         TabIndex        =   13
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Bank Address"
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
         Left            =   -74640
         TabIndex        =   11
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "BankName in Arab"
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
         Left            =   -74640
         TabIndex        =   9
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "BankName in Eng"
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
         Left            =   -74640
         TabIndex        =   7
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Bank Code"
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
         Left            =   -74640
         TabIndex        =   5
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Account Name"
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
         Left            =   -74640
         TabIndex        =   3
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Account Number"
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
         Left            =   -74640
         TabIndex        =   1
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.Label Label14 
      Caption         =   "Click right mouse button to select options ÇÖÛØ  Úáí ÒÑ ÇáÝÇÑÉ áÇÎÊíÇÑ ÇáæÙÇÆÝ"
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
      Left            =   120
      TabIndex        =   29
      Top             =   75
      Width           =   5295
   End
   Begin VB.Menu main 
      Caption         =   "Main"
      Visible         =   0   'False
      Begin VB.Menu xAddnew 
         Caption         =   "Add New"
      End
      Begin VB.Menu xEdit 
         Caption         =   "Edit"
      End
      Begin VB.Menu Delete 
         Caption         =   "Delete"
      End
      Begin VB.Menu xREfresh 
         Caption         =   "Refresh"
      End
      Begin VB.Menu bar1 
         Caption         =   "-"
      End
      Begin VB.Menu xFind 
         Caption         =   "Find"
      End
   End
End
Attribute VB_Name = "NewBankAcct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rst As New ADODB.Recordset
Dim xAcctNum
Dim xAN
Dim xBC
Dim xBNe
Dim xBNa
Dim xBAd
Dim xCN
Dim xSN
Dim xTN
Dim xFN
Dim xMN
Dim xXN
Dim xCP
Dim c1 'As Long
Dim c2 'As Long
Dim c3 'As Long
Dim c4 'As Long
Dim c5 'As Long
Dim c6 'As Long
Dim c7 'As Long
Dim c8 'As Long
Dim c9 'As Long
Dim c10 ' As Long
Dim c11 ' As Long
Dim c12 ' As Long
Dim c13 ' As Long
Dim MItem As ListItem
Sub DisplayBankAccts()
Me.ListView1.ListItems.clear
rst.MoveFirst
Do Until rst.EOF
    On Error Resume Next
    Set MItem = Me.ListView1.ListItems.Add(, , rst!accountnumber)
    MItem.SubItems(1) = rst!accountname
    MItem.SubItems(2) = rst!bankcode
    MItem.SubItems(3) = rst!banknameeng
    MItem.SubItems(4) = rst!BankNameArab
    MItem.SubItems(5) = rst!Bankaddress
    MItem.SubItems(6) = rst!ChipsNo
    MItem.SubItems(7) = rst!SwiftNo
    MItem.SubItems(8) = rst!TelNo
    MItem.SubItems(9) = rst!Faxno
    MItem.SubItems(10) = rst!Mobile
    MItem.SubItems(11) = rst!Telex
    MItem.SubItems(12) = rst!ContactPerson
    rst.MoveNext
Loop

End Sub
Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.Combo2.SetFocus
End If
End Sub

Private Sub Combo1_LostFocus()
If Me.Combo1 <> "" Then
    Me.Command2.Enabled = True
   Else
   Me.Command2.Enabled = False
End If
End Sub

Private Sub Combo10_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.Combo11.SetFocus
End If

End Sub

Private Sub Combo11_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.Combo12.SetFocus
End If

End Sub

Private Sub Combo12_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.Combo13.SetFocus
End If
End Sub

Private Sub Combo13_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.Command2.SetFocus
End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.Combo3.SetFocus
End If

End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.Combo4.SetFocus
End If

End Sub

Private Sub Combo4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.Combo5.SetFocus
End If

End Sub

Private Sub Combo5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.Combo6.SetFocus
End If

End Sub

Private Sub Combo6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.Combo7.SetFocus
End If

End Sub

Private Sub Combo7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.Combo8.SetFocus
End If

End Sub

Private Sub Combo8_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.Combo9.SetFocus
End If

End Sub

Private Sub Combo9_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.Combo10.SetFocus
End If

End Sub

Private Sub Command2_Click()
If SSTab1.TabCaption(0) <> "Edit BankAcct Info" Then
     mess = MsgBox("Do you want to save entries now?", vbOKCancel + vbQuestion, "Please confirm")
     If mess = vbOK Then
        With rst
             .addnew
             !accountnumber = Trim(Combo1)
             !accountname = Trim(Combo2)
             !bankcode = Trim(Combo3)
             !banknameeng = Trim(Combo4)
             !BankNameArab = Trim(Combo5)
             !Bankaddress = Trim(Combo6)
             !ChipsNo = Trim(Combo7)
             !SwiftNo = Trim(Combo8)
             !TelNo = Trim(Combo9)
             !Faxno = Trim(Combo10)
             !Mobile = Trim(Combo11)
             !Telex = Trim(Combo12)
             !ContactPerson = Trim(Combo13)
             .Update
         End With
       For Each Control In Me
        If TypeOf Control Is ComboBox Then
           Control.Text = ""
        End If
       Next
      End If
      
  Else
     
     mess = MsgBox("Do you want to save changes?", vbOKCancel + vbQuestion, "Please confirm")
     If mess = vbOK Then
        Dim rstEdit As New ADODB.Recordset
        xBa = Trim(Combo1)
        rstEdit.Open "Select * from BankAccount where AccountNumber=" & "'" & xBa & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
        With rstEdit
             !accountnumber = Trim(Combo1)
             !accountname = Trim(Combo2)
             !bankcode = Trim(Combo3)
             !banknameeng = Trim(Combo4)
             !BankNameArab = Trim(Combo5)
             !Bankaddress = Trim(Combo6)
             !ChipsNo = Trim(Combo7)
             !SwiftNo = Trim(Combo8)
             !TelNo = Trim(Combo9)
             !Faxno = Trim(Combo10)
             !Mobile = Trim(Combo11)
             !Telex = Trim(Combo12)
             !ContactPerson = Trim(Combo13)
             .Update
         End With
        For Each Control In Me
        If TypeOf Control Is ComboBox Then
           Control.Text = ""
        End If
       Next
      End If
      Me.Combo1.SetFocus
    End If
End Sub

Private Sub delete_Click()

cindex = Me.ListView1.SelectedItem.Index
Ac = Me.ListView1.SelectedItem.SubItems(1)
mess = MsgBox("Are you sure you want to delete " & Ac, vbQuestion + vbOKCancel + vbDefaultButton2, "Please confirm")
If mess = vbOK Then
    Dim RstBA As New ADODB.Recordset
    RstBA.Open "Delete BankAccount where AccountNumber =" & "'" & Ac & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
    Me.ListView1.ListItems.Remove cindex
End If
End Sub

Private Sub Form_Load()
c1 = Combo1.Width
c2 = Combo2.Width
c3 = Combo3.Width
c4 = Combo4.Width
c5 = Combo5.Width
c6 = Combo6.Width
c7 = Combo7.Width
c8 = Combo8.Width
c9 = Combo9.Width
c10 = Combo10.Width
c11 = Combo11.Width
c12 = Combo12.Width
c13 = Combo13.Width
rst.Open "BankAccount", constring, adOpenKeyset, adLockPessimistic, adCmdTable
DisplayBankAccts
End Sub

Private Sub Form_Resize()
On Error Resume Next
Me.SSTab1.Width = Me.Width - 150
Me.SSTab1.Height = Me.Height - 800
Me.Command2.Left = Me.SSTab1.Width / 2 - 400
Me.ListView1.Width = Me.Width - 290
Me.ListView1.Height = Me.Height - 1220

Combo1.Width = Me.SSTab1.Width - 6300 + c1
Combo2.Width = Me.SSTab1.Width - 6300 + c2
Combo3.Width = Me.SSTab1.Width - 6300 + c3
Combo4.Width = Me.SSTab1.Width - 6300 + c4
Combo5.Width = Me.SSTab1.Width - 6300 + c5
Combo6.Width = Me.SSTab1.Width - 6300 + c6
Combo7.Width = Me.SSTab1.Width - 6300 + c7
Combo8.Width = Me.SSTab1.Width - 6300 + c8
Combo9.Width = Me.SSTab1.Width - 6300 + c9
Combo10.Width = Me.SSTab1.Width - 6300 + c10
Combo11.Width = Me.SSTab1.Width - 6300 + c11
Combo12.Width = Me.SSTab1.Width - 6300 + c12
Combo13.Width = Me.SSTab1.Width - 6300 + c13
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Me.ListView1.SortKey = ColumnHeader.Index - 1
Me.ListView1.Sorted = True
End Sub

Private Sub ListView1_DblClick()
Call xedit_Click
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = 2 Then
    PopupMenu main
End If

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If Me.SSTab1.TabCaption(PreviousTab) = "Edit BankAcct Info" Or Me.SSTab1.TabCaption(PreviousTab) = "Add New BankAcct" Then
   Me.SSTab1.TabCaption(PreviousTab) = "Data Entry"
End If
End Sub

Private Sub SSTab1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = 2 Then
    PopupMenu main
End If
End Sub

Private Sub xAddnew_Click()
Me.SSTab1.SetFocus
SSTab1.TabCaption(0) = "Add New BankAcct"
For Each Control In Me
    If TypeOf Control Is ComboBox Then
       Control.Text = ""
    End If
Next
SendKeys "{Right}"

End Sub

Private Sub xedit_Click()
 Me.Combo1 = Me.ListView1.SelectedItem.Text
 Me.Combo2 = Me.ListView1.SelectedItem.SubItems(1)
 Me.Combo3 = Me.ListView1.SelectedItem.SubItems(2)
 Me.Combo4 = Me.ListView1.SelectedItem.SubItems(3)
 Me.Combo5 = Me.ListView1.SelectedItem.SubItems(4)
 Me.Combo6 = Me.ListView1.SelectedItem.SubItems(5)
 Me.Combo7 = Me.ListView1.SelectedItem.SubItems(6)
 Me.Combo8 = Me.ListView1.SelectedItem.SubItems(7)
 Me.Combo9 = Me.ListView1.SelectedItem.SubItems(8)
 Me.Combo10 = Me.ListView1.SelectedItem.SubItems(9)
 Me.Combo11 = Me.ListView1.SelectedItem.SubItems(10)
 Me.Combo12 = Me.ListView1.SelectedItem.SubItems(11)
 Me.Combo13 = Me.ListView1.SelectedItem.SubItems(12)
 xAcctNum = Combo1
 xAN = Combo2
 xBC = Combo3
 xBNe = Combo4
 xBNa = Combo5
 xBAd = Combo6
 xCN = Combo7
 xSN = Combo8
 xTN = Combo9
 xFN = Combo10
 xMN = Combo11
 xXN = Combo12
 xCP = Combo13
 Me.Command2.Enabled = True
 Me.SSTab1.SetFocus
 SSTab1.TabCaption(0) = "Edit BankAcct Info"
 SendKeys "{Right}"
End Sub

Private Sub xREfresh_Click()
DisplayBankAccts
End Sub

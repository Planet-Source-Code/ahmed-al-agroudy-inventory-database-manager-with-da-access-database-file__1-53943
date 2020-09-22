VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form NewASset 
   Caption         =   "Register New Asset  ÇáÇÔÊÑÇß ÍÓÇÈ ÌÏíÏáÇ "
   ClientHeight    =   4755
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8745
   LinkTopic       =   "Form1"
   ScaleHeight     =   4755
   ScaleWidth      =   8745
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Ok ãæÇÝÞ "
      Height          =   350
      Left            =   2160
      TabIndex        =   25
      Top             =   4320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox Combo10 
      Height          =   315
      Left            =   840
      Style           =   1  'Simple Combo
      TabIndex        =   24
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close ÛáÞ "
      Height          =   350
      Left            =   5640
      TabIndex        =   21
      Top             =   4320
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   7435
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabMaxWidth     =   3881
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Data Entry ÇÏÎÇá ÇáÈíÇäÇÊ "
      TabPicture(0)   =   "NewASset.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label15"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label7"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label17"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label21"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label19"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label18"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label8"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label9"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label10"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label11"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label12"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label13"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label14"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Combo9"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Command1"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Combo8"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Combo7"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Combo6"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Combo5"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Combo3"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Combo2"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Combo1"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Combo4"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Combo11"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).ControlCount=   30
      TabCaption(1)   =   "List  ÞÇÆãÉ"
      TabPicture(1)   =   "NewASset.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "ListView1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.ComboBox Combo11 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -72600
         Locked          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   27
         Top             =   3720
         Width           =   2775
      End
      Begin VB.ComboBox Combo4 
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
         Left            =   -72600
         Style           =   1  'Simple Combo
         TabIndex        =   18
         Top             =   3360
         Width           =   2055
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3795
         Left            =   70
         TabIndex        =   20
         Top             =   360
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   6694
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID No."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Category"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Sub-Category"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Name in English"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Name in Arabic"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Serial No."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Model No."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "GL Account Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Date Registered"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -72600
         Locked          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   2
         Top             =   480
         Width           =   1095
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -72600
         Style           =   1  'Simple Combo
         TabIndex        =   4
         Top             =   840
         Width           =   1695
      End
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
         Left            =   -72600
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1200
         Width           =   3255
      End
      Begin VB.ComboBox Combo5 
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
         Left            =   -72600
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1560
         Width           =   3255
      End
      Begin VB.ComboBox Combo6 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -72600
         Locked          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   10
         Top             =   1920
         Width           =   2055
      End
      Begin VB.ComboBox Combo7 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -72600
         Style           =   1  'Simple Combo
         TabIndex        =   12
         Top             =   2280
         Width           =   3255
      End
      Begin VB.ComboBox Combo8 
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
         Left            =   -72600
         RightToLeft     =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   14
         Top             =   2640
         Width           =   3255
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Ok ãæÇÝÞ "
         Height          =   350
         Left            =   -69360
         TabIndex        =   19
         Top             =   3720
         Width           =   1215
      End
      Begin VB.ComboBox Combo9 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -72600
         Style           =   1  'Simple Combo
         TabIndex        =   16
         Top             =   3000
         Width           =   2055
      End
      Begin VB.Label Label14 
         Caption         =   "ÇáÇÓã ÈÇáÚÑÈí "
         Height          =   375
         Left            =   -69120
         TabIndex        =   36
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label13 
         Caption         =   "ÇáÇÓã ÈÇáÇäÌáíÒí "
         Height          =   255
         Left            =   -69120
         TabIndex        =   35
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label12 
         Caption         =   "ÑÞã ÍÓÇÈÇÊ ÚÇãÉ "
         Height          =   375
         Left            =   -69120
         TabIndex        =   34
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "ÇÝÆÉ ËÇäæíÉ "
         Height          =   375
         Left            =   -69120
         TabIndex        =   33
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "ÝÆÉ "
         Height          =   375
         Left            =   -69120
         TabIndex        =   32
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "ßæÏ "
         Height          =   375
         Left            =   -69120
         TabIndex        =   31
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "ÑÞã I D"
         Height          =   375
         Left            =   -69120
         TabIndex        =   30
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label18 
         Caption         =   "áÑÞã ÇáÊÓáÓá "
         Height          =   255
         Left            =   -69120
         TabIndex        =   29
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label Label19 
         Caption         =   "ÑÞã ÇáãæÏíá "
         Height          =   255
         Left            =   -69120
         TabIndex        =   28
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Caption         =   "Accumulated Dep Code"
         Height          =   375
         Left            =   -74760
         TabIndex        =   26
         Top             =   3720
         Width           =   1695
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "Model No."
         Height          =   375
         Left            =   -74400
         TabIndex        =   17
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "ID No"
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
         Left            =   -74400
         TabIndex        =   1
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Code"
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
         Left            =   -74400
         TabIndex        =   3
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Main Category"
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
         Left            =   -74400
         TabIndex        =   5
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Sub Category"
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
         Left            =   -74400
         TabIndex        =   7
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "GL Acct No."
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
         Left            =   -74400
         TabIndex        =   9
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Name in English"
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
         Left            =   -74400
         TabIndex        =   11
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Name in Arab"
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
         Left            =   -74400
         TabIndex        =   13
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Serial No."
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
         Left            =   -74400
         TabIndex        =   15
         Top             =   3000
         Width           =   1335
      End
   End
   Begin VB.Label Label20 
      Caption         =   "&Find íæÌÏ "
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   4365
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label16 
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   4320
      Width           =   3135
   End
   Begin VB.Menu xmenu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu xAddnew 
         Caption         =   "Add new"
      End
      Begin VB.Menu xEdit 
         Caption         =   "Modify/Edit"
      End
      Begin VB.Menu xDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu bar1 
         Caption         =   "-"
      End
      Begin VB.Menu xprint 
         Caption         =   "Print..."
      End
      Begin VB.Menu xRefresh 
         Caption         =   "Refresh"
      End
      Begin VB.Menu xFind 
         Caption         =   "Find..."
      End
   End
End
Attribute VB_Name = "NewASset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Code As String
Dim NameEng As String
Dim NameArab As String
Dim SerialNo As String
Dim Modelno As String
Dim nextIdno As String

Dim ColumnName As String
Dim FindInColumn As Integer
Dim MItem As ListItem
Dim SubCat As String
Dim SubCat1 As String
Dim AssetglCode As String
Dim rssubCat1 As New ADODB.Recordset
Dim rsAssetCat As New ADODB.Recordset
Dim rsSubCat As New ADODB.Recordset
Dim rsAssetID As New ADODB.Recordset

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Me.Combo3.SetFocus
End If
End Sub

Private Sub Combo3_Click()
xSubCat = Right(Trim(Me.Combo3), 12)
SubCat = Left(xSubCat, 7)
rsSubCat.Open "select * from Level5 where Left(AccountCode,7)=" & "'" & SubCat & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
Me.Combo5.clear
Me.Combo6.clear
Do Until rsSubCat.EOF = True
  If Trim(rsSubCat!MainAcct) <> "Main Accts" Then
     Me.Combo5.AddItem rsSubCat!accountnameeng & " - " & Trim(rsSubCat!AccountCode)
  End If
  rsSubCat.MoveNext
Loop
rsSubCat.close

Dim rsAccuDep As New ADODB.Recordset
ctext = InStr(1, Me.Combo3, "-", vbTextCompare)
AssetType = Trim(Left(Me.Combo3, ctext - 1))
rsAccuDep.Open "select * from level5 where accountnameeng=" & "'" & AssetType & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
If rsAccuDep.EOF = True Then
    msg = MsgBox("Can't create this Asset, No " & AssetType & " category found in Accumulated Depreciation. Check the right spelling", vbInformation + vbOKOnly, "Message")
    Exit Sub
End If
AssetglCode = Left(rsAccuDep!AccountCode, 9)
rsAccuDep.close
rsAccuDep.Open "select Count(*)as cTotal from level6 where Left(Accountcode,9)=" & "'" & AssetglCode & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
If cTOtal >= 999 Then
    msg = MsgBox("Can't create this Asset, Total items belongs to " & AssetType & " is already reached to 999", vbInformation + vbOKOnly, "Message")
    Exit Sub
End If
cTOtal = rsAccuDep!cTOtal + 1
If Len(rsAccuDep!cTOtal) = 1 Then
    xzero = "00"
 ElseIf Len(rsAccuDep!cTOtal) = 2 Then
    xzero = "0"
 ElseIf Len(rsAccuDep!cTOtal) = 3 Then
    xzero = ""
End If
accudepCode = AssetglCode & xzero & LTrim(cTOtal)
Me.Combo11 = accudepCode
End Sub

Private Sub Combo3_GotFocus()
Const CB_SHOWDROPDOWN = &H14F
Dim Tmp
Tmp = SendMessage(Combo3.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)

End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Me.Combo5.SetFocus
End If

End Sub

Private Sub Combo4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Me.Command1.SetFocus
End If

End Sub

Private Sub Combo5_Click()

xSubCat1 = Right(Trim(Me.Combo5), 12)
SubCat1 = Left(xSubCat1, 9)
rssubCat1.Open "select count(accountCode)as xTotal from Level6 where Left(AccountCode,9)=" & "'" & SubCat1 & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
If rssubCat1!xTotal + 1 > 999 Then
    msg = MsgBox("Unable to add items because you have reached the maximum 999 records under this Sub-Category", vbExclamation + vbOKOnly, "Message")
    Exit Sub
End If
If rssubCat1!xTotal + 1 > 9 Then
  codex = "0" & LTrim(rssubCat1!xTotal + 1)
ElseIf rssubCat1!xTotal + 1 > 99 Then
  codex = "" & LTrim(rssubCat1!xTotal + 1)
Else
codex = "00" & LTrim(rssubCat1!xTotal + 1)
End If
Me.Combo6 = SubCat1 & LTrim(codex)
rssubCat1.close
End Sub

Private Sub Combo5_GotFocus()
Const CB_SHOWDROPDOWN = &H14F
Dim Tmp
Tmp = SendMessage(Combo5.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)

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

Private Sub Combo7_GotFocus()
If Trim(Me.Combo5) = "" Then
  mes = MsgBox("Please Chose Sub-Category", vbExclamation + vbOKOnly, "Message")
  Me.Combo5.SetFocus
  Exit Sub
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
 Me.Combo4.SetFocus
End If

End Sub

Private Sub Command1_Click()
If Trim(Me.Combo2) = "" Or Trim(Me.Combo4) = "" Or Trim(Me.Combo3) = "" Or Trim(Me.Combo7) = "" Or Trim(Me.Combo8) = "" Or Trim(Me.Combo9) = "" Then
   msg = MsgBox("Please complete all the fields to filled up", vbExclamation + vbOKOnly, "Message")
   Exit Sub
 End If
Dim RsRegisterAsset As New ADODB.Recordset
Dim rsFm As New ADODB.Recordset
Dim rsLevel6 As New ADODB.Recordset
rsLevel6.Open "Level6", constring, adOpenKeyset, adLockPessimistic, adCmdTable
rsFm.Open "financeMaster", constring, adOpenKeyset, adLockPessimistic, adCmdTable
RsRegisterAsset.Open "Select * from AssetREgistered where Idno=" & "'" & Trim(Me.Combo1) & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
If RsRegisterAsset.EOF = True Then
   mess = MsgBox("Do you want save your entries?", vbQuestion + vbOKCancel, "Please confirm")
  Else
    If Trim(Me.Combo2) <> Code Or Trim(Me.Combo4) <> Modelno Or Trim(Me.Combo7) <> NameEng Or Trim(Me.Combo8) <> NameArab Or Trim(Me.Combo9) <> SerialNo Then
       mess = MsgBox("Do you want save changes?", vbQuestion + vbOKCancel, "Please confirm")
    End If
End If

If mess = vbOK Then
  If RsRegisterAsset.EOF = True Then
    'update Next Idno in Setup table
    With rsAssetID
        !nextAssetIdno = !nextAssetIdno + 1
        nextIdno = !nextAssetIdno
        .Update
    End With
    
    
    'add to AssetResgistered table
    With RsRegisterAsset
        .addnew
        !Idno = Trim(Me.Combo1)
        !Code = Trim(Me.Combo2)
        !category = Trim(Left(Trim(Me.Combo3), Len(Trim(Me.Combo3)) - 14))
        !Sub_category = Trim(Left(Trim(Me.Combo5), Len(Trim(Me.Combo5)) - 14))
        !GlAcctCode = Trim(Me.Combo6)
        !NameEng = Trim(Me.Combo7)
        !NameArab = Trim(Me.Combo8)
        !SerialNo = Trim(Me.Combo9)
        !Modelno = Trim(Me.Combo4)
        !DateRegistered = Date
        !AccumulatedCode = Trim(Me.Combo11)
        !AccumulatedName = Trim(Me.Combo7)
        .Update
        
     End With
   
   Else
     
    'update AssetResgistered table
    With RsRegisterAsset
        !Idno = Trim(Me.Combo1)
        !Code = Trim(Me.Combo2)
        !category = Trim(Me.Combo3) ', Len(Trim(Me.Combo3)) - 14))
        !Sub_category = Trim(Me.Combo5) ', Len(Trim(Me.Combo5)) - 14))
        !GlAcctCode = Trim(Me.Combo6)
        !NameEng = Trim(Me.Combo7)
        !NameArab = Trim(Me.Combo8)
        !SerialNo = Trim(Me.Combo9)
        !Modelno = Trim(Me.Combo4)
        !DateRegistered = Date
        .Update
        Me.Combo2 = ""
        Me.Combo7 = ""
        Me.Combo8 = ""
        Me.Combo9 = ""
        Me.Combo4 = ""
        rsSubCat.close
        msg = MsgBox("Item modified!", vbExclamation + vbOKOnly, "Message")
        Me.SSTab1.SetFocus
        SendKeys "{Right}"
        Exit Sub
     End With
   End If
   
       'add this into financeMAster table also
       With rsFm
         If Left(SubCat, 1) = "1" Then
            !FinancialCat = "Balance Sheet"
               If Left(SubCat, 1) = "1" Then
                    SubCat = "Asset"
                ElseIf Mid(SubCat, 2, 1) = "2" Then
                    SubCat = "Contra-Asset"
               End If
             End If
            .addnew
            !Country = xCountryNAMe
            !branch = Val(xBranchCode) '& "-" & xBranchName
            !SubCat = SubCat
            !AccountCode = Trim(Me.Combo6)
            !ClientCode = xClientCode
            !accountnameeng = Trim(Me.Combo7)
            !accountnamearab = Trim(Me.Combo8)
            !LastTransType = ""
            !BeginBal = 0
            !Credit = 0
            !Debit = 0
            !EndingBal = 0
            !TotalCredit = 0
            !TotalDebit = 0
            .Update
                    
            
           'add the accumulated dEpr account
           .addnew
           If Left(AssetglCode, 1) = "1" Then
            !FinancialCat = "Balance Sheet"
               If Left(AssetglCode, 1) = "1" Then
                    SubCat = "Asset"
                ElseIf Mid(AssetglCode, 2, 1) = "2" Then
                    SubCat = "Contra-Asset"
               End If
            End If
            
            !Country = xCountryNAMe
            !branch = Val(xBranchCode) '& "-" & xBranchName
            !SubCat = SubCat
            !AccountCode = Trim(Me.Combo11)
            !ClientCode = xClientCode
            !accountnameeng = Trim(Me.Combo7)
            !accountnamearab = Trim(Me.Combo8)
            !LastTransType = ""
            !BeginBal = 0
            !Credit = 0
            !Debit = 0
            !EndingBal = 0
            !TotalCredit = 0
            !TotalDebit = 0
            .Update
         End With



     With rsLevel6
        .addnew
         !Country = Val(xCountry)
         !CountryName = Trim(xCountryNAMe)
         !TopLevelCode = "1" 'TopLevelCode
         !TopLevelName = "Balance Sheet" 'Trim(TopLevelName)
         !Level1Code = "1"  'Level1Code
         !Level1Name = "Assets"  'Trim(Level1Name)
         !Level2Code = "2" 'Level2Code
         !Level2Name = "Long Term Assets" 'Trim(Level2Name)
         !level3Code = "1" 'level3Code
         !level3Name = "Fixed Assets" 'Trim(level3Name)
         !level4Code = Val(Mid(Me.Combo6, 6, 2))
         !level4Name = Trim(Left(Trim(Me.Combo3), Len(Trim(Me.Combo3)) - 14))
         !level5Code = Val(Trim(Mid(Me.Combo6, 8, 2))) 'level5Code
         !level5Name = Trim(Left(Trim(Me.Combo5), Len(Trim(Me.Combo5)) - 14)) 'Trim(level5Name))
         !Code = Val(Right(Me.Combo6, 3))
         !AccountCode = Me.Combo6
         !accountnameeng = Trim(Me.Combo7)
         !accountnamearab = Trim(Me.Combo8)
         !remarks = ""
         !MainAcct = "Main Accts"
         !remarks = "Main Account"
         .Update
          NextCode = Val(Right(Me.Combo11, 3))
          LevelBelong = Left(Me.Combo11, 11)
          Me.Combo11 = LevelBelong & LTrim(Val(Right(Me.Combo11, 3)))
     
        'ad the accumulated depreciation side 12201
         .addnew
         !Country = Val(xCountry)
         !CountryName = Trim(xCountryNAMe)
         !TopLevelCode = "1" 'TopLevelCode
         !TopLevelName = "Balance Sheet" 'Trim(TopLevelName)
         !Level1Code = "2"  'Level1Code
         !Level1Name = "Contra-Assets"  'Trim(Level1Name)
         !Level2Code = "2" 'Level2Code
         !Level2Name = "Long Term Contra Asset" 'Trim(Level2Name)
         !level3Code = "1" 'level3Code
         !level3Name = "Fixed Assets" 'Trim(level3Name)
         !level4Code = "1"
         !level4Name = "Accumulated Depreciation"
         !level5Code = Val(Right(AssetglCode, 2))
         !level5Name = Trim(Left(Trim(Me.Combo3), Len(Trim(Me.Combo3)) - 14)) 'Trim(level5Name))
         !Code = Val(Right(Me.Combo11, 3))
         !AccountCode = Me.Combo11
         !accountnameeng = Trim(Me.Combo7)
         !accountnamearab = Trim(Me.Combo8)
         !remarks = ""
         !MainAcct = "Main Accts"
         !remarks = "Main Account"
         .Update
       End With
         
       Me.Combo2 = ""
       Me.Combo7 = ""
       Me.Combo8 = ""
       Me.Combo9 = ""
       Me.Combo4 = ""
       rsLevel6.close
       
      
       'rsSubCat.Open "select * from Level5 where Left(AccountCode,7)=" & "'" & SubCat & "'", conString, adOpenKeyset, adLockPessimistic, adCmdText
       
       Call Combo5_Click
       Call Combo3_Click
       msg = MsgBox("Item added!", vbExclamation + vbOKOnly, "Message")
       Me.SSTab1.SetFocus
       SendKeys "{Right}"
End If

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
   Dim strFindMe As String
   Dim itmFound As ListItem   ' FoundItem variable.
   If FindInColumn > 1 Then
    intSelectedOption = lvwSubItem
    Else
    intSelectedOption = lvwText
   End If
   strFindMe = Trim(Me.Combo10)
   Set itmFound = Me.ListView1.Finditem(strFindMe, intSelectedOption, , lvwPartial)
   If itmFound Is Nothing Then  ' If no match, inform user and exit.
       xmsg = MsgBox("No item found in column  áÇ íæÌÏ ÇÕäÇÝ Ýí ÇáÚãæÏ " & ColumnName, vbExclamation + vbOKOnly, "Message")
       Exit Sub
    Else
        Me.Label20.Visible = False
        Me.Combo10.Visible = False
        Me.Command3.Visible = False
        itmFound.EnsureVisible
        itmFound.Selected = True   ' Select the ListItem.
        Me.ListView1.SetFocus
    End If
End Sub

Private Sub Form_Activate()
Me.Combo2.SetFocus
End Sub

Private Sub Form_Load()


'Dim rsAssetID As New ADODB.Recordset
rsAssetID.Open "Setup", constring, adOpenKeyset, adLockPessimistic, adCmdTable
nextIdno = rsAssetID!nextAssetIdno
CodeBelongASset = "11201" 'query in level4
CodeBelongAccumulatedDep = "12201" 'query it level5 table
rsAssetCat.Open "select * from Level4 where Left(AccountCode,5)=" & "'" & CodeBelongASset & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
Do Until rsAssetCat.EOF = True
   If Trim(rsAssetCat!MainAcct) <> "Main Accts" And UCase(Trim(rsAssetCat!accountnameeng)) <> UCase("Lands") Then
    Me.Combo3.AddItem Trim(rsAssetCat!accountnameeng) & " - " & Trim(rsAssetCat!AccountCode)
   End If
   rsAssetCat.MoveNext
Loop
Me.Combo1 = nextIdno
rsAssetCat.close

Call ViewList
End Sub
Sub ViewList()
Me.ListView1.ListItems.clear
Dim rsRegAsset As New ADODB.Recordset
rsRegAsset.Open "AssetRegistered", constring, adOpenKeyset, adLockPessimistic, adCmdTable
Do Until rsRegAsset.EOF = True
    Set MItem = Me.ListView1.ListItems.Add(, , rsRegAsset!Idno)
    MItem.SubItems(1) = rsRegAsset!Code
    MItem.SubItems(2) = rsRegAsset!category
    MItem.SubItems(3) = rsRegAsset!Sub_category
    MItem.SubItems(4) = rsRegAsset!NameEng
    MItem.SubItems(5) = rsRegAsset!NameArab
     MItem.SubItems(6) = rsRegAsset!SerialNo
      MItem.SubItems(7) = rsRegAsset!Modelno
    MItem.SubItems(8) = rsRegAsset!GlAcctCode
    MItem.SubItems(9) = rsRegAsset!DateRegistered
    rsRegAsset.MoveNext
Loop
Me.Label16.caption = Me.ListView1.ListItems.Count & " item(s) in the list"
End Sub
Private Sub Form_Resize()
Me.SSTab1.Width = Me.Width - 350
Me.ListView1.Width = Me.SSTab1.Width - 130
Me.Command1.Left = Me.SSTab1.Width - 1500
Me.Command2.Left = Me.SSTab1.Width - 1200
Me.SSTab1.Height = Me.Height - 870
Me.Command2.Top = Me.SSTab1.Height + 100
Me.Command1.Top = Me.SSTab1.Height - 500
Me.ListView1.Height = Me.SSTab1.Height - 420
Me.Label20.Top = Me.SSTab1.Height + 100
Me.Combo10.Top = Me.SSTab1.Height + 70
Me.Command3.Top = Me.SSTab1.Height + 70
End Sub

Private Sub Form_Unload(Cancel As Integer)
rsAssetID.close
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
FindInColumn = ColumnHeader.Index - 1
Me.ListView1.SortKey = ColumnHeader.Index - 1
ColumnName = Me.ListView1.ColumnHeaders(ColumnHeader.Index)
Me.ListView1.Sorted = True
End Sub

Private Sub ListView1_GotFocus()
 Me.Label20.Visible = False
Me.Combo10.Visible = False
Me.Command3.Visible = False
Me.Label16.Visible = True
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Me.ListView1.ListItems.Count <> 0 Then
    Me.xdelete.Enabled = True
    Me.xEdit.Enabled = True
    Me.xFind.Enabled = True
    Me.xPrint.Enabled = True
  Else
   Me.xdelete.Enabled = False
    Me.xEdit.Enabled = False
    Me.xFind.Enabled = False
    Me.xPrint.Enabled = False
 End If
 If Button = 2 Then
    PopupMenu Me.xmenu
End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If Left(Trim(Me.SSTab1.caption), 10) = "Data Entry" Then
 Me.Label16.Visible = False
Else
 Me.Label16.Visible = True
End If
End Sub

Private Sub xAddnew_Click()
On Error Resume Next
rsAssetID.close
rsAssetID.Open "Setup", constring, adOpenKeyset, adLockPessimistic, adCmdTable
nextIdno = rsAssetID!nextAssetIdno
CodeBelongASset = "11201"
rsAssetCat.Open "select * from Level4 where Left(AccountCode,5)=" & "'" & CodeBelongASset & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
Do Until rsAssetCat.EOF = True
   If Trim(rsAssetCat!MainAcct) <> "Main Accts" And UCase(Trim(rsAssetCat!accountnameeng)) <> UCase("Lands") Then
    Me.Combo3.AddItem Trim(rsAssetCat!accountnameeng) & " - " & Trim(rsAssetCat!AccountCode)
   End If
   rsAssetCat.MoveNext
Loop
Me.Combo1 = nextIdno
rsAssetCat.close
For Each Control In Me
    On Error Resume Next
     If TypeOf Control Is ComboBox Then
         Control.Text = ""
     End If
Next
Me.Combo1.Locked = False
Me.Combo3.Locked = False
Me.Combo5.Locked = False
Me.Combo6.Locked = False
Me.Combo1 = nextIdno
Me.SSTab1.SetFocus
SendKeys "{Left}"
End Sub

Private Sub xedit_Click()
Me.Combo5.clear
Me.Combo3.clear
Me.Combo1 = Trim(Me.ListView1.SelectedItem)
Me.Combo2 = Trim(Me.ListView1.SelectedItem.SubItems(1))
Me.Combo3 = Trim(Me.ListView1.SelectedItem.SubItems(2))
Me.Combo5 = Trim(Me.ListView1.SelectedItem.SubItems(3))
Me.Combo6 = Trim(Me.ListView1.SelectedItem.SubItems(8))
Me.Combo7 = Trim(Me.ListView1.SelectedItem.SubItems(4))
Me.Combo8 = Trim(Me.ListView1.SelectedItem.SubItems(5))
Me.Combo9 = Trim(Me.ListView1.SelectedItem.SubItems(6))
Me.Combo4 = Trim(Me.ListView1.SelectedItem.SubItems(7))
Me.Combo1.Locked = True
Me.Combo3.Locked = True
Me.Combo5.Locked = True
Me.Combo6.Locked = True
Code = Trim(Me.Combo2)
NameEng = Trim(Me.Combo7)
NameArab = Trim(Me.Combo8)
SerialNo = Trim(Me.Combo9)
Modelno = Trim(Me.Combo4)


Me.SSTab1.SetFocus
SendKeys "{Left}"

End Sub

Private Sub xFind_Click()
Me.Label20.Visible = True
Me.Combo10.Visible = True
Me.Command3.Visible = True
Me.Label16.Visible = False
Me.Combo10.SetFocus
End Sub

Private Sub xPrint_Click()
ListofRegisteredAsset.Show
End Sub

Private Sub xREfresh_Click()
Call ViewList
End Sub

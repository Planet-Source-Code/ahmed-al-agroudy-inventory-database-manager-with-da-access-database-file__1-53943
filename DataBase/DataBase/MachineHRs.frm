VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form MachineHRs 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set Machine Hour Used"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   6090
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6075
      _ExtentX        =   10716
      _ExtentY        =   4895
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   1058
      TabMaxWidth     =   3528
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arabic Transparent"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "New/Edit Hour Used ÇÏÎÇá ÓÇÚÇÊ ÇáãÓÊÎÏã "
      TabPicture(0)   =   "MachineHRs.frx":0000
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
      Tab(0).Control(6)=   "MaskEdBox1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Command1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Combo1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Combo2"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "List of Hours Used  ÞÇÆãÉ ÓÇÚÇÊ ÇáãÓÊÎÏã"
      TabPicture(1)   =   "MachineHRs.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "ListView1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   -73800
         Style           =   1  'Simple Combo
         TabIndex        =   6
         Top             =   1590
         Width           =   3255
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2050
         Left            =   45
         TabIndex        =   8
         Top             =   645
         Width           =   5955
         _ExtentX        =   10504
         _ExtentY        =   3625
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "a"
            Text            =   "No of Hour Used"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "b"
            Text            =   "Date Used"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Key             =   "c"
            Text            =   "Used by"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   -73800
         Style           =   1  'Simple Combo
         TabIndex        =   2
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   325
         Left            =   -72480
         TabIndex        =   7
         Top             =   2160
         Width           =   1215
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   330
         Left            =   -73800
         TabIndex        =   4
         Top             =   1200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label6 
         Caption         =   "ÇÓÊÎÏã ÈæÇÓØÉ "
         Height          =   255
         Left            =   -70440
         TabIndex        =   11
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "ÊÇÑíÎ ÇáãÓÊÎÏã"
         Height          =   255
         Left            =   -72360
         TabIndex        =   10
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "ÓÇÚÇÊ ÇáãÓÊÎÏã"
         Height          =   255
         Left            =   -72360
         TabIndex        =   9
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Used by"
         Height          =   375
         Left            =   -74760
         TabIndex        =   5
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Date Used"
         Height          =   375
         Left            =   -74790
         TabIndex        =   3
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Hour Used"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   1
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.Menu Main 
      Caption         =   "Main"
      Visible         =   0   'False
      Begin VB.Menu xAddnew 
         Caption         =   "Add New Hour Used"
      End
      Begin VB.Menu xEdit 
         Caption         =   "Edit"
      End
   End
End
Attribute VB_Name = "MachineHRs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PrevHrBeforEdit As Integer
Dim MItem As ListItem

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.MaskEdBox1.SetFocus
End If
    
End Sub

Private Sub Combo2_Change()
If KeyAscii = 13 Then
    Me.Command1.SetFocus
End If
End Sub

Private Sub Command1_Click()
Dim rsAssetSetup As New ADODB.Recordset
Dim rsActivityBase As New ADODB.Recordset
Dim MachineHrUsed As New ADODB.Recordset
ASsetNo = AssetSetup.ListView1.SelectedItem
    
    rsActivityBase.Open "Select * from ActivitybaseAssets where AssetNo=" & "'" & ASsetNo & "'" & "and DateUsed=" & "'" & Trim(Me.MaskEdBox1.Text) & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
    If rsActivityBase.EOF = True Then
        msg = MsgBox("Do you want to save your entry?åá ÊÑíÏ ÍÝÙ ÇáÈíÇäÇÊ ", vbQuestion + vbYesNo, "PLease confirm ãä ÝÖáß")
        If msg = vbYes Then
            With rsActivityBase
                .addnew
                !ASsetNo = ASsetNo
                !AssetCode = (AssetSetup.ListView1.SelectedItem.SubItems(1))
                !AssetName = Trim(AssetSetup.ListView1.SelectedItem.SubItems(2))
                !HoursUsed = Me.Combo1
                !Dateused = Me.MaskEdBox1.Text
                !Usedby = Me.Combo2
                .Update
                 MachineHrUsed.Open "Update ASsetSetup set MachineHourUsed = " & "'" & Val(Trim(Me.Combo1)) + MachineHourUsed & "'" & " where AssetNo=" & "'" & ASsetNo & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
            End With
          Else
             Exit Sub
         End If
             
                        
      Else
       msg = MsgBox("Do you want to save changes? åá ÊÑíÏ ÍÝÙ ÇáÊÛíÑÇÊ ", vbQuestion + vbYesNo, "PLease confirm ãä ÝÖáß")
        If msg = vbYes Then
        With rsActivityBase
            !ASsetNo = ASsetNo
            !AssetCode = (AssetSetup.ListView1.SelectedItem.SubItems(1))
            !AssetName = Trim(AssetSetup.ListView1.SelectedItem.SubItems(2))
            !HoursUsed = Me.Combo1
            !Dateused = Me.MaskEdBox1.Text
            !Usedby = Me.Combo2
            .Update
            MachineHrUsed.Open "Select * from ASsetSetup where AssetNo=" & "'" & ASsetNo & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
            With MachineHrUsed
                !MachineHourUsed = (!MachineHourUsed - PrevHrBeforEdit) + Val(Me.Combo1)
                .Update
            End With
         End With
        End If
      End If
    
    Unload Me


End Sub

Private Sub Command2_Click()
  Unload Me
End Sub

Private Sub Form_Activate()
Me.Combo1.SetFocus
End Sub

Private Sub Form_Load()
Dim HourUsed As New ADODB.Recordset
ASsetNo = AssetSetup.ListView1.SelectedItem
HourUsed.Open "Select * from ActivitybaseAssets where Assetno=" & "'" & ASsetNo & "'" & " order by dateUsed", constring, adOpenKeyset, adLockPessimistic, adCmdText
Do Until HourUsed.EOF = True
    Set MItem = Me.ListView1.ListItems.Add(, , HourUsed!HoursUsed)
    MItem.SubItems(1) = HourUsed!Dateused
    MItem.SubItems(2) = HourUsed!Usedby
    HourUsed.MoveNext
Loop
Me.caption = Me.caption & "-" & AssetSetup.ListView1.SelectedItem.SubItems(2)
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)

Me.Combo1 = Me.ListView1.SelectedItem
Me.MaskEdBox1 = Me.ListView1.SelectedItem.SubItems(1)
Me.Combo2 = Me.ListView1.SelectedItem.SubItems(2)
PrevHrBeforEdit = Val(Me.Combo1)

End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = 2 Then
    PopupMenu main
End If
End Sub

Private Sub MaskEdBox1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.Combo2.SetFocus
End If
    
End Sub

Private Sub MaskEdBox1_LostFocus()
If Me.MaskEdBox1.Text = "__/__/____" Then
    Exit Sub
End If
cDay = Val(Left(Me.MaskEdBox1.Text, 2))
cMonth = Val(Mid(Me.MaskEdBox1.Text, 4, 2))
cYear = Val(Right(Me.MaskEdBox1.Text, 4))
If cDay > 31 Or cDay < 1 Then
    mess = MsgBox("Invalid Date ÈíÇäÇÊ ÎØÇð", vbInformation + vbOKOnly, "Message ÑÓÇáÉ")
    Me.MaskEdBox1.SetFocus
  ElseIf cMonth > 13 Or cMonth < 1 Then
    mess = MsgBox("Invalid Month ÔåÑ ÎØÇð", vbInformation + vbOKOnly, "Message ÑÓÇáÉ")
    Me.MaskEdBox1.SetFocus
ElseIf cYear < 1900 Or cYear > Year(Date) Then
    mess = MsgBox("Invalid Year ÓäÉ ÎØÇð", vbInformation + vbOKOnly, "Message ÑÓÇáÉ")
    Me.MaskEdBox1.SetFocus
End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If Trim(Me.SSTab1.caption) = "List of Hour Used" Then
    Me.ListView1.SetFocus
 Else
 Me.Combo1.SetFocus
End If
End Sub

Private Sub xAddnew_Click()
SendKeys "{Left}"
End Sub

Private Sub xedit_Click()
PrevHrBeforEdit = Me.Combo1
'Me.Combo1 = Me.ListView1.SelectedItem
'Me.MaskEdBox1 = Me.ListView1.SelectedItem.SubItems(1)
'Me.Combo2 = Me.ListView1.SelectedItem.Selected(2)
SendKeys "{Left}"

End Sub

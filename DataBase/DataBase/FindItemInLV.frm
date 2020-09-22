VERSION 5.00
Begin VB.Form FindItemInTV 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find "
   ClientHeight    =   1185
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "FindItemInLV.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1185
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option2 
      Appearance      =   0  'Flat
      Caption         =   "AccountN&ame"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   1440
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      Caption         =   "AccountN&umber"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2400
      TabIndex        =   3
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Find ÇáÈÍË"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Find by"
      Height          =   615
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   4455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Look at"
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   4455
      Begin VB.OptionButton Option6 
         Caption         =   "Level&4"
         Height          =   255
         Left            =   3480
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Level&3"
         Height          =   255
         Left            =   2400
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Level&2"
         Height          =   255
         Left            =   1320
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Level&1"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Type  first few letters of the Account Name your looking for ÇßÊÈ ÇáÇÍÑÝ ÇáÇæáì ãä ÇÓã ÇáÍÓÇÈ ÇáÐí ÊÈÍË Úäå"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "Find What ÇáÈÍË Úä"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   1695
   End
End
Attribute VB_Name = "FindItemInTV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Long
Dim X As Long
Dim Finditem As Boolean

Private Sub Command1_Click()
Dim xKey As String
Static CancelFind As Boolean

xKey = Trim(Me.Text1.Text)
cLen = Len(xKey)

If CancelFind Then
   Me.Command1.caption = "Continue"
   CancelFind = False
  Else
    Me.Command1.caption = "Stop"
    CancelFind = True
    If i = AccTreeView.TreeView1.Nodes.Count And Finditem = True Then
         xmsg = MsgBox("The specified region has been searched, Nothing follows. " & vbCrLf & _
                     "ÎÕíÕ åÐÇ ÇáÍÞá áÈÍË áÇÔíÁ ", vbInformation + vbOKOnly, "Message ÑÓÇáÉ ")
         i = 1
         Exit Sub
    End If
   For i = X To AccTreeView.TreeView1.Nodes.Count
       AccTreeView.TreeView1.Nodes.Item(i).Selected = True
       X = AccTreeView.TreeView1.SelectedItem.Index
       xItem = Trim(Mid(AccTreeView.TreeView1.Nodes.Item(i).Text, 1, 15))
       n = 0
       For n = 1 To Len(xItem)
           
           If Mid(xItem, n, 1) = "(" Then Exit For
           If Mid(xItem, n, 1) <> "-" Then
             ItemToFind = ItemToFind & Mid(xItem, n, 1)
             Else
             ItemToFind = ""
           End If
           
        Next
        If Left(Trim(UCase(ItemToFind)), cLen) = Trim(UCase(xKey)) Then
           i = i + 1
           Beep
           Finditem = True
           On Error Resume Next
           CancelFind = False
           FindItemInTV.Command1.caption = "Find Next"
           AccTreeView.TreeView1.Nodes.Item(i).Selected = True
           X = AccTreeView.TreeView1.SelectedItem.Index
           AccTreeView.TreeView1.Nodes.Item(i - 1).Bold = True
           On Error GoTo 0
        Exit Sub
        End If
        AccTreeView.TreeView1.Nodes.Item(i).Expanded = False
       DoEvents
     If CancelFind = False Then
         Exit For
         Exit Sub
        End If
    Next
  If CancelFind = False Then
   CancelFind = False
   Exit Sub
  End If
  If Finditem = False Then
    i = 1
    CancelFind = False
    xmsg = MsgBox("Walang makita pare(Mafi Shof Sadik)áÇíæÌÏ  ", vbInformation + vbOKOnly, "Message ÑÓÇáÉ ")
    AccTreeView.TreeView1.Nodes.Item(i).Selected = True
   Else
   i = 1
   CancelFind = False
   xmsg = MsgBox("The specified region has been searched ÎÕíÕ åÐÇ ÇáÍÞá áÈÍË  ", vbInformation + vbOKOnly, "MessageÑÓÇáÉ")
   AccTreeView.TreeView1.Nodes.Item(i).Selected = True
  End If
End If
End Sub

Private Sub Form_Activate()


Me.Option1.SetFocus
Me.Option6.Value = True
End Sub

Private Sub Option1_Click()
Me.Text1.SetFocus
End Sub

Private Sub Option2_Click()
Me.Text1.SetFocus
End Sub

Private Sub Option3_Click()
Me.Text1.SetFocus
End Sub

Private Sub Option4_Click()
Me.Text1.SetFocus
End Sub

Private Sub Option5_Click()
Me.Text1.SetFocus
End Sub

Private Sub Option6_Click()
Me.Text1.SetFocus
End Sub

Private Sub Text1_Change()
X = 1 'AccTreeView.TreeView1.SelectedItem.Index
If Len(Me.Text1.Text) = 0 Then
    Me.Command1.Enabled = False
  Else
     Me.Command1.Enabled = True
End If
If Me.Text1 = "" Then
  Me.Command1.Enabled = False
 Else
 Me.Command1.Enabled = True
End If
Me.Command1.caption = "&Find"
End Sub

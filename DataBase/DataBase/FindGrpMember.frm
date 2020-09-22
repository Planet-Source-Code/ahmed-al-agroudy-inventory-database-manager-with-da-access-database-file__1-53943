VERSION 5.00
Begin VB.Form FindGrpMember 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   Icon            =   "FindGrpMember.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   6135
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Find by"
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   5655
      Begin VB.OptionButton Option3 
         Caption         =   "Account Name Arab"
         Height          =   255
         Left            =   3720
         TabIndex        =   7
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton Option2 
         Caption         =   "AccountName in Eng"
         Height          =   255
         Left            =   1680
         TabIndex        =   6
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         Caption         =   "AccountNo"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Find ÇáÈÍË"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Find What ãÇåæ ÇáÈÍË"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Type the first few letters of an item you're looking for.ÇßÊÈ ÇáÇÍÑÝ ÇáÇæáì ãä ÇÓã Çá "
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
      TabIndex        =   2
      Top             =   720
      Width           =   4215
   End
End
Attribute VB_Name = "FindGrpMember"
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
If CancelFind Then
   Me.Command1.caption = "Continue"
   CancelFind = False
  Else
    Me.Command1.caption = "Stop"
    CancelFind = True
    If i = Grouping.TreeView1.Nodes.Count And Finditem = True Then
         xmsg = MsgBox("The specified region has been searched, Nothing follows. " & vbCrLf & _
                     "ÎÕíÕ åÐÇ ÇáÍÞá áÈÍË áÇÔíÁ ", vbInformation + vbOKOnly, "Message ÑÓÇáÉ ")
         i = 1
         Exit Sub
    End If
   i = 1
   X = 1
   For i = X To Grouping.TreeView1.Nodes.Count
       Grouping.TreeView1.Nodes.Item(i).Selected = True
       X = Grouping.TreeView1.SelectedItem.Index
       If Me.Option1.Value = True Then
         xItem = Trim(Mid(Grouping.TreeView1.Nodes.Item(i).Text, 1, 12))
       ElseIf Me.Option2.Value = True Then
         PartialItem = Trim(Mid(Grouping.TreeView1.Nodes.Item(i).Text, 14, 100))
         cLen = InStr(14, PartialItem, "\", vbTextCompare)
         clen1 = InStr(14, PartialItem, "\", vbTextCompare)
         On Error Resume Next
         xItem = Trim(Left(PartialItem, cLen - 1))
         On Error GoTo 0
       End If
       ItemToFind = xItem
       n = 0
       'For n = 1 To Len(xItem)
           'ItemToFind = xItem
           'If Mid(xItem, n, 1) = "(" Then Exit For
           'If Mid(xItem, n, 1) <> "-" Then
           '  ItemToFind = ItemToFind & Mid(xItem, n, 1)
           '  Else
           '  ItemToFind = ""
           'End If
           
       ' Next
        If Left(Trim(UCase(ItemToFind)), Len(xKey)) = Trim(UCase(xKey)) Then
           i = i + 1
           Beep
           Finditem = True
           On Error Resume Next
           CancelFind = False
           Me.Command1.caption = "Find Next"
          Grouping.TreeView1.Nodes.Item(i).Selected = True
           X = Grouping.TreeView1.SelectedItem.Index
           Grouping.TreeView1.Nodes.Item(i - 1).Bold = True
           ItemToFind = ""
           On Error GoTo 0
        Exit Sub
        End If
        Grouping.TreeView1.Nodes.Item(i).Expanded = False
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
    Grouping.TreeView1.Nodes.Item(i).Selected = True
   Else
   i = 1
   CancelFind = False
   xmsg = MsgBox("The specified region has been searched ÎÕíÕ åÐÇ ÇáÍÞá áÈÍË  ", vbInformation + vbOKOnly, "MessageÑÓÇáÉ")
   Grouping.TreeView1.Nodes.Item(i).Selected = True
  End If
End If

End Sub

Private Sub Text1_Change()
If Me.Text1.Text = "" Then
    Me.Command1.Enabled = False
   Else
     Me.Command1.Enabled = True
End If
End Sub

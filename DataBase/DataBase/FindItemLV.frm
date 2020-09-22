VERSION 5.00
Begin VB.Form FindItemLV 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find Item ÈÍË Úä ÇáÕäÝ"
   ClientHeight    =   870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3840
   BeginProperty Font 
      Name            =   "Arabic Transparent"
      Size            =   9.75
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FindItemLV.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   870
   ScaleWidth      =   3840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton command1 
      Caption         =   "Find ÇÈÍË"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   2400
      TabIndex        =   2
      Top             =   360
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   345
      Left            =   120
      Style           =   1  'Simple Combo
      TabIndex        =   1
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Úä ãÇÐÇ ÊÑíÏ ÇáÈÍË"
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Find what?"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "FindItemLV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim intSelectedOption As Integer
Private Sub Combo1_KeyPress(KeyAscii As Integer)

If KeyAscii = 27 Then
    Unload Me
End If
End Sub

Private Sub Command1_Click()
  
   Dim strFindMe As String
   Dim itmFound As ListItem   ' FoundItem variable.
   intSelectedOption = lvwSubItem
   strFindMe = Trim(Me.Combo1)
   Set itmFound = GenJournalEntry.ListView2.Finditem(strFindMe, intSelectedOption, , lvwPartial)
   If itmFound Is Nothing Then  ' If no match, inform user and exit.
       xmsg = MsgBox("No item found ", vbExclamation + vbOKOnly, "Message")
       Exit Sub
    Else
        Unload Me
        itmFound.EnsureVisible
        itmFound.Selected = True   ' Select the ListItem.
       ' Return focus to the control to see selection.
        
        GenJournalEntry.ListView2.SetFocus
    End If
     
End Sub

Private Sub Form_Load()
On Error Resume Next
whatCol = Trim(GenJournalEntry.ListView2.ColumnHeaders(2))
SubItem = Trim(GenJournalEntry.ListView2.ColumnHeaders(8))
If whatCol = "" Then
    Exit Sub
End If
End Sub

Private Sub Option1_Click()
On Error Resume Next
Me.Combo1.SetFocus
End Sub

Private Sub Option2_Click()
Me.Combo1.SetFocus
End Sub

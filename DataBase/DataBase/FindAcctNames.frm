VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FindAcctNames 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find AccountNames  ÇáÈÍË Úä ÍÓÇÈ"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9630
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FindAcctNames.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   9630
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   4170
      Width           =   9630
      _ExtentX        =   16986
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Find byæÍÏÉ ÇáÈÍË"
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
      TabIndex        =   8
      Top             =   0
      Width           =   8895
      Begin VB.OptionButton Option2 
         Caption         =   "Arabic NameÇáÇÓã ÈÇáÚÑÈí"
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
         Left            =   4200
         TabIndex        =   10
         Top             =   200
         Width           =   2775
      End
      Begin VB.OptionButton Option1 
         Caption         =   "English NameÇáÇÓã ÈÇáÇäÌáíÒí"
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
         Left            =   240
         TabIndex        =   9
         Top             =   200
         Value           =   -1  'True
         Width           =   2895
      End
   End
   Begin RichTextLib.RichTextBox RichTextBox2 
      Height          =   30
      Left            =   2040
      TabIndex        =   6
      Top             =   4680
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   53
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"FindAcctNames.frx":0442
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   20
      Style           =   1  'Simple Combo
      TabIndex        =   4
      Top             =   860
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&CloseÇáÛáÞ"
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
      Left            =   5160
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command21 
      Caption         =   "&Find ÇáÈÍË"
      Default         =   -1  'True
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
      Left            =   3600
      TabIndex        =   2
      Top             =   840
      Width           =   1455
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2985
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   5265
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      PictureAlignment=   1
      _Version        =   393217
      Icons           =   "ImageList2"
      SmallIcons      =   "ImageList2"
      ForeColor       =   -2147483640
      BackColor       =   16777215
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
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   375
      Left            =   2280
      TabIndex        =   7
      Top             =   3600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393217
      TextRTF         =   $"FindAcctNames.frx":0510
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FindAcctNames.frx":05DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FindAcctNames.frx":0A30
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FindAcctNames.frx":0E82
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FindAcctNames.frx":12D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FindAcctNames.frx":15EE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   3480
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Type any few letters of the accountName you're looking for.ÇÏÎá ÇÍÑÝ ÞáíáÉ ãä ÇÓã ÇáÍÓÇÈ ÇáÐí ÊæÏ ÇáÈÍË Úäå:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   45
      TabIndex        =   1
      Top             =   600
      Width           =   7695
   End
End
Attribute VB_Name = "FindAcctNames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sqltable As Boolean
Dim mConnek As Boolean
Dim rst As New ADODB.Recordset
Dim xtable As String
Dim xxtable As String
Dim CON1 As ADODB.Connection
Dim con11 As ADODB.Connection
Dim xxClass As HabitatClass
Dim COPTION As Integer
Dim MItem As ListItem
Dim rstt As New ADODB.Recordset
Dim i As Integer
Dim xcol As ColumnHeader
Dim constring As String
Dim Rec As Long
Dim CurrSearch As String
Dim PrevSearch As String
Dim ArabicSelect As Boolean
Dim xvalue As String

Private Sub Combo1_Change()
If Trim(Me.Combo1) = "" Then
    Me.Command21.Enabled = False
  Else
    Me.Command21.Enabled = True
 End If
 If Trim(Me.Combo1) <> PrevSearch Then
     Command21.caption = "&Find ÇáÈÍË"
    End If
End Sub

Private Sub Combo1_GotFocus()
Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(Combo1.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
   'SendKeys "{Down}"
End Sub


Private Sub Combo1_LostFocus()
CurrSearch = Me.Combo1
Rec = 0
End Sub

Private Sub Command1_Click()
Me.Hide
'Unload Me
End Sub

Private Sub Command21_Click()
Dim strFindMe As String
Dim intSelectedOption As String
Dim itmFound As ListItem   ' FoundItem variable.
intSelectedOption = lvwText
Dim Tobeseach As String
Dim i As Integer
Dim FoundPos As Integer
Dim FoundLine As Integer
Dim words As String
Dim xwords As String
Me.Command1.Enabled = False
Me.StatusBar1.Panels(1).Text = Str(Rec) & " items(s) found"
Static CancelFind As Boolean
'CurrSearch = Trim(Me.Combo1)
If ArabicSelect = True And CancelFind = False And Rec = 0 Then
  If CurrSearch <> PrevSearch Then
    rst.MoveFirst
    Me.ListView1.ListItems.clear
  End If
 ElseIf ArabicSelect = False And CancelFind = False And Rec = 0 Then
   If CurrSearch <> PrevSearch Then
    rst.MoveFirst
    Me.ListView1.ListItems.clear
   End If
End If

If CancelFind Then
   CancelFind = False
  Else
    
       Me.Command21.caption = "Stop"
       CancelFind = True
       PrevSearch = Trim(Me.Combo1)
       If Me.Option2.Value = True Then
        ArabicSelect = True
        Else
        ArabicSelect = False
       End If
       Tobeseach = Trim(Me.Combo1)
       Dim acctnames As String
       Dim acctNo As String
       If Rec = 0 Then
          Me.ListView1.ListItems.clear
          Set MItem = Me.ListView1.ListItems.Add(, , "Searching...")
       End If
       With rst
          Do Until .EOF = True
           
           If Me.Option1 = True Then
            Me.RichTextBox1.Text = IIf(IsNull(rst!accountnameeng) = True, "", rst!accountnameeng)
            Else
            Me.RichTextBox1.Text = rst!accountnamearab
           End If
           FoundPos = Me.RichTextBox1.find(Tobeseach, , , rtfCFText)
           words = Me.RichTextBox1.find(Tobeseach, , , rtfCFText)
           If words >= 0 Then
             Rec = Rec + 1
             If Rec = 1 Then
                Me.ListView1.ListItems.clear
             End If
             Me.StatusBar1.Panels(1).Text = Str(Rec) & " items(s) found"
             acctnames = rst!accountnameeng
             acctNo = rst!AccountCode
             Set MItem = Me.ListView1.ListItems.Add(, , acctnames, , IIf(rst!Active = 1, 1, 4))
             MItem.SubItems(1) = acctNo
             MItem.SubItems(2) = rst!accountnamearab
             MItem.SubItems(3) = IIf(rst!Active = 1, "Active", "Inactive")
             Set itmFound = Me.ListView1.Finditem(acctnames, intSelectedOption, , lvwPartial)
             If itmFound Is Nothing Then  ' If no match, inform user and exit.
               Else
                 itmFound.EnsureVisible
             End If
            End If
           .MoveNext
           DoEvents
          If CancelFind = False Then
             Exit Do
            Exit Sub
          End If
         Loop
          If rst.EOF = False And CancelFind = False Then
            Command21.caption = "Continue ÇÓÊãÑÇÑ"
            Me.Command1.Enabled = True
            CancelFind = False
             Exit Sub
           Else
           Command21.caption = "&Find ÇáÈÍË"
           CancelFind = False
           Me.Command1.Enabled = True
           On Error Resume Next
           Me.Combo1.SetFocus
           If Rec = 0 Then
              Me.ListView1.ListItems.clear
              Set MItem = Me.ListView1.ListItems.Add(, , "Kalas Habebe, Mafi shof!")
           End If
           rst.MoveFirst
           Rec = 0
           Exit Sub
         End If
        End With
               
       
    
End If
'CancelFind = True
End Sub
Private Sub Command2_Click()
 Unload Me
End Sub
Private Sub Form_Activate()
Me.Combo1.SetFocus
End Sub
Private Sub Form_Load()
ListView1.ListItems.clear
ListView1.ColumnHeaders.clear
Set xcol = Me.ListView1.ColumnHeaders.Add(, , "Account NAme(Eng)", 2800)
Set xcol = Me.ListView1.ColumnHeaders.Add(, , "Account Number", 1580)
Set xcol = Me.ListView1.ColumnHeaders.Add(, , "Account NAme(Arab)", 2900)
Set xcol = Me.ListView1.ColumnHeaders.Add(, , "Status", 1200)
Me.ListView1.ColumnHeaders(2).Alignment = lvwColumnRight
Me.ListView1.ColumnHeaders(3).Alignment = lvwColumnRight
Dim CON1 As ADODB.Connection
Set CON1 = New ADODB.Connection
Left = 50
Set rst = New ADODB.Recordset
constring = "dsN=fINANCE;UID=SA;PWD=;"
CON1.Open constring
rst.Open "FinanceMaster", CON1, adOpenDynamic, adLockOptimistic, adCmdTable

'Dim xClass As HabitatClass
'Dim xxClass As HabitatClass
'Set xxClass = New HabitatClass
'
'Set xClass = New HabitatClass
'Set rst = New ADODB.Recordset
'Set con1 = New ADODB.Connection
'Set con11 = New ADODB.Connection
'
'Set rstAmount = New ADODB.Recordset
'xtable = "FinanceMaster"
'xClass.GetTables rst, con1, xtable, conString, SQLtable
'Do Until rst.EOF
'     On Error Resume Next
'     rst.MoveNext
'    On Error GoTo 0
'  Loop
''Rst.Close

End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Command21_Click
  Combo1.SetFocus
  End If
End Sub

Private Sub TxtAmount_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim mess As Integer
If Me.Command21.caption = "Stop" Then
    Call Command21_Click
    mess = MsgBox("Do you want to cancel searching?", vbQuestion + vbYesNo + vbDefaultButton2, "Please confirm")
    If mess = vbNo Then
        Cancel = -1
        Call Command21_Click
      Else
      Me.Command21.caption = "&Find ÇáÈÍË"
    End If
End If
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Me.ListView1.SortKey = ColumnHeader.Index - 1
Me.ListView1.Sorted = True

End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim an As String
Dim ano As String
Dim anoarab As String
an = Me.ListView1.SelectedItem.Text & "\" & Me.ListView1.SelectedItem.SubItems(2)
ano = Me.ListView1.SelectedItem.SubItems(1)
anoarab = Me.ListView1.SelectedItem.SubItems(2)
Dim catName As String
Dim Prevcap As String
Dim acctNo As String
acctNo = Trim(ano)
Prevcap = Trim(Me.caption)
Call DisplayCats(Prevcap, acctNo, catName)
Me.ListView1.ToolTipText = "Classification:" & catName


If FindAcctNAme = True Then
If GenJournalEntry.Command3 = True Or InvJournalEntry.Command3 = True Or FrmPaymentAnalysis.Command1 = True Or _
    BankTransaction1.Command3 = True Then
  
     GenJournalEntry.Combo6 = ano
     GenJournalEntry.Combo12 = an
     InvJournalEntry.Combo6 = ano
     InvJournalEntry.Combo12 = an
     FrmPaymentAnalysis.txtAccNo = ano
     FrmPaymentAnalysis.txtPartic = an
     BankTransaction1.Combo6 = ano
     BankTransaction1.Combo12 = an
     Else
     GenJournalEntry.Combo7 = ano
     GenJournalEntry.Combo13 = an
     InvJournalEntry.Combo7 = ano
     InvJournalEntry.Combo13 = an
     BankTransaction1.Combo7 = ano
     BankTransaction1.Combo13 = an
     FrmPaymentAnalysis.txtDBAccNo = ano
     FrmPaymentAnalysis.txtDBPartic = an
    End If
   End If
   ' AddNewAccts.Combo2 = Me.ListView1.SelectedItem.Text
  '  AddNewAccts.Combo3 = anoarab
 
 
End Sub

Private Sub Option1_Click()
PrevSearch = ""
ArabicSelect = False
Me.Combo1.RightToLeft = False
Me.Combo1.SetFocus
End Sub

Private Sub Option2_Click()
PrevSearch = ""
ArabicSelect = True
Me.Combo1.RightToLeft = True
Me.ListView1.ColumnHeaders(2).Alignment = lvwColumnRight
Me.Combo1.SetFocus
End Sub

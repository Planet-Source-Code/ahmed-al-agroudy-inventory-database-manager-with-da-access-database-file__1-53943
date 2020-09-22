VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form NewAccts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Country ÈáÏ"
   ClientHeight    =   5130
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   9975
   Icon            =   "NewAccts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   9975
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BackColor       =   &H0071B9B4&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "  Please wait...Sadik"
      Top             =   2400
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Close ÛáÞ "
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
      Left            =   8640
      TabIndex        =   3
      Top             =   4680
      Width           =   1200
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Next > ÇãÇã"
      Default         =   -1  'True
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
      Left            =   7320
      TabIndex        =   2
      Top             =   4680
      Width           =   1200
   End
   Begin VB.CommandButton Command1 
      Caption         =   "< &Back ÎáÝ "
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
      Left            =   6120
      TabIndex        =   1
      Top             =   4680
      Width           =   1200
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4215
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   7435
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
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
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList ImageList1 
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
            Picture         =   "NewAccts.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NewAccts.frx":0894
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NewAccts.frx":0CE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NewAccts.frx":1138
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "NewAccts.frx":158A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
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
      TabIndex        =   6
      Top             =   4680
      Width           =   3495
   End
   Begin VB.Label Label2 
      Caption         =   "Click right mouse button to select options ÇÖÛØ íãíä ÇáÝÇÑÉ áÎÊíÇÑ ÇáæÙÇÆÝ"
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
      TabIndex        =   5
      Top             =   120
      Width           =   6375
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   1800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Menu Main 
      Caption         =   "Main"
      Visible         =   0   'False
      Begin VB.Menu AddItem 
         Caption         =   "Add Item... ÇÖÇÝÉ ÕäÝ  "
      End
      Begin VB.Menu delete 
         Caption         =   "&Delete... ÇáÛÇÁ "
      End
      Begin VB.Menu bar1 
         Caption         =   "-"
      End
      Begin VB.Menu Rename 
         Caption         =   "&Rename English Name"
         Visible         =   0   'False
      End
      Begin VB.Menu RenameArabName 
         Caption         =   "&Rename ÇÚÏÉ ÇáÇÓã "
      End
      Begin VB.Menu bar2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu xRefresh 
         Caption         =   "&Refresh ÊÍÏíË"
      End
      Begin VB.Menu xPrint 
         Caption         =   "Print... ØÈÇÚÉ"
      End
   End
End
Attribute VB_Name = "NewAccts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Dim MItem As ListItem
Dim oldstring As String
Dim LevelCode As String
Dim xcol As ColumnHeader
Dim constring As String


Private Sub addnew_Click()
End Sub

Private Sub AddItem_Click()
 If UCase(UserRole) <> UCase("admin") Then
  If UCase(Left(UserRole, 5)) <> UCase(Left("Accoun", 5)) Then
   mess = MsgBox("You are not authorized to add new item", vbExclamation + vbOKOnly, "Message")
   Exit Sub
 End If
 End If
'creating main account in the level 1-2 is not allowed
If FormNo < 4 Then
    Exit Sub
End If
   
   
   If cErr = 380 Then
       cErr = 0
       mess = MsgBox("I'm going to click first the Back Button now in order to refresh the list" & vbCrLf & _
                      "ÓæÝ ÇÐåÈ áÖÛØ ÇæáÇð ÒÑ áÊÍÏíË  ÇáÞÇÆãÉ", vbInformation + vbOKOnly, "Message ÑÓÇáÉ")
       Call Command1_Click
       Call Command2_Click
       Exit Sub
    End If

    
 
 'nextCode# will be blank because zero items in listview
 'we're going to place a value for the next level
 If Me.ListView1.ListItems.Count = 0 Then
  If FormNo + 1 = 2 Then
        TopLevelCode = "X"
        TopLevelName = "X"
    ElseIf FormNo + 1 = 3 Then
        Level1Name = "X"
        Level1Code = "X"
    ElseIf FormNo + 1 = 4 Then
        Level2Name = "X"
        Level2Code = "X"
    ElseIf FormNo + 1 = 5 Then
        level3Name = "X"
        level3Code = "X"
    ElseIf FormNo + 1 = 6 Then
        level4Name = "X"
        level4Code = "X"
    ElseIf FormNo + 1 = 7 Then
        level5Name = "X"
        level5Code = "X"
    End If
    AddNewAccts.Show 1
    Exit Sub
 End If
' Else
'  If FormNo + 1 = 2 Then
'        TopLevelCode = Trim(Right(Me.ListView1.SelectedItem.SubItems(1), 1))
'        TopLevelName = (Me.ListView1.SelectedItem.Text)
'    ElseIf FormNo + 1 = 3 Then
'        Level1Name = (Me.ListView1.SelectedItem.Text)
'        Level1Code = Trim(Right(Me.ListView1.SelectedItem.SubItems(1), 1))
'    ElseIf FormNo + 1 = 4 Then
'        Level2Name = (Me.ListView1.SelectedItem.Text)
'        Level2Code = Trim(Right(Me.ListView1.SelectedItem.SubItems(1), 1))
'    ElseIf FormNo + 1 = 5 Then
'        level3Name = (Me.ListView1.SelectedItem.Text)
'        level3Code = Trim(Right(Me.ListView1.SelectedItem.SubItems(1), 1))
'    ElseIf FormNo + 1 = 6 Then
'        level4Name = (Me.ListView1.SelectedItem.Text)
'        level4Code = Trim(Right(Me.ListView1.SelectedItem.SubItems(1), 1))
'    ElseIf FormNo + 1 = 7 Then
'        level5Name = (Me.ListView1.SelectedItem.Text)
'        level5Code = Trim(Right(Me.ListView1.SelectedItem.SubItems(1), 1))
'    End If
'End If
Dim CanICreateNEwAcct As Boolean
CanICreateNEwAcct = True
cTotalItems = Me.ListView1.ListItems.Item(Me.ListView1.ListItems.Count).SubItems(2)  'Me.ListView1.ListItems.Count
If FormNo = 0 Then
  AddNewAccts.Show 1
ElseIf FormNo - 1 = 1 Then
 If cTotalItems = 9 Then
    CanICreateNEwAcct = False
    UnabletoOPenNewAcct
    Exit Sub
   Else
    AddNewAccts.Show 1
  End If


ElseIf FormNo - 1 = 2 Then
  If cTotalItems = 9 Then
    CanICreateNEwAcct = False
    UnabletoOPenNewAcct
    Exit Sub
   Else
    AddNewAccts.Show 1
  End If
  
 ElseIf FormNo - 1 = 3 Then
  If cTotalItems = 99 Then
    CanICreateNEwAcct = False
    UnabletoOPenNewAcct
    Exit Sub
   Else
    AddNewAccts.Show 1
  End If
 
ElseIf FormNo - 1 = 4 Then
  If cTotalItems = 99 Then
    CanICreateNEwAcct = False
    UnabletoOPenNewAcct
    Exit Sub
   Else
    AddNewAccts.Show 1
  End If
  
ElseIf FormNo - 1 = 5 Then
  If cTotalItems = 99 Then
    CanICreateNEwAcct = False
    UnabletoOPenNewAcct
    Exit Sub
   Else
    AddNewAccts.Show 1
  End If
  
 ElseIf FormNo - 1 = 6 Then
  If cTotalItems = 999 Then
    CanICreateNEwAcct = False
    UnabletoOPenNewAcct
    Exit Sub
   Else
    AddNewAccts.Show 1
  End If
  
End If
End Sub
Sub UnabletoOPenNewAcct()
   mess = MsgBox("Unable to open new Account", vbExclamation + vbOKOnly, "Message")
End Sub
Private Sub Command1_Click()

If FormNo = 1 Then
    TopLevelCode = ""
ElseIf FormNo = 2 Then
    TopLevelCode = ""
    Level1Code = ""
    PrevItem1 = ""
    PrevItem2 = ""
    PrevItem3 = ""
    PrevItem4 = ""
    PrevItem5 = ""
ElseIf FormNo = 3 Then
    Level1Code = ""
    Level2Code = ""
    PrevItem2 = ""
    PrevItem3 = ""
    PrevItem4 = ""
    PrevItem5 = ""
ElseIf FormNo = 4 Then
    Level2Code = ""
    level3Code = ""
    PrevItem3 = ""
    PrevItem4 = ""
    PrevItem5 = ""
ElseIf FormNo = 5 Then
    level3Code = ""
    level4Code = ""
    PrevItem4 = ""
    PrevItem5 = ""
ElseIf FormNo = 6 Then
    level4Code = ""
    level5Code = ""
    PrevItem5 = ""
ElseIf FormNo = 7 Then
    level5Code = ""
    level6Code = ""
    
End If

FormNo = FormNo - 1
If FormNo = 0 Then
    Me.Command1.Enabled = False
End If
CancelAll = True
Unload Me
CancelAll = False
End Sub

Private Sub Command2_Click()
constring = "dsN=fINANCE;UID=SA;PWD=;"
Dim rst1 As New ADODB.Recordset
Dim WhatLevel As String

Set Newform = New NewAccts
Newform.Command1.Enabled = True
Newform.ListView1.ListItems.clear
Newform.ListView1.ColumnHeaders.clear

'PrevItem = Trim(Me.ListView1.SelectedItem.Text)
If Me.ListView1.ListItems.Count <> 0 Then
    If FormNo = 0 Then
     
     xCountry = Me.ListView1.SelectedItem.SubItems(1)
     xCountryNAMe = Me.ListView1.SelectedItem.Text
    End If

    'we don't want to add item if selected item is the main account
    If FormNo > 3 And FormNo < 7 Then
      xRem = Left(Me.ListView1.SelectedItem.SubItems(4), 4)
      If xRem = "Main" Then
        Beep
        mess = MsgBox("This is the Main Account and nothing follows" & vbCrLf & _
                       "åÐå Êßæä ÞÇÆãÉ ÇáÍÓÇÈÇÊ æáÇ íæÌÏ ÔíÁ ", vbInformation + vbOKOnly, "Message ÑÓÇáÉ ")
        Exit Sub
      End If
     End If
    
    'if it is the last level, stop of pressing next button
    If FormNo = 7 Then
        Beep
        mess = MsgBox("This is the Last Level of the Chart of Account" & vbCrLf & _
                      "åÐÇ íßæä ÇÎÑ ãÓÊæíÝí ÌÏæá ÇáÍÓÇÈÇÊ  ", vbInformation + vbOKOnly, "Message ÑÓÇáÉ")
        Exit Sub
    End If
    FormNo = FormNo + 1 ' increment form no.
    
    If FormNo = 2 Then
        PrevItem1 = Trim(Me.ListView1.SelectedItem.Text)
        TopLevelCode = Trim(Right(Me.ListView1.SelectedItem.SubItems(2), 2))
        TopLevelName = (Me.ListView1.SelectedItem.Text)
    ElseIf FormNo = 3 Then
        PrevItem2 = Trim(Me.ListView1.SelectedItem.Text)
        Level1Name = (Me.ListView1.SelectedItem.Text)
        Level1Code = Trim(Right(Me.ListView1.SelectedItem.SubItems(2), 2))
    ElseIf FormNo = 4 Then
        PrevItem3 = Trim(Me.ListView1.SelectedItem.Text)
        Level2Name = (Me.ListView1.SelectedItem.Text)
        Level2Code = Trim(Right(Me.ListView1.SelectedItem.SubItems(2), 2))
    ElseIf FormNo = 5 Then
        PrevItem4 = Trim(Me.ListView1.SelectedItem.Text)
        level3Name = (Me.ListView1.SelectedItem.Text)
        level3Code = Trim(Right(Me.ListView1.SelectedItem.SubItems(2), 2))
    ElseIf FormNo = 6 Then
        PrevItem5 = Trim(Me.ListView1.SelectedItem.Text)
        level4Name = (Me.ListView1.SelectedItem.Text)
        level4Code = Trim(Right(Me.ListView1.SelectedItem.SubItems(2), 2))
    ElseIf FormNo = 7 Then
        PrevItem6 = Trim(Me.ListView1.SelectedItem.Text)
        level5Name = (Me.ListView1.SelectedItem.Text)
        level5Code = Trim(Right(Me.ListView1.SelectedItem.SubItems(2), 2))
    End If
Else
    Exit Sub
End If
If FormNo = 8 Then
    Me.Command2.Enabled = False
    Exit Sub
End If

    

If FormNo = 1 Then
    WhatLevel = "TopLevel"
    rst1.Open "Select * from TopLevel where country =" & "'" & xCountry & "'" & "order by AccountNameEng ", constring, adOpenKeyset, adLockPessimistic, adCmdText

  ElseIf FormNo = 2 Then
    WhatLevel = "Level1"
    rst1.Open "Select * from Level1 where country =" & "'" & xCountry & "'" _
    & " and TOpLevelCode = " & "'" & TopLevelCode & "'" & "order by AccountNameEng ", constring, adOpenKeyset, adLockPessimistic, adCmdText

  ElseIf FormNo = 3 Then
    WhatLevel = "Level2"
    rst1.Open "Select * from Level2 where country =" & "'" & xCountry & "'" _
    & " and TOpLevelCode = " & " '" & TopLevelCode & "'" & " and Level1Code=" & "'" _
    & Level1Code & " '" & " order by AccountNameEng ", constring, adOpenKeyset, adLockPessimistic, adCmdText
 
    
  ElseIf FormNo = 4 Then
    WhatLevel = "Level3"
    rst1.Open "Select * from Level3 where country =" & "'" & xCountry & "'" _
    & " and TOpLevelCode = " & " '" & TopLevelCode & "'" _
    & " and Level1Code=" & "'" & Level1Code & " '" _
    & " and Level2Code = " & "'" & Level2Code & "'" _
    & " order by AccountNameEng ", constring, adOpenKeyset, adLockPessimistic, adCmdText
    
  ElseIf FormNo = 5 Then
    WhatLevel = "Level4"
    rst1.Open "Select * from Level4 where country =" & "'" & xCountry & "'" _
    & " and TOpLevelCode = " & " '" & TopLevelCode & "'" _
    & " and Level1Code=" & "'" & Level1Code & " '" _
    & " and Level2Code = " & "'" & Level2Code & "'" _
    & " and Level3Code = " & "'" & level3Code & "'" _
    & " order by AccountNameEng ", constring, adOpenKeyset, adLockPessimistic, adCmdText
    
  ElseIf FormNo = 6 Then
    WhatLevel = "Level5"
    rst1.Open "Select * from Level5 where country =" & "'" & xCountry & "'" _
    & " and TOpLevelCode = " & " '" & TopLevelCode & "'" _
    & " and Level1Code=" & "'" & Level1Code & " '" _
    & " and Level2Code = " & "'" & Level2Code & "'" _
    & " and Level3Code = " & "'" & level3Code & "'" _
    & " and Level4Code = " & "'" & level4Code & "'" _
    & " order by AccountNameEng ", constring, adOpenKeyset, adLockPessimistic, adCmdText
 

  ElseIf FormNo = 7 Then
    WhatLevel = "Level6"
    rst1.Open "Select * from Level6 where country =" & "'" & xCountry & "'" _
    & " and TOpLevelCode = " & " '" & TopLevelCode & "'" _
    & " and Level1Code=" & "'" & Level1Code & " '" _
    & " and Level2Code = " & "'" & Level2Code & "'" _
    & " and Level3Code = " & "'" & level3Code & "'" _
    & " and Level4Code = " & "'" & level4Code & "'" _
    & " and Level5Code = " & "'" & level5Code & "'" _
    & " order by AccountNameEng ", constring, adOpenKeyset, adLockPessimistic, adCmdText
 End If
 If UCase(WhatLevel) = UCase("TOpLevel") Then
    Newform.caption = "Financial Position Category ÝÆÇÊ æÙÇÆÝ ÇáÍÓÇÈÇÊ  "
 Else
    Newform.caption = WhatLevel & " - Under " & PrevItem1 & Chr(187) & PrevItem2 & Chr(187) & PrevItem3 & Chr(187) & PrevItem4 & Chr(187) & PrevItem5 & Chr(187) & PrevItem6
End If
    

Me.Command1.Enabled = True
Set xcol = Newform.ListView1.ColumnHeaders.Add(, "a", "AccountName(Englisi) ÇÓã ÇáÍÓÇÈ ÇäÌáíÒí ", 2500)
Set xcol = Newform.ListView1.ColumnHeaders.Add(, "b", "AccountName(Arabi)ÇÓã ÇáÍÓÇÈ ÈÇáÚÑÈí", 2630)
Set xcol = Newform.ListView1.ColumnHeaders.Add(, "c", "Sub-Level# ãÓÊæí ËÇäæí", 1000)
Set xcol = Newform.ListView1.ColumnHeaders.Add(, "d", "AccountCode ßæÏ ÇáÍÓÇÈÇÊ ", 1740)
Set xcol = Newform.ListView1.ColumnHeaders.Add(, "e", "Remarks ãáÇÍÙÇÊ ", 2000)
Newform.ListView1.ColumnHeaders(2).Alignment = lvwColumnRight
Newform.ListView1.ColumnHeaders(3).Alignment = lvwColumnRight
Newform.ListView1.ColumnHeaders(4).Alignment = lvwColumnCenter
Do Until rst1.EOF
    Set MItem = Newform.ListView1.ListItems.Add(, , Trim(rst1!accountnameeng), , IIf(Left(rst1!remarks, 4) = "Main", 4, 1))
    MItem.SubItems(1) = IIf(Trim(rst1!accountnamearab) = "", " ", Trim(rst1!accountnamearab))
    MItem.SubItems(2) = Trim(rst1!Code)
    MItem.SubItems(3) = Trim(rst1!AccountCode)
    MItem.SubItems(4) = Trim(rst1!remarks)
    rst1.MoveNext
Loop
rst1.close
If FormNo > 0 Then
 Newform.ListView1.SortKey = Newform.ListView1.ColumnHeaders(3).Index
 WhatColumnclick = 3
End If
Newform.ListView1.Sorted = True

If FormNo > 1 Then
    Newform.Label1.caption = Level1Code & Level2Code & level3Code & level4Code & level5Code
End If

Newform.Label3 = Newform.ListView1.ListItems.Count & " Total item(s) found! íÌÏ ãÌãæÚ ÇáÇÕäÇÝ "
Newform.Show 1

End Sub

Private Sub Command3_Click()

Unload Me
End Sub

Private Sub Command4_Click()

End Sub

Private Sub delete_Click()
 If UCase(UserRole) <> UCase("admin") Then
  If UCase(Left(UserRole, 5)) <> UCase(Left("Accoun", 5)) Then

   mess = MsgBox("You are not authorized to delete the item", vbExclamation + vbOKOnly, "Message")
   Exit Sub
 End If
 End If


Dim rsDelrec As New ADODB.Recordset
Dim CanDelete As Boolean
Dim Acctname As String
On Error Resume Next
Acctname = Trim(Me.ListView1.SelectedItem.Text)
cindex = Me.ListView1.SelectedItem.Index
AcctCode = Trim(Me.ListView1.SelectedItem.SubItems(3))
constring = "dsN=fINANCE;UID=SA;PWD=;"
mess = MsgBox("Are you sure you want delete " & Acctname & "?", vbOKCancel + vbQuestion + vbDefaultButton2, "Please confirm")
  If mess = vbOK Then
    rsDelrec.Open "Select * from FinanceMaster where AccountCode = " & " '" & AcctCode & "'", constring, adOpenKeyset, adLockOptimistic, adCmdText
    If rsDelrec.EOF = False Then
      If rsDelrec!TotalTrans <> 0 Then
        mess = MsgBox("Account has as a transactions,Deletion is not allowed", vbExclamation + vbOKOnly, "Message")
        Exit Sub
       Else
      'End If
     'Else
      rsDelrec.close
      Dim rsJOurnals As New ADODB.Recordset
      Dim HasTrans As Boolean
      Dim rsTRans As New ADODB.Recordset
      rsJOurnals.Open "select * from JOurnalCode order by code", constring, adOpenKeyset, adLockPessimistic, adCmdText
      
      Do Until rsJOurnals.EOF = True
        If Trim(rsJOurnals!TableName) <> "" Then
         xtable = Trim(rsJOurnals!TableName)
         If Trim(rsJOurnals!Code) <> "PTC" And Trim(rsJOurnals!Code) <> "PYB" Then
            rsTRans.Open "Select * from " & xtable & " where accountnumber =" & "'" & AcctCode & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
           ElseIf rsJOurnals!Code = "PTC" Then
            rsTRans.Open "Select * from " & xtable & " where accoutno =" & "'" & AcctCode & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
           ElseIf rsJOurnals!Code = "PYB" Then
            rsTRans.Open "Select * from " & xtable & " where accno =" & "'" & AcctCode & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
         End If
         If rsTRans.EOF = True Then
            rsTRans.close
            rsJOurnals.MoveNext
           Else
            HasTrans = True
            mess = MsgBox("Account has as an unposted transactions,Deletion is not allowed otherwise delete its transactions", vbExclamation + vbOKOnly, "Message")
            Exit Sub
          End If
         End If
          rsJOurnals.MoveNext
      Loop
     End If
    End If
    On Error Resume Next
    rsDelrec.close
    rsDelrec.Open "delete FinanceMaster where AccountCode = " & " '" & AcctCode & "'", constring, adOpenKeyset, adLockOptimistic, adCmdText
    On Error GoTo 0
    
    If FormNo - 1 = 1 Then
       'Let's Check first on the next level if this item on the current level has no sub cat, If positive we don't allow to delete unless subcat will be deleted.
       rsDelrec.Open "Select * from Level2 where Left(AccountCode,2)=" & "'" & Left(AcctCode, 2) & "'", constring, adOpenKeyset, adLockOptimistic, adCmdText
       Call ConfirmToDel(rsDelrec, CanDelete, Acctname)
       If CanDelete = True Then
         rsDelrec.Open "Delete Level1 where AccountCode=" & "'" & AcctCode & "'", constring, adOpenKeyset, adLockOptimistic, adCmdText
       End If
    
    ElseIf FormNo - 1 = 2 Then
        'Let's Check first on the next level if this item on the current level has no sub cat, If positive we don't allow to delete unless subcat will be deleted.
       rsDelrec.Open "Select * from Level3 where Left(AccountCode,3)=" & "'" & Left(AcctCode, 3) & "'", constring, adOpenKeyset, adLockOptimistic, adCmdText
       Call ConfirmToDel(rsDelrec, CanDelete, Acctname)
       If CanDelete = True Then
         rsDelrec.Open "Delete Level2 where AccountCode=" & "'" & AcctCode & "'", constring, adOpenKeyset, adLockOptimistic, adCmdText
       End If
    
    ElseIf FormNo - 1 = 3 Then
       'Let's Check first on the next level if this item on the current level has no sub cat, If positive we don't allow to delete unless subcat will be deleted.
       rsDelrec.Open "Select * from Level4 where Left(AccountCode,5)=" & "'" & Left(AcctCode, 5) & "'", constring, adOpenKeyset, adLockOptimistic, adCmdText
       Call ConfirmToDel(rsDelrec, CanDelete, Acctname)
       If CanDelete = True Then
         rsDelrec.Open "Delete Level3 where AccountCode=" & "'" & AcctCode & "'", constring, adOpenKeyset, adLockOptimistic, adCmdText
       End If
    
    ElseIf FormNo - 1 = 4 Then
       'Let's Check first on the next level if this item on the current level has no sub cat, If positive we don't allow to delete unless subcat will be deleted.
       rsDelrec.Open "Select * from Level5 where Left(AccountCode,7)=" & "'" & Left(AcctCode, 7) & "'", constring, adOpenKeyset, adLockOptimistic, adCmdText
      Call ConfirmToDel(rsDelrec, CanDelete, Acctname)
       If CanDelete = True Then
         rsDelrec.Open "Delete Level4 where AccountCode=" & "'" & AcctCode & "'", constring, adOpenKeyset, adLockOptimistic, adCmdText
       End If
    
    ElseIf FormNo - 1 = 5 Then
      'Let's Check first on the next level if this item on the current level has no sub cat, If positive we don't allow to delete unless subcat will be deleted.
       rsDelrec.Open "Select * from Level6 where Left(AccountCode,9)=" & "'" & Left(AcctCode, 9) & "'", constring, adOpenKeyset, adLockOptimistic, adCmdText
       Call ConfirmToDel(rsDelrec, CanDelete, Acctname)
       If CanDelete = True Then
         rsDelrec.Open "Delete Level5 where AccountCode=" & "'" & AcctCode & "'", constring, adOpenKeyset, adLockOptimistic, adCmdText
       End If
       
    ElseIf FormNo - 1 = 6 Then
       rsDelrec.Open "Delete Level6 where AccountCode=" & "'" & AcctCode & "'", constring, adOpenKeyset, adLockOptimistic, adCmdText
    End If
    If CanDelete = True Or FormNo - 1 = 6 Then
       CanDelete = False
        Me.ListView1.ListItems.Remove cindex
     mess = MsgBox("Successfully deleted ", vbExclamation + vbOKOnly, "Message")
    End If
End If

End Sub
Sub ConfirmToDel(rsDelrec As Recordset, CanDelete As Boolean, Acctname As String)
If rsDelrec.EOF = False Then
   If FormNo - 1 = 5 Then
     msg = MsgBox(Acctname & " has a Main Account(s) on Level " & Str(FormNo) & ", Deletion is not allowed ", vbInformation + vbOKOnly, "Message")
    Else
    msg = MsgBox(Acctname & " has a sub-category on Level " & Str(FormNo) & ", Deletion is not allowed ", vbInformation + vbOKOnly, "Message")
   End If
    Exit Sub
  Else
  CanDelete = True
  rsDelrec.close
End If
End Sub
Private Sub Form_Activate()
If CancelAll = True Then
   Unload Me
End If
On Error Resume Next
Me.ListView1.SetFocus
cTotalItems = Me.ListView1.ListItems.Count
End Sub

Private Sub Form_Load()
If xCountry = "" Then
    Dim rst As New ADODB.Recordset
    constring = "dsN=fINANCE;UID=SA;PWD=;"
    rst.Open "Select * from country order by country", constring, adOpenKeyset, adLockPessimistic, adCmdText
    Set xcol = Me.ListView1.ColumnHeaders.Add(, "a", "Country ÈáÏ", 4000)
    Set xcol = Me.ListView1.ColumnHeaders.Add(, "b", "Code ßæÏ", 1000)
    Do Until rst.EOF
        Set MItem = Me.ListView1.ListItems.Add(, , rst!Country, , 1)
        MItem.SubItems(1) = rst!Code
        rst.MoveNext
    Loop
End If

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = 2 Then
    PopupMenu Main
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
 If CancelAll = False Then
    msgs = MsgBox("Are you sure you want to Close?åá ÇäÊ ãÊÇßÏ ãä ÇáÛáÞ  ", vbQuestion + vbYesNo, "Please confirm ãä ÝÖáß ÇáÊÇßí")
 Else
  Unload Me
 Exit Sub
End If
    On Error Resume Next
    If msgs = vbYes Then
        CancelAll = True
        Unload Newform
        Else
        Cancel = -1
    
End If
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = 2 Then
    PopupMenu Main
End If
End Sub

Private Sub ListView1_AfterLabelEdit(Cancel As Integer, Newstring As String)
xNewstring = Newstring
constring = "dsN=fINANCE;UID=SA;PWD=;"
Dim rstEdit As New ADODB.Recordset

If oldstring <> xNewstring Then
  mess = MsgBox("Do you want to save changes?åá ÊÑíÏ ÍÝÙ ÇáÊÛíÑÇÊ ", vbOKCancel + vbQuestion, "Please confirm ãä ÝÖáß ÇáÊÇßíÏ")
  If mess = vbOK Then
    On Error Resume Next
    AcctCode = Trim(Newform.ListView1.SelectedItem.SubItems(3))
    If AcctCode = "" Then
       AcctCode = Trim(Me.ListView1.SelectedItem.SubItems(3))
    End If
    rstEdit.Open "Update FinanceMaster Set accountnameEng=" & "'" & Newstring & "'" & " Where AccountCode = " & " '" & AcctCode & "'", constring, adOpenKeyset, adLockOptimistic, adCmdText
    On Error GoTo EditCountry
    If Newform.caption = "Country" Then
       rstEdit.Open "Update Country Set accountnameEng=" & "'" & Newstring & "'" & " Where AccountNameEng = " & " '" & oldstring & "'", constring, adOpenKeyset, adLockOptimistic, adCmdText
      
      ElseIf Newform.caption = "Financial Position Category" Then
         rstEdit.Open "Update TopLevel Set accountnameEng=" & "'" & Newstring & "'" & " Where AccountCode = " & " '" & AcctCode & "'", constring, adOpenKeyset, adLockOptimistic, adCmdText

      ElseIf Left(Newform.caption, 6) = "Level1" Then
         rstEdit.Open "Update Level1 Set accountnameEng=" & "'" & Newstring & "'" & " Where AccountCode = " & " '" & AcctCode & "'", constring, adOpenKeyset, adLockOptimistic, adCmdText
      ElseIf Left(Newform.caption, 6) = "Level2" Then
         rstEdit.Open "Update level2 Set accountnameEng=" & "'" & Newstring & "'" & " Where AccountCode = " & " '" & AcctCode & "'", constring, adOpenKeyset, adLockOptimistic, adCmdText
      ElseIf Left(Newform.caption, 6) = "Level3" Then
          rstEdit.Open "Update level3 Set accountnameEng=" & "'" & Newstring & "'" & " Where AccountCode = " & " '" & AcctCode & "'", constring, adOpenKeyset, adLockOptimistic, adCmdText
      ElseIf Left(Newform.caption, 6) = "Level4" Then
         rstEdit.Open "Update level4 Set accountnameEng=" & "'" & Newstring & "'" & " Where AccountCode = " & " '" & AcctCode & "'", constring, adOpenKeyset, adLockOptimistic, adCmdText
      ElseIf Left(Newform.caption, 6) = "Level5" Then
          rstEdit.Open "Update level5 Set accountnameEng=" & "'" & Newstring & "'" & " Where AccountCode = " & " '" & AcctCode & "'", constring, adOpenKeyset, adLockOptimistic, adCmdText
      ElseIf Left(Newform.caption, 6) = "Level6" Then
         rstEdit.Open "Update level6 Set accountnameEng=" & "'" & Newstring & "'" & " Where AccountCode = " & " '" & AcctCode & "'", constring, adOpenKeyset, adLockOptimistic, adCmdText
    End If
    Exit Sub
   Else
    Cancel = -1
   End If
End If
EditCountry:
X = Err.Description
If Me.caption = "Country" Then
       rstEdit.Open "Update Country Set country=" & "'" & Newstring & "'" & " Where country = " & " '" & oldstring & "'", constring, adOpenKeyset, adLockOptimistic, adCmdText
      ElseIf Me.caption = "Financial Position Category" Then
         rstEdit.Open "Update TopLevel Set accountnameEng=" & "'" & Newstring & "'" & " Where AccountCode = " & " '" & AcctCode & "'", constring, adOpenKeyset, adLockOptimistic, adCmdText
      ElseIf Me.caption = "Level1" Then
         rstEdit.Open "Update Level1 Set accountnameEng=" & "'" & Newstring & "'" & " Where AccountCode = " & " '" & AcctCode & "'", constring, adOpenKeyset, adLockOptimistic, adCmdText
      ElseIf Me.caption = "Level2" Then
         rstEdit.Open "Update level2 Set accountnameEng=" & "'" & Newstring & "'" & " Where AccountCode = " & " '" & AcctCode & "'", constring, adOpenKeyset, adLockOptimistic, adCmdText
      ElseIf Me.caption = "Level3" Then
          rstEdit.Open "Update level3 Set accountnameEng=" & "'" & Newstring & "'" & " Where AccountCode = " & " '" & AcctCode & "'", constring, adOpenKeyset, adLockOptimistic, adCmdText
      ElseIf Me.caption = "Level4" Then
         rstEdit.Open "Update level4 Set accountnameEng=" & "'" & Newstring & "'" & " Where AccountCode = " & " '" & AcctCode & "'", constring, adOpenKeyset, adLockOptimistic, adCmdText
      ElseIf Me.caption = "Level5" Then
          rstEdit.Open "Update level5 Set accountnameEng=" & "'" & Newstring & "'" & " Where AccountCode = " & " '" & AcctCode & "'", constring, adOpenKeyset, adLockOptimistic, adCmdText
      ElseIf Me.caption = "Level6" Then
         rstEdit.Open "Update level6 Set accountnameEng=" & "'" & Newstring & "'" & " Where AccountCode = " & " '" & AcctCode & "'", constring, adOpenKeyset, adLockOptimistic, adCmdText
 Exit Sub
End If
End Sub

Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)
On Error GoTo EditCountry
oldstring = Newform.ListView1.SelectedItem.Text
LevelCode = Trim(Right(Me.ListView1.SelectedItem.SubItems(2), 1))

EditCountry:
 If Me.caption = "Country" Then
    oldstring = Me.ListView1.SelectedItem.Text
End If
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo CountryIndex
Newform.ListView1.SortKey = ColumnHeader.Index - 1
If ColumnHeader.Index = 1 Or ColumnHeader.Index = 2 Or ColumnHeader.Index = 5 Then
 Me.RenameArabName.caption = "Rename " & Trim(Me.ListView1.ColumnHeaders.Item(ColumnHeader.Index).Text)
 Me.RenameArabName.Enabled = True
Else
 Me.RenameArabName.Enabled = False
 Me.RenameArabName.caption = "Rename ÇÚÇÏÉ ÇáÇÓã"
End If
Newform.ListView1.Sorted = True
CountryIndex:
If Err.Number > 0 Then
    Me.ListView1.SortKey = ColumnHeader.Index - 1
    Me.ListView1.Sorted = True
End If
WhatColumnclick = ColumnHeader.Index - 1

End Sub

Private Sub ListView1_DblClick()
Call Command2_Click
End Sub

Private Sub ListView1_GotFocus()

On Error Resume Next
If FormNo > 1 Then
    Me.Label1.caption = Me.ListView1.SelectedItem.SubItems(1)
End If
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error Resume Next
If FormNo > 1 Then
    Me.Label1.caption = Me.ListView1.SelectedItem.SubItems(1)
End If
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Call Command1_Click
End If
End Sub

Private Sub ListView1_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = 113 Then
  If FormNo > 3 Then
    Call AddItem_Click
  End If
End If
If KeyCode = 114 Then
    Call Rename_Click
End If

If KeyCode = 46 Then
   If FormNo > 3 Then
    Call delete_Click
   End If
End If

End Sub


Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)


If Me.ListView1.ListItems.Count = 0 Then
  Me.RenameArabName.Enabled = False
  Me.delete.Enabled = False
  Me.xRefresh.Enabled = False
  Me.xPrint.Enabled = False
 Else
  Me.RenameArabName.Enabled = True
  Me.delete.Enabled = True
  Me.xRefresh.Enabled = True
  Me.xPrint.Enabled = True
 End If
 
If FormNo < 2 Then
    Me.xPrint.Enabled = False
   Else
   Me.xPrint.Enabled = True
End If
 
If FormNo < 4 Then
  Me.AddItem.Enabled = False
  Me.delete.Enabled = False
 Else
  Me.AddItem.Enabled = True
  Me.delete.Enabled = True
End If

If Button = 2 Then
    PopupMenu Main
End If

End Sub

Private Sub Rename_Click()
'Me.ListView1.StartLabelEdit
End Sub

Private Sub RenameArabName_Click()
 If UCase(UserRole) <> UCase("admin") Then
  If UCase(Left(UserRole, 5)) <> UCase(Left("Accoun", 5)) Then
   mess = MsgBox("You are not authorized to rename the item", vbExclamation + vbOKOnly, "Message")
   Exit Sub
 End If
End If
RenArabName.Combo1 = Trim(Me.ListView1.SelectedItem.SubItems(3))
RenArabName.Combo2 = Trim(Me.ListView1.SelectedItem)
 
If WhatColumnclick = 1 Then
  RenArabName.Combo3.RightToLeft = True
  RenArabName.Combo3 = Trim(Me.ListView1.SelectedItem.SubItems(1))
  RenArabName.Label3 = "AccountName Arab ÇÓã ÇáÍÓÇÈ ÈÇáÚÑÈí"
  
ElseIf WhatColumnclick = 4 Then
  RenArabName.Combo3.RightToLeft = False
  RenArabName.Combo3 = Trim(Me.ListView1.SelectedItem.SubItems(4))
  RenArabName.Label3 = "Remarks ãáÇÍÙÇÊ "
  Me.RenameArabName.caption = Trim(Me.ListView1.ColumnHeaders.Item(WhatColumnclick).Text)
ElseIf WhatColumnclick = 0 Then
Me.ListView1.StartLabelEdit
Exit Sub
ElseIf WhatColumnclick = 2 Or WhatColumnclick = 3 Then
    mess = MsgBox(Trim(Me.ListView1.ColumnHeaders.Item(WhatColumnclick + 1).Text) & " is not editable ")
    Exit Sub
End If

RenArabName.Show 1
End Sub

Private Sub xPrint_Click()
If FormNo = 0 Then
    Exit Sub
End If
    
Dim acctNo As String
Dim caption As String
Dim rstTempLevel As New ADODB.Recordset
Dim rstLevel As New ADODB.Recordset
prevItem = Me.ListView1.SelectedItem
caption = Trim(Me.ListView1.SelectedItem)
acctNo = Me.ListView1.SelectedItem.SubItems(3)

constring = "dsN=fINANCE;UID=SA;PWD=;"
If FormNo - 1 = 0 Then
    xtable = "Select * from financemaster where Left(accountcode,1)=" & "'" & Left(acctNo, 1) & "'"
ElseIf FormNo - 1 = 1 Then
   xtable = "Select * from financemaster where Left(accountcode,2)=" & "'" & Left(acctNo, 2) & "'"
ElseIf FormNo - 1 = 2 Then
   xtable = "Select * from financemaster where Left(accountcode,3)=" & "'" & Left(acctNo, 3) & "'"
ElseIf FormNo - 1 = 3 Then
   xtable = "Select * from financemaster where Left(accountcode,5)=" & "'" & Left(acctNo, 5) & "'"
ElseIf FormNo - 1 = 4 Then
   xtable = "Select * from financemaster where Left(accountcode,7)=" & "'" & Left(acctNo, 7) & "'"
ElseIf FormNo - 1 = 5 Then
   xtable = "Select * from financemaster where Left(accountcode,9)=" & "'" & Left(acctNo, 9) & "'"
ElseIf FormNo - 1 = 6 Then
   mess = MsgBox("This is the last heirarchy of the chart", vbExclamation + vbOKOnly, "Message")
   Exit Sub
End If
rstTempLevel.Open "delete TempLevel", constring, adOpenKeyset, adLockPessimistic, adCmdText
rstTempLevel.Open "TempLevel", constring, adOpenKeyset, adLockPessimistic, adCmdTable
rstLevel.Open xtable, constring, adOpenKeyset, adLockPessimistic, adCmdText
Do Until rstLevel.EOF = True
    Me.Text1.Visible = True
    i = 0
    With rstTempLevel
       .addnew
       !AccountCode = rstLevel!AccountCode
       !accountnameeng = rstLevel!accountnameeng
       !accountnamearab = rstLevel!accountnamearab
       !remarks = IIf(rstLevel!Active = 1, "Active", "Inactive")
       'If FormNo - 1 = 1 Then
        Dim catName As String
        Dim Prevcap As String
        Dim acctNum As String
        acctNum = Trim(!AccountCode)
        Prevcap = Trim(Me.caption)
        Call DisplayCats(Prevcap, acctNum, catName)
        DrCat = catName
        !MotherCat = DrCat
'               !MotherCat = Trim(rstLevel!TopLevelName)
'            ElseIf FormNo - 1 = 2 Then
'               !MotherCat = Trim(rstLevel!Level1Name) '& "<-" & Trim(rstLevel!TopLevelName)
'            ElseIf FormNo - 1 = 3 Then
'               !MotherCat = Trim(rstLevel!Level1Name) & Chr(187) & Trim(rstLevel!Level2Name)
'            ElseIf FormNo - 1 = 4 Then
'               !MotherCat = Trim(rstLevel!Level1Name) & Chr(187) & Trim(rstLevel!Level2Name) & Chr(187) & Trim(rstLevel!level3Name) '& "<-" & Trim(rstLevel!TopLevelName)
'            ElseIf FormNo - 1 = 5 Then
'               !MotherCat = Trim(rstLevel!Level1Name) & Chr(187) & Trim(rstLevel!Level2Name) & Chr(187) & Trim(rstLevel!level3Name) & Chr(187) & Trim(rstLevel!level4Name) ' & "<-" & Trim(rstLevel!TopLevelName)
'            ElseIf FormNo - 1 = 6 Then
'               !MotherCat = Trim(rstLevel!Level1Name) & Chr(187) & Trim(rstLevel!Level2Name) & Chr(187) & Trim(rstLevel!level3Name) & Chr(187) & Trim(rstLevel!level4Name) & Chr(187) & Trim(rstLevel!level5Name) ' & "<-" & Trim(rstLevel!TopLevelName)
        'End If
       
       .Update
    End With
    DoEvents
   rstLevel.MoveNext
Loop
Me.Text1.Visible = False
cWhatLevelName TempLevelPrn.Sections(1).Controls("Label5"), caption
On Error Resume Next
FinanceDE.rsTempLevel.close
TempLevelPrn.Show 1
End Sub
Private Sub cWhatLevelName(lblX As RptLabel, caption As String)
   With lblX
      .CanGrow = True
      .caption = "List of Accounts Under " & caption
   End With
End Sub

Private Sub xREfresh_Click()
Dim rst1 As New ADODB.Recordset
constring = "dsN=fINANCE;UID=SA;PWD=;"
FormNo = FormNo '+ 1
If FormNo = 1 Then
    WhatLevel = "TopLevel"
    rst1.Open "Select * from TopLevel where country =" & "'" & xCountry & "'" & "order by AccountNameEng ", constring, adOpenKeyset, adLockPessimistic, adCmdText

  ElseIf FormNo = 2 Then
    WhatLevel = "Level1"
    rst1.Open "Select * from Level1 where country =" & "'" & xCountry & "'" _
    & " and TOpLevelCode = " & "'" & TopLevelCode & "'" & "order by AccountNameEng ", constring, adOpenKeyset, adLockPessimistic, adCmdText

  ElseIf FormNo = 3 Then
    WhatLevel = "Level2"
    rst1.Open "Select * from Level2 where country =" & "'" & xCountry & "'" _
    & " and TOpLevelCode = " & " '" & TopLevelCode & "'" & " and Level1Code=" & "'" _
    & Level1Code & " '" & " order by AccountNameEng ", constring, adOpenKeyset, adLockPessimistic, adCmdText
 
    
  ElseIf FormNo = 4 Then
    WhatLevel = "Level3"
    rst1.Open "Select * from Level3 where country =" & "'" & xCountry & "'" _
    & " and TOpLevelCode = " & " '" & TopLevelCode & "'" _
    & " and Level1Code=" & "'" & Level1Code & " '" _
    & " and Level2Code = " & "'" & Level2Code & "'" _
    & " order by AccountNameEng ", constring, adOpenKeyset, adLockPessimistic, adCmdText
    
  ElseIf FormNo = 5 Then
    WhatLevel = "Level4"
    rst1.Open "Select * from Level4 where country =" & "'" & xCountry & "'" _
    & " and TOpLevelCode = " & " '" & TopLevelCode & "'" _
    & " and Level1Code=" & "'" & Level1Code & " '" _
    & " and Level2Code = " & "'" & Level2Code & "'" _
    & " and Level3Code = " & "'" & level3Code & "'" _
    & " order by AccountNameEng ", constring, adOpenKeyset, adLockPessimistic, adCmdText
    
  ElseIf FormNo = 6 Then
    WhatLevel = "Level5"
    rst1.Open "Select * from Level5 where country =" & "'" & xCountry & "'" _
    & " and TOpLevelCode = " & " '" & TopLevelCode & "'" _
    & " and Level1Code=" & "'" & Level1Code & " '" _
    & " and Level2Code = " & "'" & Level2Code & "'" _
    & " and Level3Code = " & "'" & level3Code & "'" _
    & " and Level4Code = " & "'" & level4Code & "'" _
    & " order by AccountNameEng ", constring, adOpenKeyset, adLockPessimistic, adCmdText
 ElseIf FormNo = 7 Then
    WhatLevel = "Level6"
    rst1.Open "Select * from Level6 where country =" & "'" & xCountry & "'" _
    & " and TOpLevelCode = " & " '" & TopLevelCode & "'" _
    & " and Level1Code=" & "'" & Level1Code & " '" _
    & " and Level2Code = " & "'" & Level2Code & "'" _
    & " and Level3Code = " & "'" & level3Code & "'" _
    & " and Level4Code = " & "'" & level4Code & "'" _
     & " and Level5Code = " & "'" & level5Code & "'" _
    & " order by AccountNameEng ", constring, adOpenKeyset, adLockPessimistic, adCmdText
  End If


Newform.ListView1.ListItems.clear
On Error GoTo Nelson
Do Until rst1.EOF
    Set MItem = Newform.ListView1.ListItems.Add(, , Trim(rst1!accountnameeng), , IIf(Left(rst1!remarks, 4) = "Main", 4, 1))
    MItem.SubItems(1) = IIf(Trim(rst1!accountnamearab) = "", " ", Trim(rst1!accountnamearab))
    MItem.SubItems(2) = Trim(rst1!Code)
    MItem.SubItems(3) = Trim(rst1!AccountCode)
    MItem.SubItems(4) = Trim(rst1!remarks)
    
    rst1.MoveNext
Loop
rst1.close
Nelson:
c = Err.Number
If c = 380 Then
    cErr = c
     
End If

End Sub

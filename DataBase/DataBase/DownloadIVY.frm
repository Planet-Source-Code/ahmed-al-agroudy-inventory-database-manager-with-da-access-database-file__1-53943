VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form DownLoadIVY 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Please confirm"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Height          =   350
      Left            =   120
      ScaleHeight     =   285
      ScaleWidth      =   4395
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   4455
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   330
         Left            =   -15
         TabIndex        =   2
         Top             =   -15
         Width           =   4440
         _ExtentX        =   7832
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel ÇáÛÇÁ"
      Height          =   350
      Left            =   2640
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Continue ÇÓÊãÑÇÑ"
      Height          =   350
      Left            =   1080
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
   Begin MSComCtl2.Animation Animation1 
      Height          =   375
      Left            =   4560
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      _Version        =   393216
      AutoPlay        =   -1  'True
      BackColor       =   16777215
      FullWidth       =   33
      FullHeight      =   25
   End
   Begin VB.Label Label4 
      Caption         =   "Found: 0"
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
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Downloading Inventory Items...ÊÍãíá ÇÕäÇÝ ÇáãÎÒæä "
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "Do you want to continue for Downloading item taken?"
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   320
      Width           =   3855
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "DownloadIVY.frx":0000
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "DownLoadIVY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsIVY As New ADODB.Recordset
Dim rsIvy2 As New ADODB.Recordset
Dim trandate As Date
Dim MItem As ListItem
Dim i As Long
Dim cTotrec As Long
Dim ErrorConek As Boolean
Dim Taken As Long
Dim cVal As Long
Sub DownLoadIvy()
'Dim rstJOurcode As New ADODB.Recordset
'rstJOurcode.Open "Select * from JOurnalCode order by Code", constring, adOpenKeyset, adLockPessimistic, adCmdText
'rstJOurcode.Move 4, 1
'xCode = rstJOurcode!JOurnalName
'Dim trandate As Date
'trandate = rstJOurcode!lastpostingdate

'open Warehousecode
Dim rsWhCode As New ADODB.Recordset
rsWhCode.Open "WareHouseCode", constring, adOpenKeyset, adLockPessimistic, adCmdTable

Static CancelDownLoading As Boolean
Me.caption = "Please wait...ÇáÑÌÇÁ ÇáÇäÊÙÇÑ"
Me.Command1.Left = 1800
Me.Picture1.Visible = True
Me.Label1.Visible = True
Me.Label4.Visible = True
Me.Command2.Visible = False
Me.Label2.Visible = False
Me.Image1.Visible = False
If CancelDownLoading Then
   CancelDownLoading = False
  Else
    Me.Command1.caption = "Stop ÞÝ"
    CancelDownLoading = True
    
    Do Until rsIvy2.EOF = True
        i = i + 1
        If i = cTotrec Then
            cVal = cVal + 1
            i = 0
            If cVal <= 100 Then
            Me.ProgressBar1.Value = cVal
            Me.Label3.caption = cVal & "%"
            End If
            
        End If
     'If rsIvy2!trDate > trandate Then
        If Trim(UCase(rsIvy2!remarks1)) = "TAKEN" Then
         If rsIvy2!waveCost <> 0 Or rsIvy2!Cost <> 0 Then
           Taken = Taken + 1
           Me.Label4.caption = "Found :" & LTrim(Taken)
           Set MItem = InvJournalEntry.ListView3.ListItems.Add(, , rsIvy2!trDate)
           MItem.SubItems(1) = rsIvy2!grno
           MItem.SubItems(2) = rsIvy2!InvCat
           MItem.SubItems(3) = rsIvy2!Voucher
           MItem.SubItems(4) = rsIvy2!TRnu
           If Trim(rsIvy2!trType) <> "A" Then
              MItem.SubItems(5) = rsIvy2!waveCost
            Else
              MItem.SubItems(5) = rsIvy2!Cost
           End If
           
           'look for the WHCode
           FrCC = Trim(rsIvy2!TranRele)
           rsWhCode.MoveFirst
           Do Until rsWhCode.EOF = True
                If FrCC = Trim(rsIvy2!TranRele) Then
                   WHCode = rsWhCode!WCCode
                   Exit Do
                  End If
                rsWhCode.MoveNext
           Loop
           
           MItem.SubItems(6) = WHCode & "                   " & WHCode
           MItem.SubItems(7) = rsIvy2!ProdUnit & "                 " & Trim(rsIvy2!Depart)
           MItem.SubItems(8) = rsIvy2!Purpose
           MItem.SubItems(9) = rsIvy2!Work_no
           If Trim(rsIvy2!trType) = "A" Then
              MItem.SubItems(10) = "RR"
            ElseIf Trim(rsIvy2!trType) = "B" Then
              MItem.SubItems(10) = "DO"
            ElseIf Trim(rsIvy2!trType) = "C" Then
              MItem.SubItems(10) = "RO"
            ElseIf Trim(rsIvy2!trType) = "D" Then
              MItem.SubItems(10) = "TO"
            End If
            With rsIVY
                .AddNew
                rsIVY!TRansDate = rsIvy2!trDate
                rsIVY!grno = rsIvy2!grno
               
               'look for the WHCode
                FrCC = Trim(rsIvy2!TranRele)
                rsWhCode.MoveFirst
                Do Until rsWhCode.EOF = True
                     If FrCC = Trim(rsIvy2!TranRele) Then
                        WHCode = rsWhCode!WCCode
                        Exit Do
                     End If
                     rsWhCode.MoveNext
                Loop
                rsIVY!fr_CostCenter = WHCode
                rsIVY!fr_dept = WHCode
                rsIVY!To_Costcenter = rsIvy2!ProdUnit
                rsIVY!to_dept = rsIvy2!Depart
                rsIVY!Purpose = rsIvy2!Purpose
                rsIVY!WOno = rsIvy2!Work_no
                rsIVY!InvCat = rsIvy2!InvCat
                rsIVY!Trantype = rsIvy2!trType
               .Update
             End With
           End If
          End If
         'End If
         rsIvy2.MoveNext
         DoEvents
         If CancelDownLoading = False Then
            Exit Do
             Exit Sub
         End If
    Loop
    If CancelDownLoading = False Then
        Me.Animation1.Visible = False
        Exit Sub
      Else
      Beep
       mess = MsgBox("Finished downloading! " & Str(Taken) & " item(s) found " & "  ÇäÊåÇÁ ÇáÊÍãíá", vbInformation + vbOKOnly, "Message")
       CancelDownLoading = False
       i = 0
       cVal = 0
       rsIvy2.Close
       rsIVY.Close
       Me.Command2.Visible = False
       Me.Command1.Visible = False
       Me.Height = 1400
       Unload Me
       Exit Sub
       
    End If
    Unload Me
End If
Command1.caption = "ContinueÇÓÊãÑÇÑ"
CancelDownLoading = False
Exit Sub
End Sub


Private Sub Command1_Click()
DownLoadIvy

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Activate()
If ErrorConek = True Then
    ErrorConek = False
    Unload Me
End If
End Sub

Private Sub Form_Load()
If i = 0 Then
    On Error GoTo MyMsg
    rsIVY.Open "Delete Inventory", constring, adOpenKeyset, adLockPessimistic, adCmdText
    rsIVY.Open "Inventory", constring, adOpenKeyset, adLockPessimistic, adCmdTable
    InvJournalEntry.caption = "Please wait, Connecting to external Database..."
    Dim xcls As New HabitatClass
    Dim xtable As String
    Dim conRsIV2 As New ADODB.Connection
    Dim sqltable As Boolean
    sqltable = True
    DoEvents
    InvJournalEntry.MousePointer = 99
    InvJournalEntry.MouseIcon = LoadPicture(APp.Path & "\" & "wait_m.cur")
    DoEvents
    xcls.GetTables rsIvy2, conRsIV2, "select * from warehist where remarks1='TAKEN'", "dsN=WAreHist;UID=SA;PWD=;", sqltable
    cTotrec = Int(rsIvy2.RecordCount / 100)
    Totrec = cTotrec
    i = 0
    cVal = 0
    InvJournalEntry.MousePointer = 0
End If
InvJournalEntry.caption = "Inventory Journal Entry   ÅÏÎÇá ÞíæÏ ÇáíæãíÉ "



MyMsg:
 c = Err.Number
 d = Err.Description
If c = 3705 Then
    
   
   Else
   X = Err.Description
   If c = 3705 Then
    ErrorConek = True
   MsgBox ("Maybe file is used by other user,try again later" & vbCrLf & _
               " ÑÈãÇ ÇáãáÝ íÓÊÎÏã ÈæÇÓØÉ ãÓÊÎÏã ÇÎÑ Íæá Ýí æÞÊ ÇÎÑ  ")
   'Unload Me
   End If
 End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Nelson
If rsInv2.EOF = False And cVal <> 0 Then
    mess = MsgBox("Are you sure you want to cancel downloading?" & vbCrLf & _
           " åá ÇäÊ ãÊÇßÏ ãä ÇáÛÇÁ ÇáÊÍãíá ", vbQuestion + vbYesNo, "Please confirm ãä ÝÖáß ÇáÊÇßíÏ")
    If mess = vbNo Then
        Cancel = -1
        'SendKeys "{Enter}"
      Else
         Unload Me
    End If
   Else
End If
Nelson:
Unload Me
End Sub

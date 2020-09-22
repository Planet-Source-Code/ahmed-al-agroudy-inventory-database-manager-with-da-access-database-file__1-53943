VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmSQL 
   Caption         =   "MySQL v.1.0.0 by Nelson Rosell  "
   ClientHeight    =   4665
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8880
   BeginProperty Font 
      Name            =   "Courier"
      Size            =   9.75
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SQL.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4665
   ScaleWidth      =   8880
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2280
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   1480
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "E&xecute"
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
      Height          =   375
      Left            =   50
      TabIndex        =   2
      Top             =   1480
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   1095
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   5175
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   1920
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   4683
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1025
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1025
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   0
      Top             =   4320
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=Ledger"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "Ledger"
      OtherAttributes =   ""
      UserName        =   "sa"
      Password        =   ""
      RecordSource    =   "select * from accounts"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Top             =   1560
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Enter SQL Statement"
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
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "FrmSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Sqlstring As String

Private Sub Combo1_Click()
Me.Text1.Text = Me.Combo1
End Sub

Private Sub Command1_Click()
Dim rst As New ADODB.Recordset
Dim con As New ADODB.Connection
On Error GoTo MyErr
If Sqlstring = Me.Text1.Text Then
 Exit Sub
End If
'conString = "dsn=LEdger;uid=sa;pwd=;"
rst.CursorLocation = adUseClient
con.Open constring
i = 0
For i = 1 To Len(Me.Text1.Text)
   If Mid(Me.Text1.Text, i, 1) = "&" Then Exit For
     xstring = xstring & Mid(Me.Text1.Text, i, 1)
     
Next
If UCase(Left(Trim(xstring), 6)) = UCase("Delete") Or UCase(Left(Trim(xstring), 6)) = UCase("Update") Or UCase(Left(Trim(xstring), 6)) = UCase("Insert") Then
  If UCase(Left(Trim(xstring), 6)) = UCase("Update") Then
    mess = MsgBox("Replace it?", vbQuestion + vbOKCancel + vbDefaultButton2, "Please confirm")
   ElseIf UCase(Left(Trim(xstring), 6)) = UCase("Delete") Then
    mess = MsgBox("Delete it?", vbQuestion + vbOKCancel + vbDefaultButton2, "Please confirm")
   ElseIf UCase(Left(Trim(xstring), 6)) = UCase("insert") Then
    mess = MsgBox("Insert it?", vbQuestion + vbOKCancel + vbDefaultButton2, "Please confirm")
  End If
  If mess = vbOK Then
         rst.Open xstring, con, adOpenDynamic, adLockOptimistic, adCmdText
  End If
 Else
 rst.Open xstring, con ', adOpenKeyset, adLockPessimistic, adCmdText
End If
On Error Resume Next
Set Me.Adodc1.Recordset = rst
Me.Adodc1.Refresh
Set Me.DataGrid1.DataSource = Me.Adodc1
Me.Label2 = rst.RecordCount & " record(s) affected."
Me.Combo1.Visible = True
rst.Close
Me.Combo1.AddItem xstring
Sqlstring = xstring


MyErr:
c = Err.Number
X = Err.Description
If c = 3704 Then
    Exit Sub
End If
If c <> 0 Then
    xmess = MsgBox(Err.Description, vbInformation + vbOKOnly, "Error on Statement")
    Me.Text1.SetFocus
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Activate()

Me.caption = "MySQL v.1.0.0 by Nelson Rosell  " & Chr(169) & " 2002 Copyright"
Me.Text1.SetFocus
If Me.Text1.Text <> "" Then
    Call Command1_Click
End If
If UCase(UserRole) <> UCase("Admin") Then
    Me.Text1.Enabled = False
    Me.Command1.Enabled = False
   Else
     Me.Text1.Enabled = True
    Me.Command1.Enabled = True
End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
Me.Text1.Width = Me.Width - 150
Me.Combo1.Width = Me.Width - 2400
Me.DataGrid1.Width = Me.Width - 150
Me.DataGrid1.Top = Me.Text1.Height + 800
Me.DataGrid1.Height = Me.Height - 2500
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Text1.Text = ""
Sqlstring = ""
End Sub

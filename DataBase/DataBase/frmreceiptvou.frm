VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmrecieptvou 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Receipt Voucher ÓäÏ ÇáÞÈÖ"
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11415
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   9
      Charset         =   177
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmreceiptvou.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   11415
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timaccountnumber 
      Interval        =   500
      Left            =   4560
      Top             =   6840
   End
   Begin VB.Timer Timer2 
      Left            =   3480
      Top             =   6000
   End
   Begin MSComctlLib.ListView listcurrency 
      Height          =   1455
      Left            =   6960
      TabIndex        =   5
      Top             =   5760
      Visible         =   0   'False
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   2566
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Currency"
         Object.Width           =   1676
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Rate"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Lastest Update Date & Time"
         Object.Width           =   4410
      EndProperty
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "A&DD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   340
      Left            =   120
      TabIndex        =   70
      Top             =   7800
      Width           =   855
   End
   Begin VB.TextBox txtallinvoice 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   68
      Text            =   " "
      Top             =   7440
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      DataField       =   "INVC_NO"
      DataSource      =   "data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   66
      Text            =   "Text2"
      Top             =   6360
      Width           =   3135
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   570
      Left            =   120
      Top             =   5640
      Width           =   3120
      _ExtentX        =   5503
      _ExtentY        =   1005
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Dsn=Visual FoxPro Tables;sourcedb=\\Server2_obour\invoice\data\inv;Exclusive=No;"
      OLEDBString     =   "Dsn=Visual FoxPro Tables;sourcedb=\\Server2_obour\invoice\data\inv;Exclusive=No;"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "marcusfl"
      Caption         =   "this marcusfl runtime only"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   615
      Left            =   120
      Top             =   6720
      Width           =   3120
      _ExtentX        =   5503
      _ExtentY        =   1085
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
      Connect         =   "DSN=anufoxpro"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "anufoxpro"
      OtherAttributes =   ""
      UserName        =   "sa"
      Password        =   ""
      RecordSource    =   "select * from sjmaster"
      Caption         =   "adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5535
      Left            =   0
      TabIndex        =   24
      Top             =   120
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   9763
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   5
      TabHeight       =   617
      ShowFocusRect   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arabic Transparent"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Data Entry  ÇÏÎÇá ÇáÈíÇäÇÊ "
      TabPicture(0)   =   "frmreceiptvou.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtreceiptnumber"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "framebutton"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "langopt"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "List of View ÞÇÆãÉ "
      TabPicture(1)   =   "frmreceiptvou.frx":045E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "ListView2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.CheckBox langopt 
         Caption         =   "ÇÎÊíÇÑ ÇáØÈÇÚÉ (English)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   310
         Left            =   -70200
         Picture         =   "frmreceiptvou.frx":047A
         TabIndex        =   72
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   2385
      End
      Begin VB.Frame framebutton 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   -64920
         TabIndex        =   69
         Top             =   2640
         Width           =   1095
         Begin VB.CommandButton cmdclose 
            Caption         =   "Cl&ose"
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
            Left            =   120
            TabIndex        =   22
            Top             =   2160
            Width           =   855
         End
         Begin VB.CommandButton cmdnewrecord 
            Caption         =   "&New"
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
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton cmdprint 
            Caption         =   "&Print"
            Enabled         =   0   'False
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
            Left            =   120
            TabIndex        =   21
            Top             =   1680
            Width           =   855
         End
         Begin VB.CommandButton cmdsave 
            Caption         =   "&Save"
            Enabled         =   0   'False
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
            Left            =   120
            TabIndex        =   19
            Top             =   720
            Width           =   855
         End
         Begin VB.CommandButton cmdcancel 
            Caption         =   "&Cancel"
            Enabled         =   0   'False
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
            Left            =   120
            TabIndex        =   20
            Top             =   1200
            Width           =   855
         End
      End
      Begin MSMask.MaskEdBox txtreceiptnumber 
         Height          =   315
         Left            =   -73680
         TabIndex        =   33
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "########"
         PromptChar      =   " "
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   4935
         Left            =   120
         TabIndex        =   34
         Top             =   480
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   8705
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Receipt#"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Date"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Payment To"
            Object.Width           =   8114
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Pay-Type"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Pay-Option"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Amount"
            Object.Width           =   2293
         EndProperty
      End
      Begin VB.Frame Frame1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5055
         Left            =   -74880
         TabIndex        =   25
         Top             =   360
         Width           =   11175
         Begin VB.ComboBox txtdetails 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1200
            Style           =   1  'Simple Combo
            TabIndex        =   71
            Top             =   1680
            Width           =   5655
         End
         Begin VB.CommandButton cmdinvoice 
            Caption         =   ".s."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3840
            TabIndex        =   67
            Top             =   1320
            Width           =   375
         End
         Begin VB.ComboBox comreceivedfrom 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1200
            TabIndex        =   1
            Top             =   600
            Width           =   1215
         End
         Begin VB.ComboBox comfromname 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   2400
            TabIndex        =   2
            Top             =   600
            Width           =   4455
         End
         Begin VB.TextBox txttempreceipt 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   8400
            TabIndex        =   0
            Top             =   240
            Width           =   1815
         End
         Begin VB.ComboBox comsalesman 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   8400
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   1320
            Width           =   1815
         End
         Begin VB.ComboBox ccname 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   2760
            Sorted          =   -1  'True
            TabIndex        =   15
            Top             =   3000
            Width           =   5055
         End
         Begin VB.ComboBox comcreditaccountnumber 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   14
            Top             =   3000
            Width           =   2535
         End
         Begin VB.ComboBox cname 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   2760
            Sorted          =   -1  'True
            TabIndex        =   13
            Top             =   2400
            Width           =   5055
         End
         Begin VB.ComboBox combdebitaccnum 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            ItemData        =   "frmreceiptvou.frx":1344
            Left            =   120
            List            =   "frmreceiptvou.frx":1346
            Sorted          =   -1  'True
            TabIndex        =   12
            Top             =   2400
            Width           =   2535
         End
         Begin VB.TextBox creditamount 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   7920
            TabIndex        =   16
            Top             =   3000
            Width           =   1815
         End
         Begin VB.ComboBox comsetmode 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1200
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   1320
            Width           =   2655
         End
         Begin VB.ComboBox cmbpaymenttype 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   960
            Width           =   2655
         End
         Begin VB.TextBox txtdebitamt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   8950
            TabIndex        =   3
            Top             =   600
            Width           =   1265
         End
         Begin MSComCtl2.DTPicker txtcheckdate 
            Height          =   345
            Left            =   8400
            TabIndex        =   7
            Top             =   960
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   16777215
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   63176705
            CurrentDate     =   37548
         End
         Begin MSMask.MaskEdBox txtchecknumber 
            Height          =   315
            Left            =   5520
            TabIndex        =   6
            Top             =   960
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
            MaxLength       =   14
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "0"
            PromptChar      =   " "
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1620
            Left            =   120
            TabIndex        =   17
            Top             =   3360
            Width           =   9615
            _ExtentX        =   16960
            _ExtentY        =   2858
            View            =   3
            LabelEdit       =   1
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            HotTracking     =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Account Number"
               Object.Width           =   2999
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Account Name"
               Object.Width           =   10054
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Text            =   "Amount"
               Object.Width           =   3175
            EndProperty
         End
         Begin MSMask.MaskEdBox txtinvoicenum 
            Height          =   315
            Left            =   5520
            TabIndex        =   9
            Top             =   1320
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
            MaxLength       =   14
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "0"
            PromptChar      =   " "
         End
         Begin MSComCtl2.DTPicker dtpinvoicedate 
            Height          =   315
            Left            =   8400
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   1680
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   63176705
            CurrentDate     =   37193
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "ÇáãÈáÛ"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   9360
            TabIndex        =   76
            Top             =   2760
            Width           =   345
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "ÇáãÈáÛ"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   10680
            TabIndex        =   75
            Top             =   600
            Width           =   345
         End
         Begin VB.Label lblcurrency 
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "S.R"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   177
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   8400
            MouseIcon       =   "frmreceiptvou.frx":1348
            MousePointer    =   99  'Custom
            TabIndex        =   23
            Top             =   600
            Width           =   570
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "ÇÓÊáãÊ ãä"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   6840
            RightToLeft     =   -1  'True
            TabIndex        =   65
            Top             =   600
            Width           =   690
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ÓäÏ ãÄÞÊ"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   10440
            RightToLeft     =   -1  'True
            TabIndex        =   64
            Top             =   285
            Width           =   615
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Temp Invoice #"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   6960
            TabIndex        =   63
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            Caption         =   "ÊÇÑíÎ ÇáÝÇÊæÑÉ"
            BeginProperty Font 
               Name            =   "Arabic Transparent"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   10320
            RightToLeft     =   -1  'True
            TabIndex        =   62
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label25 
            Caption         =   "In. Date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7680
            TabIndex        =   61
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "ãáÇÍÙÇÊ "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   6840
            RightToLeft     =   -1  'True
            TabIndex        =   60
            Top             =   1320
            Width           =   660
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "ãÞÇÈá"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   4440
            TabIndex        =   59
            Top             =   1320
            Width           =   315
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "ØÑíÞÉ ÇáÏÝÚ"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   4080
            TabIndex        =   58
            Top             =   960
            Width           =   720
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Sa. Man"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   7680
            TabIndex        =   57
            Top             =   1320
            Width           =   600
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "ÇáÈÇÆÚ"
            BeginProperty Font 
               Name            =   "Arabic Transparent"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   10785
            RightToLeft     =   -1  'True
            TabIndex        =   56
            Top             =   1320
            Width           =   270
         End
         Begin VB.Label lblref 
            AutoSize        =   -1  'True
            Caption         =   "Ref No."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4920
            TabIndex        =   55
            Top             =   1320
            Width           =   555
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Account Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2760
            TabIndex        =   54
            Top             =   2760
            Width           =   1455
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Credit Acct Number"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   2760
            Width           =   1695
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Amount"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   7920
            TabIndex        =   52
            Top             =   2760
            Width           =   540
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Debit Acct Number"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   51
            Top             =   2160
            Width           =   1350
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Account Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2760
            TabIndex        =   50
            Top             =   2160
            Width           =   1455
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Amount"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   7920
            TabIndex        =   49
            Top             =   2160
            Width           =   540
         End
         Begin VB.Label lbldebitamount 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   7920
            TabIndex        =   48
            Top             =   2400
            Width           =   1815
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ÑÞã ÍÓÇÈ ÇáãÏíä "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1440
            TabIndex        =   47
            Top             =   2160
            Width           =   1140
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ÑÞã ÍÓÇÈ ÇáÏÇÆä "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1560
            TabIndex        =   46
            Top             =   2760
            Width           =   1110
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            Caption         =   "ÇÓã ÇáÍÓÇÈ "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   6960
            TabIndex        =   45
            Top             =   2760
            Width           =   810
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            Caption         =   "ÇáãÈáÛ"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   9360
            TabIndex        =   44
            Top             =   2160
            Width           =   345
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "ÇÓã ÇáÍÓÇÈ "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   6960
            TabIndex        =   43
            Top             =   2160
            Width           =   810
         End
         Begin VB.Label Label26 
            Caption         =   "Label26"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2400
            TabIndex        =   42
            Top             =   3600
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Voucher No :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Record Date  :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4080
            TabIndex        =   40
            Top             =   240
            Width           =   1185
         End
         Begin VB.Label lbldate 
            AutoSize        =   -1  'True
            Caption         =   "01/01/2003"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   5400
            TabIndex        =   39
            Top             =   240
            Width           =   1200
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ÑÞã ÇáÓäÏ"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   3000
            TabIndex        =   38
            Top             =   285
            Width           =   555
         End
         Begin VB.Label Label36 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "ãáÇÍÙÇÊ "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   6840
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Top             =   1680
            Width           =   660
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "ÑÞã ÇáÔíß"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   6900
            RightToLeft     =   -1  'True
            TabIndex        =   36
            Top             =   960
            Width           =   615
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "ÊÇÑíÎ ÇáÔíß "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   10200
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Top             =   960
            Width           =   810
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Description"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   32
            Top             =   1680
            Width           =   795
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Received From"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   31
            Top             =   600
            Width           =   1065
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Date"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   7680
            TabIndex        =   30
            Top             =   960
            Width           =   345
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Amount"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   7680
            TabIndex        =   29
            Top             =   600
            Width           =   555
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Chq No."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4920
            TabIndex        =   28
            Top             =   960
            Width           =   585
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Payment For"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   1320
            Width           =   915
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Payment Mode"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   26
            Top             =   960
            Width           =   1065
         End
      End
   End
   Begin VB.Label Label51 
      Caption         =   "Label51"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5640
      TabIndex        =   74
      Top             =   6000
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label37 
      Caption         =   "Label37"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   73
      Top             =   6000
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu delete 
         Caption         =   "Delete Item"
      End
   End
End
Attribute VB_Name = "frmrecieptvou"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As ADODB.Connection
Dim cnmar As New ADODB.Connection
Dim recor As New ADODB.Recordset
Dim creditcount As Integer ' this is for add button
Dim creditamountcount As Currency ' thsi is for count the balance
Dim recnum As New ADODB.Recordset
Dim recacc As New ADODB.Recordset
Dim recemp As New ADODB.Recordset
Dim recmar As New ADODB.Recordset
Dim rectemp As New ADODB.Recordset
Dim recte As New ADODB.Recordset
Dim reclist As New ADODB.Recordset
Dim CON1 As New ADODB.Connection
Dim recset As New ADODB.Recordset
Dim myclass As New HabitatClass
Dim takejournal As String
Dim sqltable As Boolean
Dim xtable As String
Dim xdecimal As Integer
Dim recpay As New ADODB.Recordset ' this is for paymode
Dim recmod As New ADODB.Recordset
Dim recteinvoice12 As New ADODB.Recordset
Dim recchecktemp As New ADODB.Recordset ' to check the temporary receipt number
Public fromwho As String
Public wrongcount As Integer
Public bringinvoice As String 'this takes the
Public stopstop1 As Integer
Dim cc As String
Dim recmarxx As New ADODB.Recordset
Dim conmar As New ADODB.Connection
Dim xClass As New HabitatClass
Dim slqtable As Boolean
Public checknumberanddate As String
Dim constring As String
Dim reccurrency As New ADODB.Recordset


Private Sub ccname_Click()
Dim recfindacc As New ADODB.Recordset
namenamenum = InStr(1, Trim(ccname.Text), "\", vbTextCompare)
If namenamenum > 0 Then
namename = Mid(Trim(ccname.Text), 1, (namenamenum - 1))
End If
recfindacc.Open "Select * from financemaster where active <> '0' and accountnameeng = " & "'" & namename & "'", CON1, adOpenKeyset, adLockOptimistic
If recfindacc.BOF = False Then
    comcreditaccountnumber.Text = recfindacc!AccountCode
End If
 recfindacc.close
 
'this is for the mother name
getmothername nonumber, namename, cc
Me.caption = "Receipt Voucher  " & cc
'end mother name

End Sub

Private Sub ccname_GotFocus()
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(ccname.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)

End Sub

Private Sub ccname_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyPageDown
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(ccname.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Select

End Sub

Private Sub ccname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    creditamount.SetFocus
End If
End Sub

Private Sub ccname_LostFocus()
Dim recfindacc As New ADODB.Recordset
namenamenum = InStr(1, Trim(ccname.Text), "\", vbTextCompare)
If namenamenum > 0 Then
    namename = Mid(Trim(ccname.Text), 1, (namenamenum - 1))
Else
namename = Trim(ccname.Text)
End If
recfindacc.Open "Select * from financemaster where active <> '0' and accountnameeng = " & "'" & namename & "'", CON1, adOpenKeyset, adLockOptimistic
If recfindacc.BOF = True Then
    recfindacc.close
    MsgBox "Please choose the Correct Account Number" & vbCrLf & " ãä ÝÖááß ÇÎÊÑ ÑÞã ÇáÍÓÇÈ ÇáÕÍíÍ", vbInformation, "Invalid Account Number"
    ccname.SetFocus
    Exit Sub
Else
    comcreditaccountnumber.Text = Trim(recfindacc!AccountCode)
End If

End Sub

Private Sub Frame2_DragDrop(Source As Control, X As Single, y As Single)

End Sub

Private Sub langopt_Click()

SendKeys "%" + "+"

If langopt.Value = 1 Then
    langopt.caption = "ÇÎÊíÇÑ ÇáØÈÇÚÉ (Arabic)"
    comfromname.RightToLeft = True
    comreceivedfrom.RightToLeft = True
    txtdetails.RightToLeft = True
Else
    langopt.caption = "ÇÎÊíÇÑ ÇáØÈÇÚÉ (English)"
    comfromname.RightToLeft = False
    comreceivedfrom.RightToLeft = False
    txtdetails.RightToLeft = False
End If
If comfromname.Enabled = True Then
    comfromname.SetFocus
End If
End Sub

Private Sub cmbpaymenttype_Click()
reccurrency.MoveFirst
If Trim(Mid(Trim(cmbpaymenttype.Text), 1, 3)) = "02" Then
    listcurrency.ListItems.clear
    i = 1
    While reccurrency.EOF = False
        If Trim(reccurrency!detail) = "Card" Then
            listcurrency.ListItems.Add , , reccurrency!currency
            listcurrency.ListItems(i).ListSubItems.Add , , Format(reccurrency!rate, "###0.###0")
            listcurrency.ListItems(i).ListSubItems.Add , , reccurrency!latestupdate
            i = i + 1
        End If
        reccurrency.MoveNext
    Wend
    lblcurrency.caption = Trim(listcurrency.ListItems(1).Text)
ElseIf Trim(Mid(Trim(cmbpaymenttype.Text), 1, 3)) = "04" Or Trim(Mid(Trim(cmbpaymenttype.Text), 1, 3)) = "05" Then
    listcurrency.ListItems.clear
    i = 1
    While reccurrency.EOF = False
        If Trim(reccurrency!detail) = "Cash" And IsNull(reccurrency!Status) Then
            listcurrency.ListItems.Add , , reccurrency!currency
            listcurrency.ListItems(i).ListSubItems.Add , , Format(reccurrency!rate, "###0.###0")
            listcurrency.ListItems(i).ListSubItems.Add , , reccurrency!latestupdate
            i = i + 1
            If Trim(reccurrency!Status) = "default" Then
            End If
        End If
        reccurrency.MoveNext
    Wend
Else
    listcurrency.ListItems.clear
    i = 1
    While reccurrency.EOF = False
        If Trim(reccurrency!detail) = "Cash" And Trim(reccurrency!Status) = "default" Then
            listcurrency.ListItems.Add , , reccurrency!currency
            listcurrency.ListItems(i).ListSubItems.Add , , Format(reccurrency!rate, "###0.###0")
            listcurrency.ListItems(i).ListSubItems.Add , , reccurrency!latestupdate
            i = i + 1
        End If
        reccurrency.MoveNext
    Wend
End If

lblcurrency.caption = Trim(listcurrency.ListItems(1).Text)
'end add currency options

If Trim(cmbpaymenttype.Text) <> "" Then
    If Val(Trim(txtdebitamt.Text)) > 0 And (Trim(Mid(Trim(cmbpaymenttype.Text), 1, 3)) = "04" Or Trim(Mid(Trim(cmbpaymenttype.Text), 1, 3)) = "05") Then
        Dim findrate As New ADODB.Recordset
        findrate.Open "Select * from currencytable where currency = '" & Trim(lblcurrency.caption) & "'", CON1, adOpenKeyset, adLockOptimistic
        lbldebitamount.caption = Format(Val(Trim(txtdebitamt.Text)) * Val(findrate!rate), "############0.#0")
        findrate.close
    Else
        lbldebitamount.caption = Format(Trim(txtdebitamt.Text), "############0.#0")
    End If
Else
    'lbldebitamount.Caption = ""
End If


If Val(Mid(Trim(cmbpaymenttype.Text), 1, 2)) <> 3 And Val(Mid(Trim(cmbpaymenttype.Text), 1, 2)) <> 4 Then
    txtchecknumber.Enabled = False
    txtcheckdate.Enabled = False
Else
    txtchecknumber.Enabled = True
    txtcheckdate.Enabled = True
End If

If Val(Mid(Trim(cmbpaymenttype.Text), 1, 2)) = 2 Then
    Label15.caption = "Card No."
    txtchecknumber.Enabled = True
    txtcheckdate.Enabled = False
Else
    Label15.caption = "Chq No."
    txtcheckdate.Enabled = True
End If
End Sub

Private Sub cmbpaymenttype_GotFocus()
    listcurrency.Visible = True
    listcurrency.Top = 1440
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(cmbpaymenttype.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
    
End Sub

Private Sub cmbpaymenttype_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyPageDown
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(cmbpaymenttype.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Select
End Sub

Private Sub cmbpaymenttype_KeyPress(KeyAscii As Integer)
If Trim(cmbpaymenttype.Text) = "" Then
    cmbpaymenttype.SetFocus
    Exit Sub
End If
If KeyAscii = 13 Then
    If txtchecknumber.Enabled = True Then
    txtchecknumber.SetFocus
    Else
    comsetmode.SetFocus
    End If
End If
End Sub

Private Sub cmbpaymenttype_LostFocus()
If Trim(cmbpaymenttype.Text) <> "" Then
    If Val(Trim(txtdebitamt.Text)) > 0 And (Trim(Mid(Trim(cmbpaymenttype.Text), 1, 3)) = "04" Or Trim(Mid(Trim(cmbpaymenttype.Text), 1, 3)) = "05") Then
        Dim findrate As New ADODB.Recordset
        findrate.Open "Select * from currencytable where currency = '" & Trim(lblcurrency.caption) & "'", CON1, adOpenKeyset, adLockOptimistic
        lbldebitamount.caption = Format(Val(Trim(txtdebitamt.Text)) * Val(findrate!rate), "############0.#0")
        findrate.close
    Else
        lbldebitamount.caption = Format(Trim(txtdebitamt.Text), "############0.#0")
    End If
End If

If Trim(Mid(Trim(cmbpaymenttype.Text), 1, 2)) = "10" And Trim(comreceivedfrom.Text) <> "O900002" Then
    MsgBox "Please Check Your Client Code or Payment Mode", vbInformation, "Invalid Choice"
    cmbpaymenttype.SetFocus
    Exit Sub
End If

If Val(Trim(Mid(Trim(comsetmode.Text), 1, 3))) = 7 And Trim(comsetmode.Text) <> "" And Val(Trim(Mid(Trim(cmbpaymenttype.Text), 1, 3))) <> 9 Then
    MsgBox "Please Check Your Payment Mode", vbInformation, "Invalid Choice"
    cmbpaymenttype.SetFocus
    Exit Sub
End If

If Trim(Mid(Trim(cmbpaymenttype.Text), 1, 2)) = "10" Then
    frmreturncheck.Show 1
    Exit Sub
End If

listcurrency.Visible = False
End Sub

Private Sub cmdAdd_Click()
If Trim(comcreditaccountnumber.Text) = "" Or Trim(ccname.Text) = "" Or Val(Trim(creditamount.Text)) <= "0" Then
    MsgBox "Please check the Credit Account Number and Credit Amount", vbInformation, "Empty Data"
    comcreditaccountnumber.SetFocus
    Exit Sub
End If
    
creditamountcount = Val(creditamountcount) + Val(creditamount.Text) 'this is text box
If Val(lbldebitamount.caption) < Val(creditamountcount) Then
    MsgBox "Please Check Your Payment Amount", vbInformation, "Invalid Amount"
    creditamountcount = Val(creditamountcount) - Val(creditamount.Text)
    creditamount.SetFocus
    Exit Sub
End If

If Val(txtdebitamt.Text) = Val(creditamountcount) Then
    'cmdadd.Enabled = False
    cmdsave.SetFocus
Else
    comcreditaccountnumber.SetFocus
End If

'to check all the values
creditcount = creditcount + 1
ListView1.ListItems.Add , , Trim(comcreditaccountnumber.Text)
ListView1.ListItems(creditcount).ListSubItems.Add , , Trim(ccname.Text)
ListView1.ListItems(creditcount).ListSubItems.Add , , Format(Trim(creditamount.Text), "#########0.00#")
If creditcount = 29 Then
    cmdsave.SetFocus
End If
End Sub

Private Sub cmdcancel_Click()

txtdebitamt.Enabled = True
Call prcclear
cmdsave.Enabled = False
cmdcancel.Enabled = False
Frame1.Enabled = False
cmdnewrecord.Enabled = True
txtreceiptnumber.Text = ""
txtreceiptnumber.Enabled = False

combdebitaccnum.Text = ""
comcreditaccountnumber.Text = ""
cname.Text = ""
ccname.Text = ""

On Error Resume Next
rectemp.close
On Error GoTo 0
rectemp.Open "Delete from tempinvoice", CON1, adOpenKeyset, adLockOptimistic
rectemp.Open "delete from checkreturntemp", CON1, adOpenKeyset, adLockOptimistic
dtpinvoicedate.Value = Format(Date, "dd/mm/yyyy")
End Sub
Public Sub cmdCancelclearclear_Click()
txtdebitamt.Enabled = True
Call prcclear
cmdsave.Enabled = False
cmdcancel.Enabled = False
Frame1.Enabled = False
cmdnewrecord.Enabled = True
txtreceiptnumber.Text = ""
txtreceiptnumber.Enabled = False
On Error Resume Next
rectemp.close
On Error GoTo 0
rectemp.Open "Delete from tempinvoice", CON1, adOpenKeyset, adLockOptimistic
End Sub

Private Sub cmdinvoice_Click()
If Trim(comreceivedfrom.Text) <> "" And Val(Trim(Mid(Trim(comsetmode.Text), 1, 5))) = 1 And Trim(txtdebitamt.Text) <> "" And (Mid(Trim(cmbpaymenttype.Text), 1, 2) = "01" Or Mid(Trim(cmbpaymenttype.Text), 1, 2) = "09") Then
frminvoice.activeform = 0
frminvoice.anucustcode = Trim(comreceivedfrom.Text)
On Error GoTo er:
frminvoice.activeform = 1
frminvoice.Show 1
End If
er:
If Err.Number = 360 Then
End If

End Sub

Private Sub cmdnewrecord_Click()
Dim reccurrencychange As New ADODB.Recordset
'the same recordset is checking for any data whether posted or not
reccurrencychange.Open "Select * from glmaster where left(journalno,3)= 'CSR' and recorddate >= '" & Format(Date, "mm/dd/yyyy") & "'", CON1, adOpenKeyset, adLockOptimistic
If reccurrencychange.RecordCount > 0 Then
    MsgBox "Please Check The Date; Because This Date Posted Before" & vbCrLf & " Contact Your System Administrator", vbInformation, "Error Add New"
    reccurrencychange.close
    Exit Sub
End If
reccurrencychange.close
'end checking for the posting

'change the currency to default

reccurrencychange.Open "select * from currencytable where detail = 'Cash' and status='default'", CON1, adOpenKeyset, adLockOptimistic
lblcurrency.caption = Trim(reccurrencychange!currency)
reccurrencychange.close
'end checking for the currency

Dim reccheck As New ADODB.Recordset ' to check whether any transaction for tomorrow
reccheck.Open "select * from vouchers where receiptdate > '" & Format(Date, "mm/dd/yyyy") & "'", CON1, adOpenKeyset, adLockOptimistic
    If reccheck.BOF = False Then
        MsgBox "Please check The Date and Add The Transaction", vbInformation, "Check Date"
    End If
reccheck.close

Frame1.Enabled = True
cmdcancel.Enabled = True
txtreceiptnumber.Enabled = True
comreceivedfrom.Enabled = True
comfromname.Enabled = True
txtdetails.Enabled = True
txtchecknumber.Enabled = True
txtdebitamt.Enabled = True
comsetmode.Enabled = True
'cmdinvoice.Enabled = True

txtinvoicenum.Enabled = True
cmbpaymenttype.Enabled = True
cmdsave.Enabled = True
cmdnewrecord.Enabled = False
cmdprint.Enabled = False
'cmdedit.Enabled = False

recnum.Requery
txtreceiptnumber.Text = Val(recnum!receiptnumber)
txtreceiptnumber.Enabled = False

'Frame2.Enabled = False
txttempreceipt.Enabled = True
txttempreceipt.SetFocus
frminvoice.transferamount = 0
frminvoice.totalvalidatedebitamount = 0
frminvoice.validatedebitamount = 0

Call prcclear
creditamount.Text = " "
fromwho = "c"
End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub CMDEDIT_Click()

recemp.MoveFirst
comverified.clear
comenterd.clear
While recemp.EOF = False
On Error Resume Next
    comverified.AddItem recemp!Userid
    comenterd.AddItem recemp!Userid
    recemp.MoveNext
Wend
On Error GoTo 0
recacc.MoveFirst

Call prcclear
Dim recaccadd As New ADODB.Recordset
creditcount = 0
creditamount = 0
creditamountcount = 0
cmdnewrecord.Enabled = False
CmdEdit.Enabled = False
'cmdadd.Enabled = True
cmdsave.Enabled = True
cmdcancel.Enabled = True
If UCase(cLogUser) = UCase("Cashier") Then
    cmdprint.Enabled = True
End If
txtreceiptnumber.Enabled = True
txtreceiptnumber.Text = ""
SendKeys "{home}+{end}"
txtreceiptnumber.SetFocus

fromwho = "a"
End Sub

Private Sub cmdPrint_Click()
'frmlanguagemessage.Show 1

If MsgBox("Select the Language to Print      ÇÎÊÇÑ áÛÉ ÇáØÈÇÚÉ  " & vbCrLf & "Yes :- For Arabic                                       ÈÇáÚÑÈí " & vbCrLf & "No  :- For English                                  ÈÇáÇäÌáíÒí", vbInformation + vbYesNo, "Language Option  ÇÎÊíÇÑ ÇááÛÉ") = vbYes Then
    frmrecieptvou.langopt.Value = 1
Else
     frmrecieptvou.langopt.Value = 2
End If

'this is for the amount in word
Dim AmtInFigure As Currency
Dim AmtInwords As String
Dim classamount As New HabitatClass

AmtInFigure = IIf(IsNull(txtdebitamt.Text) = True, 0, txtdebitamt.Text)
If langopt.Value = 0 Then
    classamount.EnglishAmountInWords AmtInFigure, AmtInwords, "Riyals", "Halalas"
Else
    classamount.ArabicAmountInWords AmtInFigure, AmtInwords, "ÑíÇá", "åááå"
End If
AmtInwords = classamount.Kupal

On Error Resume Next
Printer.PaperSize = 19
On Error GoTo 0

On Error GoTo er
Printer.FontItalic = False
Printer.FontBold = False
Printer.FontName = "ARABIC TRANSPARENT"
Printer.FontSize = 15
Printer.Print ""
Printer.Print ""
Printer.Print ""
Printer.FontBold = True
Printer.Print ""
Printer.Print ""
Printer.FontSize = 6
Printer.Print ""
Printer.FontSize = 15
receiptx = 6 - Len(Trim(txtreceiptnumber.Text))
ireceipt = 1
For ireceipt = 1 To receiptx
    addzero = addzero & "0"
Next

Printer.Print Tab(22); "Official Receipt No.:-  "; addzero & Trim(txtreceiptnumber.Text); "  ÓäÏ ÞÈÖ  "

'this is for different payment modes
Printer.FontSize = 13
Dim takepaymode As New ADODB.Recordset
takepaymode.Open "Select * from paymode where newcode = '" & Trim(Mid(Trim(cmbpaymenttype.Text), 1, 3)) & "'", CON1, adOpenKeyset, adLockOptimistic, adCmdText
Printer.Print Tab(25); "  #  "; Trim(takepaymode!neweng); "   #"
takepaymode.close
'end  show the payment mode

Printer.FontBold = False
Printer.FontSize = 5
Printer.Print ""
Printer.FontSize = 10
Printer.Print Tab(1); "_____________________________________________________________________________________________________________________________"
Printer.FontSize = 2
Printer.Print ""
Printer.FontSize = 10
Printer.FontBold = True
Printer.Print "Date";
Printer.FontBold = False
Printer.Print Tab(25); Format(Date, "dd/mm/yyyy");
Printer.FontBold = True
Printer.Print Tab(65 - Len("..... ÇáÊÇÑíÎ")); "..... ÇáÊÇÑíÎ";
Printer.Print ; Tab(80); "Amount";
Printer.FontBold = False
Printer.Print Tab(110 - Len(amount & " " & Trim(lblcurrency.caption) & " ")); Format(Trim(txtdebitamt.Text), "###,###,###,###0.#0"); " "; Trim(lblcurrency.caption);
Printer.FontBold = True
Printer.Print Tab(132 - Len("..... ÇáãÈáÛ")); "..... ÇáãÈáÛ"
Printer.FontBold = False
Printer.FontSize = 2
Printer.Print ""


Printer.FontSize = 10
Printer.FontBold = True
Printer.Print "Received From";
Printer.FontBold = False

Printer.FontSize = 8
If langopt.Value = 0 Then
    Printer.Print Tab(31); Trim(Trim(comreceivedfrom.Text) & "  " & UCase(Trim(comfromname.Text)));
Else
    Printer.RightToLeft = True
    Printer.Print Tab(71); Trim(Trim(comreceivedfrom.Text) & " " & UCase(Trim(comfromname.Text)));
    Printer.RightToLeft = False
End If
Printer.FontSize = 10

Printer.FontBold = True
Printer.Print ; Tab(80); "Sales Man";
Printer.FontBold = False
Printer.Print Tab(100); Trim(comsalesman.Text);
Printer.FontBold = True
Printer.Print Tab(132 - Len(".... ÇÓÊáãäÇ ãä")); ".... ÇÓÊáãäÇ ãä"
Printer.FontSize = 2
Printer.Print ""


Printer.FontSize = 10
Printer.FontBold = True
Printer.Print "Amount in Words";
Printer.FontBold = False

If langopt.Value = 0 Then
    Printer.Print Tab(25); Trim(AmtInwords);
Else
    Printer.RightToLeft = True
    Printer.FontSize = 9
    Printer.Print Tab(23); Trim(AmtInwords);
    Printer.RightToLeft = False
End If
Printer.FontSize = 10

Printer.FontBold = True
Printer.Print Tab(132 - Len("..... ÇáãÈáÛ ÈÇáÍÑæÝ")); "..... ÇáãÈáÛ ÈÇáÍÑæÝ"
Printer.FontSize = 2
Printer.Print ""


Printer.FontSize = 10
Printer.Print "Payment Mode";
Printer.FontBold = False

namenamenum = InStr(1, Trim(cmbpaymenttype.Text), "~", vbTextCompare)
nameara = Trim(Mid(Trim(cmbpaymenttype.Text), (namenamenum + 1), 30))
If namenamenum > 0 Then
NameEng = Mid(Trim(cmbpaymenttype.Text), 1, (namenamenum - 1))
Else
NameEng = Trim(cmbpaymenttype.Text)
End If

If langopt.Value = 0 Then
    Printer.Print Tab(25); Trim(Mid(NameEng, 3, 30));
Else
    Printer.RightToLeft = True
    Printer.Print Tab(21); Trim(nameara);
    Printer.RightToLeft = False
End If
Printer.FontBold = True
Printer.Print Tab(133 - Len("....... äæÚ ÇáÏÝÚ ")); "....... äæÚ ÇáÏÝÚ"
Printer.FontSize = 2
Printer.Print ""
    
    If Trim(Mid(Trim(cmbpaymenttype.Text), 1, 3)) = "03" Or Trim(Mid(Trim(cmbpaymenttype.Text), 1, 3)) = "02" Then
    Printer.FontBold = True
    Printer.FontSize = 10
    Printer.Print "Reference No";
    Printer.FontBold = False
    Printer.Print Tab(25); Trim(txtchecknumber.Text);
    Printer.FontBold = True
    Printer.Print Tab(51); "Due Date";
    Printer.FontBold = False
    Printer.Print Tab(70); Format(txtcheckdate.Value, "dd/mm/yyyy");
    Printer.FontBold = True
    Printer.Print Tab(133 - Len("..... Çáíæã ãÓÊÍÞ ÊÇÑíÎ")); "..... Çáíæã ãÓÊÍÞ ÊÇÑíÎ"
    Printer.FontSize = 10
    ElseIf Trim(Mid(Trim(cmbpaymenttype.Text), 1, 3)) = "10" Then
        Printer.FontBold = True
        Printer.FontSize = 10
        Printer.Print "Check No & Date ";
        Printer.FontBold = False
        Printer.Print Tab(25);
        Printer.FontSize = 8
        Printer.Print Trim(Mid(checknumberanddate, 1, 106))
        If Trim(Mid(checknumberanddate, 106, 211)) <> "" Then
            Printer.FontSize = 10
            Printer.Print Tab(25);
            Printer.FontSize = 8
            Printer.Print Trim(Mid(checknumberanddate, 106, 211))
        End If
        If Trim(Mid(checknumberanddate, 212, 317)) <> "" Then
            Printer.FontSize = 10
            Printer.Print Tab(25);
            Printer.FontSize = 8
            Printer.Print Trim(Mid(checknumberanddate, 212, 317))
        End If
    End If
Printer.FontSize = 2
Printer.Print ""

Printer.FontBold = True
Printer.FontSize = 10
Printer.Print "Payment Againts";
Printer.FontBold = False

namenamenum = InStr(1, Trim(comsetmode.Text), "~", vbTextCompare)
nameara = Trim(Mid(Trim(comsetmode.Text), (namenamenum + 1), 30))
If namenamenum > 0 Then
NameEng = Mid(Trim(comsetmode.Text), 1, (namenamenum - 1))
Else
NameEng = Trim(comsetmode.Text)
End If

If langopt.Value = 0 Then
    Printer.Print Tab(25); Trim(Trim(Mid(NameEng, 4, 30)));
Else
    Printer.RightToLeft = True
    Printer.Print Tab(21); Trim(nameara);
    Printer.RightToLeft = False
End If
Printer.FontBold = True
Printer.Print Tab(133 - Len("....... ÇáÏÝÚ ãÞÇÈá")); "....... ÇáÏÝÚ ãÞÇÈá"
Printer.FontBold = False
'this will print all invoice numbers
If Trim(Mid(Trim(comsetmode.Text), 1, 4)) = "001" Then
    If langopt.Value = 0 Then
        Printer.Print Tab(25); Trim(txtinvoicenum.Text); Trim("- " & Mid(Trim(txtallinvoice.Text), 2, Len(Trim(txtallinvoice.Text))))
    Else
        Printer.RightToLeft = True
        Printer.Print Tab(21); Trim(txtinvoicenum.Text); Trim("- " & Mid(Trim(txtallinvoice.Text), 2, Len(Trim(txtallinvoice.Text))))
        Printer.RightToLeft = False
    End If
End If
Printer.FontSize = 2
Printer.Print ""

Printer.FontBold = True
Printer.FontSize = 10
Printer.Print "Description ";
Printer.FontBold = False
If langopt.Value = 0 Then
    Printer.Print Tab(25); Trim(txtdetails.Text);
Else
    Printer.RightToLeft = True
    Printer.Print Tab(21); Trim(txtdetails.Text);
    Printer.RightToLeft = False
End If
Printer.FontBold = True
Printer.Print Tab(133 - Len("..... ÇáÈíÇä ")); "..... ÇáÈíÇä "
Printer.FontSize = 2
Printer.Print ""
Printer.FontSize = 10
Printer.Print "Account Numbers:-";

'this is for account number
Printer.RightToLeft = True
For i = 1 To ListView1.ListItems.Count
    Printer.Print Tab(19); Trim(ListView1.ListItems(i).ListSubItems(2).Text);
    Printer.Print Tab(68); Trim(ListView1.ListItems(i).Text);
    Printer.Print Tab(90); i
Next
Printer.RightToLeft = False
'end accountnumber

Printer.Print Tab(1); "__________________________________________________________________________________________________________________________________________"
Printer.FontSize = 2
Printer.Print ""
Printer.Print ""

Printer.FontSize = 10
Printer.Print ""

Printer.Print Tab(12); "............................."; Tab(53); "............................."; Tab(95); "............................."
Printer.FontSize = 10
Printer.Print Tab(13); "  (Manager/ÇáãÏíÑ)    "; Tab(51); "     (Auditor/ÇáãÑÇÌÚ)    "; Tab(92); "     (Cashier/Ããíä ÇáÕäÏæÞ)    "
'Printer.Print ""
Printer.FontSize = 6
Printer.Print ""

Printer.FontSize = 10
Printer.FontBold = True
Printer.FontItalic = True
If Trim(txttempreceipt.Text) <> "" Then
Printer.Print Tab(1); "Temporary Invoice Number is  :-  " & Trim(txttempreceipt.Text)
Printer.FontItalic = False
End If
On Error GoTo er:
Printer.EndDoc
MsgBox "Receipt Voucher Printed Successfully" & vbCrLf & " Êã ÇáØÈÚ ÈäÌÇÍ ", vbInformation, "Print Conformation"

er:
If Err.Number = 482 Then
    If MsgBox("Please check your Printer, Turn ON And Press Yes." & vbCrLf & " ÑÇÌÚ ÇáØÇÈÚå .ÇÝÊÍåÇ æÇÖÛØ ãæÇÝÞ", vbYesNo, "Not Ready") = vbYes Then
        cmdPrint_Click
    Else
        Exit Sub
    End If
End If

End Sub

Private Sub cmdsave_Click()

Dim findrate As New ADODB.Recordset
findrate.Open "Select * from currencytable where status = 'default'", CON1, adOpenKeyset, adLockOptimistic
    If findrate!currency <> Trim(lblcurrency.caption) Then
        If (Trim(Mid(Trim(cmbpaymenttype.Text), 1, 3)) <> "02" And Trim(Mid(Trim(cmbpaymenttype.Text), 1, 3)) <> "04" And Trim(Mid(Trim(cmbpaymenttype.Text), 1, 3)) <> "05") Then
            MsgBox "Please Check Your Currency That is Not Compatable", vbInformation, "Invalid Choice"
            listcurrency.Visible = True
            findrate.close
            Exit Sub
        End If
    Else
        If (Trim(Mid(Trim(cmbpaymenttype.Text), 1, 3)) = "02" And Trim(Mid(Trim(cmbpaymenttype.Text), 1, 3)) = "04" And Trim(Mid(Trim(cmbpaymenttype.Text), 1, 3)) = "05") Then
            MsgBox "Please Check Your Currency That is Not Compatable", vbInformation, "Invalid Choice"
            listcurrency.Visible = True
            findrate.close
            Exit Sub
        End If
    End If
findrate.close

If Trim(comfromname.Text) = "" Then
    MsgBox "Please Check Your Client Name That is Missing", vbInformation, "Empty Client Name"
    comfromname.SetFocus
    Exit Sub
End If

If Val(txtdebitamt.Text) <= 0 Then
    MsgBox "Please Check Your Receipt Amount", vbInformation, "Amount Missing"
    txtdebitamt.SetFocus
    Exit Sub
End If

If Trim(Mid(Trim(cmbpaymenttype.Text), 1, 3)) = "10" And Trim(checknumberanddate) = "" Then
    MsgBox "Please Choose the Checks From The List", vbInformation, "Empty Numbers"
    cmbpaymenttype.SetFocus
    Exit Sub
End If

If Trim(Mid(Trim(cmbpaymenttype.Text), 1, 3)) = "03" Or Trim(Mid(Trim(cmbpaymenttype.Text), 1, 3)) = "04" Then
    If IsNull(txtcheckdate.Value) = True Then
        MsgBox "Your Check Due Date That is Missing", vbInformation, "No check Date"
        Exit Sub
    End If
    
    If Trim(txtchecknumber.Text) = "" Then
        MsgBox "Your check Number That is Missing", vbInformation, "Check Number is missing"
        txtchecknumber.SetFocus
        Exit Sub
    End If
End If

If Trim(cmbpaymenttype.Text) = "" Then
    MsgBox "Plese check Your Payment Type", vbInformation, "Invalid PaymentMode"
    cmbpaymenttype.SetFocus
    Exit Sub
End If

If Trim(comsetmode.Text) = "" Then
    MsgBox "Plese check Your Payment For", vbInformation, "Invalid PaymentFor"
    comsetmode.SetFocus
    Exit Sub
End If

If Trim(comsalesman.Text) = "" Then
    MsgBox "Plese check SalesMan Number & Name", vbInformation, "Empty SalesMan"
    comsalesman.SetFocus
    Exit Sub
End If

If Trim(combdebitaccnum.Text) = "" Or Trim(cname.Text) = "" Then
    MsgBox "Plese check Debit Account Number or Debit Account Name", vbInformation, "Invalid PaymentMode"
    If combdebitaccnum.Enabled = True Then
        combdebitaccnum.SetFocus
    End If
    Exit Sub
End If

ListCount = ListView1.ListItems.Count ' this is to check the analysis
For i = 1 To ListCount
    checkallamount = checkallamount + Val(Trim(ListView1.ListItems(i).ListSubItems(2).Text))
Next

If Val(checkallamount) <> Val(Trim(lbldebitamount.caption)) Then
    MsgBox "Please Check Credit Amount it is not Equal to Debit Amount", vbInformation, "Amount Overflow"
    checkallamount = 0
    Exit Sub
End If
checkallamount = 0 ' end analysis

frmconformpassword.Show 1

If stopstop1 = 1 Then
        stopstop1 = 0
        Exit Sub
End If
    
On Error Resume Next
rectemp.close
On Error GoTo 0

rectemp.Open "Select * from tempinvoice", CON1, adOpenKeyset, adLockOptimistic
txtdebitamt.Enabled = True

recor.addnew
findrate.Open "Select * from currencytable where currency = '" & Trim(lblcurrency.caption) & "'", CON1, adOpenKeyset, adLockOptimistic
    
    recor!currencyrate = Val(findrate!rate)
    recor!currencymark = Trim(lblcurrency.caption)
    
    If findrate!Status = "default" Then
        recor!currencydefault = 1
    End If
findrate.close
'add the currency rate to voucher table

recor!receiptno = Trim(txtreceiptnumber.Text)
recor!receiptdate = Format(Date, "mm/dd/yyyy")
recor!custno = Trim(comreceivedfrom.Text)
recor!custname = Trim(comfromname.Text)
recor!salesman = Trim(comsalesman.Text)
recor!invoicedate = Format(dtpinvoicedate.Value, "mm/dd/yyyy")
recor!deleted = "0"
recor!okprint = "0"

If Trim(Mid(Trim(cmbpaymenttype.Text), 1, 3)) = "10" Then
    recor!moderef = "Return Checks"
Else
    recor!moderef = Trim(txtchecknumber.Text)
End If
recor!receiptamount = Trim(txtdebitamt.Text)
recor!payopt = Trim(comsetmode.Text)
recor!paymode = Trim(cmbpaymenttype.Text)
recor!optref = Trim(txtinvoicenum.Text) & " "

If IsNull(txtcheckdate.Value) = False Then
    recor!chkdue = Format(txtcheckdate.Value, "mm/dd/yyyy")
End If
    recset.Requery 'this is for take the journal number
    totaljournal = "CSR-" & recset!CurrentMoYr & "-" & Trim(recset!nextjn)
    recor!JournalNumber = totaljournal
    totaljournal = "0000" & (Val(Trim(recset!nextjn)) + 1)
    recset!nextjn = Right(totaljournal, 5)
    recset.Update

recor!Post = "no"
recor!svoucher = "Collections"
recor!accountnumber = Trim(combdebitaccnum.Text)

'this is for find the mother name
getmothername Trim(combdebitaccnum.Text), noname, cc
recor!mothername = cc
lenlen = 0
instrlen = 0
lenlen = Len(Trim(cname.Text))
instrlen = InStr(1, Trim(cname.Text), "\", vbTextCompare)
If instrlen > 0 Then
    recor!accountname = Mid(Trim(cname.Text), 1, instrlen - 1)
    instrlen = lenlen - instrlen
    recor!accountnamearab = Right(Trim(cname.Text), instrlen)
Else
    recor!accountname = Trim(cname.Text)
    recor!accountnamearab = "  "
End If

recor!againtsnumber = Trim(ListView1.ListItems(1).Text)
recor!againtsname = Trim(ListView1.ListItems(1).ListSubItems(1).Text)

recor!LogUser = "Last Modified By : " & cLogUser
lastuser = "Last Modified By : " & cLogUser
recor!tempinvoice = Trim(txttempreceipt.Text)
recor!remarks = Trim(txtdetails.Text)
recor.Update

cmdsave.Enabled = False
recnum!receiptnumber = Val(Val(Trim(txtreceiptnumber.Text)) + 1)
recnum.Update

'this is all for temporary tales
Dim recbulk As New ADODB.Recordset
recbulk.Open "delete from tempinvoice where applied <=0", CON1, adOpenKeyset, adLockOptimistic
rectemp.Requery

If rectemp.BOF = False And Mid(Trim(comsetmode.Text), 1, 4) = "001" Then
    'this is like bulk copy tempinvoice2
        recbulk.Open "INSERT INTO tempinvoice2 (receiptno,custid,invoiceno,receiptdate,amount,unpaid,applied)" & _
        "SELECT '" & Trim(txtreceiptnumber.Text) & "',cusnumber,invoicenumber,invdate,amount,unpaid,applied FROM tempinvoice where applied > 0 order by num", CON1, adOpenKeyset, adLockOptimistic
        
        Dim recfindsj As New ADODB.Recordset ' this is to update the sjmaster table
        Dim consj As New ADODB.Connection
        Dim conanumar As New ADODB.Connection ' this is for the creditmain connection
        Dim reccre As New ADODB.Recordset
        conanumar.Mode = adModeShareDenyNone
        conanumar.Open "Dsn=anufoxpro;uid=sa;pwd=;"
        consj.Open "Dsn=anufoxpro;uid=sa;pwd=;"
        
        rectemp.MoveFirst
        While rectemp.EOF = False
            If Trim(rectemp!Applied) > 0 Then
                recfindsj.Open "Select * from sjmaster where invc_no ='" & Trim(rectemp!InvoiceNumber) & "'", consj, adOpenKeyset, adLockOptimistic
                recfindsj!unpaidamt = Trim(rectemp!unpaid)
                recfindsj!paidamt = Val(Trim(recfindsj!paidamt)) + Val(Trim(rectemp!Applied))
                recfindsj.Update
                recfindsj.close
                
                If Trim(rectemp!Applied) > 0 Then 'this is for sjmaster
                    If Mid(Trim(rectemp!InvoiceNumber), 1, 3) <> "ODN" Then
                      recfindsj.Open "Select * from sjmaster where invc_no ='" & Trim(rectemp!InvoiceNumber) & _
                      "'", consj, adOpenKeyset, adLockOptimistic
                      
                      If recfindsj.BOF = False Then
                          recfindsj!unpaidamt = Trim(rectemp!unpaid)
                          recfindsj!paidamt = Val(Trim(recfindsj!paidamt)) + Val(Trim(rectemp!Applied))
                            If recfindsj!unpaidamt = 0 Then
                                recfindsj!paid = True
                            End If
                          recfindsj.Update
                      End If
                      recfindsj.close
                    Else 'this is for credmain updation
                      reccre.Open "Select * from credmain where invc_no ='" & Trim(rectemp!InvoiceNumber) & _
                      "'", consj, adOpenKeyset, adLockOptimistic
                      
                      If reccre.BOF = False Then
                          reccre!paidamt = Val(Trim(reccre!paidamt)) + Val(Trim(rectemp!Applied))
                              If Val(reccre!tot_amt) - Val(Trim(reccre!paidamt)) = 0 Then
                                  reccre!Paidmark = "1"
                              End If
                          reccre.Update
                      End If
                      reccre.close
                    End If
                End If
            End If
'                ProgressBar1.Visible = True
'                ProgressBar1.Min = 0
'                ProgressBar1.Max = rectemp.RecordCount + 1
'                If ProgressBar1.Value <> ProgressBar1.Max Then
'                    ProgressBar1.Value = ProgressBar1.Value + 1
'                Else
'                    ProgressBar1.Value = 0
'                End If
            rectemp.MoveNext
        Wend
        consj.close
        conanumar.close
        rectemp.MoveFirst
        If Trim(rectemp!Applied) > 0 Then   'this is for add  menu to tempagaintsinvoice
             recteinvoice12.addnew
             recteinvoice12!receiptno = Trim(txtreceiptnumber.Text)
             recteinvoice12!invoiceno = "Invoice Number"
             recteinvoice12!Menu = "Amount"
             recteinvoice12!svoucher = "Collections"
             recteinvoice12!custno = Trim(comreceivedfrom.Text)
             recteinvoice12.Update
             
             recbulk.Open "INSERT INTO tempagaintsinvoice SELECT '" & Trim(txtreceiptnumber.Text) & _
             "',invoicenumber,cusnumber,applied,' ','Collections',display,rtrim(ltrim(left(invdate,15)))" & _
             " FROM tempinvoice where applied > 0", CON1, adOpenKeyset, adLockOptimistic
               
             recteinvoice12.addnew ' to record total amount
             recteinvoice12!receiptno = Trim(txtreceiptnumber.Text)
             recteinvoice12!invoiceno = "Invoice SubTotal"
             recteinvoice12!Applied = Val(Trim(txtdebitamt.Text)) - Val(rectemp!unappliedbalance)
             recteinvoice12!svoucher = "Collections"
             recteinvoice12!invoicedate = Mid(Trim(rectemp!InvDate), 1, 12)
             recteinvoice12!custno = Trim(comreceivedfrom.Text)
             recteinvoice12.Update
             
             If Val(Trim(rectemp!unappliedbalance)) > 0 Then
                recteinvoice12.addnew ' to record un applied amount
                recteinvoice12!receiptno = Trim(txtreceiptnumber.Text)
                recteinvoice12!invoiceno = "Un Applied Amount"
                recteinvoice12!Applied = Trim(rectemp!unappliedbalance)
                recteinvoice12!svoucher = "Collections"
                recteinvoice12!invoicedate = Mid(Trim(rectemp!InvDate), 1, 12)
                recteinvoice12!custno = Trim(comreceivedfrom.Text)
                recteinvoice12.Update
             End If
             notequla = 1
         End If
End If

Dim recdelete As New ADODB.Recordset ' this is to clear the tempinvoice table
recdelete.Open "Delete from tempinvoice", CON1, adOpenKeyset, adLockOptimistic
      
'ProgressBar1.Visible = False
'progresslabel.Visible = False
'cmdadd.Enabled = True
cmdcancel.Enabled = False
creditcount = 0
creditamountcount = 0
Frame1.Enabled = False
fromwho = "none"
cmdprint.Enabled = True

'to insert the first ticket to cashjournal
If Trim(Mid(Trim(cmbpaymenttype.Text), 1, 3)) = "07" Then
    recdelete.Open "insert into cashjournal SELECT JournalNumber, '1', '0', Accountnumber, AccountName, accountnamearab, mothername, " & _
    "convert(nvarchar(10),vouchers.receiptno) + ' \ ' + isnull(vouchers.remarks,'') + '\'+ ltrim(substring(vouchers.paymode,4,20)) + '\' + isnull(vouchers.moderef,'') + '-' + isnull(convert(nvarchar(11),vouchers.chkdue),'') " & _
    ",againtsname+' \ '+'" & lastuser & "',ReceiptNo,'" & Format(Date, "mm/dd/yyyy") & "',receiptamount,0,NULL,'UnPosted','R' FROM vouchers where receiptno = '" & Trim(txtreceiptnumber.Text) & "' and svoucher = 'Collections'", CON1, adOpenKeyset, adLockOptimistic
Else
    recdelete.Open "insert into cashjournal SELECT JournalNumber, '1', '0', Accountnumber, AccountName, accountnamearab, mothername, " & _
    "convert(nvarchar(10),vouchers.receiptno) + ' \ ' + isnull(vouchers.remarks,'') + '\'+ ltrim(substring(vouchers.paymode,4,20)) + '\' + isnull(vouchers.moderef,'') + '-' + isnull(convert(nvarchar(11),vouchers.chkdue),'') " & _
    ",againtsname+' \ '+'" & lastuser & "',ReceiptNo,'" & Format(Date, "mm/dd/yyyy") & "',receiptamount,0,NULL,'UnPosted','R' FROM vouchers where receiptno = '" & Trim(txtreceiptnumber.Text) & "' and svoucher = 'Collections'", CON1, adOpenKeyset, adLockOptimistic
End If

'end add data to the
Dim reccashjournal As New ADODB.Recordset
reccashjournal.Open "Select * from cashjournal", CON1, adOpenKeyset, adLockOptimistic
    For xxx = 11 To ListCount + 10
        With reccashjournal
            .addnew
            !SerialNo = recor!JournalNumber
            !ticket = xxx - 9
            !deletemark = "0"
            !accountnumber = Trim(ListView1.ListItems(xxx - 10).Text)
            lenlen = Len(Trim(ListView1.ListItems(xxx - 10).ListSubItems(1).Text))
            instrlen = InStr(1, Trim(ListView1.ListItems(xxx - 10).ListSubItems(1).Text), "\", vbTextCompare)
            
            If instrlen > 0 Then
            !accountname = Mid(Trim(ListView1.ListItems(xxx - 10).ListSubItems(1).Text), 1, instrlen - 1)
            instrlen = lenlen - instrlen
            !accountnamearab = Right(Trim(ListView1.ListItems(xxx - 10).ListSubItems(1).Text), instrlen)
            Else
            !accountname = Trim(ListView1.ListItems(xxx - 10).ListSubItems(1).Text)
            !accountnamearab = "  "
            End If
            
            getmothername Trim(ListView1.ListItems(xxx - 10).Text), noname, cc
            !mothername = cc
            
            !cashreceiptno = Trim(txtreceiptnumber.Text)
            !Description = Trim(cname.Text) & " \ " & "Last Modified By : " & cLogUser
            !TRansDate = Format(Date, "mm/dd/yyyy")
            addParticulars = recor!receiptno & " \ " & recor!remarks & " \  " & Trim(Mid(recor!paymode, 3, 20)) & IIf(Trim(recor!moderef) = "", " ", " \ " & recor!moderef) & IIf(IsNull(Trim(recor!chkdue)), " ", " - " & Trim(recor!chkdue))
            !particulars = Trim(addParticulars) ' later this will transfer the particulars
            !creditamount = Trim(ListView1.ListItems(xxx - 10).ListSubItems(2).Text)
            !DebitAmount = 0
            !Status = "UnPosted"
            !Trantype = "R"
            .Update
'            If ProgressBar1.Value <> ProgressBar1.Max Then
'                ProgressBar1.Value = ProgressBar1.Value + 1
'            Else
'                ProgressBar1.Value = 0
'            End If
            addParticulars = ""
        End With
Next
        reccashjournal.close
   
MsgBox "Record Updated Successfully" & vbCrLf & "  Êã ÇáÍÝÙ", vbInformation, "Data Saved"
cmdPrint_Click
cmdsave.Enabled = True
Printer.FontItalic = False
'ProgressBar1.Visible = False
cmdsave.Enabled = False
txtreceiptnumber.Enabled = False
'cmdadd.Enabled = True
cmdcancel.Enabled = False
creditcount = 0
creditamountcount = 0
Frame1.Enabled = False
cmdnewrecord.Enabled = True
txttempreceipt.Enabled = False
txtchecknumber.MaxLength = 8
Me.caption = "Receipt Voucher...                                                                                                                                                ÓäÏ ÞÈÖ  "
cc = 0

'this is for add listview2 for viewing purpose
reclist.Requery
l2 = 0
        If reclist.BOF = False Then
            reclist.MoveFirst
            ListView2.ListItems.clear
                While reclist.EOF = False
                    With ListView2
                        l2 = l2 + 1
                        .ListItems.Add , , Trim(reclist!receiptno)
                        .ListItems(l2).ListSubItems.Add , , Trim(reclist!receiptdate)
                        .ListItems(l2).ListSubItems.Add , , Trim(reclist!custname)
                        .ListItems(l2).ListSubItems.Add , , " " & reclist!paymode 'IsNull(Trim(reclist!paymode), " ", Trim(reclist!paymode))
                        .ListItems(l2).ListSubItems.Add , , " " & Trim(reclist!payopt)
                        .ListItems(l2).ListSubItems.Add , , Format(Trim(receiptamount), "############0.#0")
                    End With
                reclist.MoveNext
                 DoEvents
                Wend
        End If
End Sub


Private Sub cmdshowclient_Click()
connectclientcode.Combo1.Text = Trim(comreceivedfrom.Text)
connectclientcode.Show 1
End Sub

Private Sub cmdshowinvoice_Click()
If Trim(comreceivedfrom.Text) <> "" Then  'And Val(Trim(Mid(Trim(comsetmode.Text), 1, 5))) = 1 And Trim(txtdebitamt.Text) <> "" Then
frminvoice.activeform = 0
frminvoice.anucustcode = Trim(comreceivedfrom.Text)
frminvoice.receiptno = Trim(txtreceiptnumber.Text)
On Error GoTo er:
frminvoice.Label5.Visible = False
frminvoice.Label6.Visible = False
'frminvoice.Label7.Visible = False
frminvoice.cmdautodistribute.Visible = False
frminvoice.l8.Visible = False
frminvoice.ListView1.Enabled = False
frminvoice.cmdcancel.Visible = False
frminvoice.cmdclose.caption = "Close"
frminvoice.Timer1.Interval = 0
frminvoice.Show 1
End If

er:
If Err.Number = 360 Then
End If

End Sub

Private Sub cname_Click()
Dim recfindacc As New ADODB.Recordset
namenamenum = InStr(1, Trim(cname.Text), "\", vbTextCompare)

If namenamenum > 0 Then
namename = Mid(Trim(cname.Text), 1, (namenamenum - 1))
Else
namename = Trim(cname.Text)
End If
recfindacc.Open "Select * from financemaster where active <> '0' and accountnameeng = " & "'" & namename & "'", CON1, adOpenKeyset, adLockOptimistic
If recfindacc.BOF = False Then
    combdebitaccnum.Text = recfindacc!AccountCode
End If
 recfindacc.close
 'this is for the mother name
getmothername nonumber, namename, cc
Me.caption = "Receipt Voucher  " & cc
'end mother name

End Sub

Private Sub cname_GotFocus()
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(cname.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)

End Sub

Private Sub cname_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyPageDown
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(cname.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Select

End Sub

Private Sub cname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Trim(cname.Text) <> "" Then
    comcreditaccountnumber.SetFocus
End If
End Sub

Private Sub cname_LostFocus()
Dim recfindacc As New ADODB.Recordset
namenamenum = InStr(1, Trim(cname.Text), "\", vbTextCompare)
If namenamenum > 0 Then
namename = Mid(Trim(cname.Text), 1, (namenamenum - 1))
End If
recfindacc.Open "Select * from financemaster where active <> '0' and accountnameeng = " & "'" & namename & "'", CON1, adOpenKeyset, adLockOptimistic
If recfindacc.BOF = True Then
    MsgBox "Please choose the Correct Account", vbInformation, "Invalid accountNumber"
    cname.SetFocus
    Exit Sub
Else
    combdebitaccnum.Text = recfindacc!AccountCode
End If
recfindacc.close

End Sub

Private Sub combdebitaccnum_Click()
'this is for find the account name
Dim recfindaccount As New ADODB.Recordset
recfindaccount.Open "select * from financemaster where active <> '0' and accountcode = " & "'" & Trim(combdebitaccnum.Text) & "'", CON1, adOpenKeyset, adLockOptimistic
If recfindaccount.BOF = False Then
    cname.Text = Trim(recfindaccount!accountnameeng) & "\" & Trim(recfindaccount!accountnamearab)
End If
recfindaccount.close
'end find the account name

'this is for the mother name
getmothername Trim(combdebitaccnum.Text), noname, cc
Me.caption = "Receipt Voucher  " & cc
'end mother name
End Sub

Private Sub combdebitaccnum_GotFocus()
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(combdebitaccnum.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)

End Sub

Private Sub combdebitaccnum_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyPageDown
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(combdebitaccnum.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Select

End Sub

Private Sub combdebitaccnum_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then 'And Trim(combdebitaccnum.Text) <> ""
    cname.SetFocus
End If
End Sub

Private Sub combdebitaccnum_LostFocus()
Dim recfindaccount As New ADODB.Recordset
If Trim(combdebitaccnum.Text) <> "" Then
recfindaccount.Open "select * from financemaster where active <> '0' and accountcode = " & "'" & Trim(combdebitaccnum.Text) & "'", CON1, adOpenKeyset, adLockOptimistic
If recfindaccount.BOF = True Then
    MsgBox "Please Choose the Correct Account", vbInformation, "Invalid AccountNumber"
    combdebitaccnum.SetFocus
    Exit Sub

Else
    cname.Text = recfindaccount!accountnameeng & "\" & recfindaccount!accountnamearab

End If
recfindaccount.close
End If
Me.caption = "Receipt Voucher  " & cc
End Sub

Private Sub comcreditaccountnumber_Click()
Dim recfindaccount As New ADODB.Recordset
recfindaccount.Open "select * from financemaster where active <> '0' and accountcode = " & "'" & Trim(comcreditaccountnumber.Text) & "'", CON1, adOpenKeyset, adLockOptimistic
If recfindaccount.BOF = False Then
    ccname.Text = Trim(recfindaccount!accountnameeng) & "\" & Trim(recfindaccount!accountnamearab)
End If
recfindaccount.close
'this is for mother name
getmothername Trim(comcreditaccountnumber.Text), noname, c
Me.caption = "Receipt Voucher  " & cc
End Sub

Private Sub comcreditaccountnumber_GotFocus()
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(comcreditaccountnumber.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)

End Sub

Private Sub comcreditaccountnumber_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyPageDown
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(comcreditaccountnumber.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Select

End Sub

Private Sub comcreditaccountnumber_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then 'And Trim(comcreditaccountnumber.Text) <> ""
    ccname.SetFocus
End If
End Sub

Private Sub comcreditaccountnumber_LostFocus()
Dim recfindaccount As New ADODB.Recordset
If Trim(comcreditaccountnumber.Text) <> "" Then
recfindaccount.Open "select * from financemaster where active <> '0' and accountcode = " & "'" & Trim(comcreditaccountnumber.Text) & "'", CON1, adOpenKeyset, adLockOptimistic
If recfindaccount.BOF = True Then
    MsgBox "Please Choose the Currect Account", vbInformation, "Invalid AccountNumber"
    comcreditaccountnumber.SetFocus
    Exit Sub
Else
    ccname.Text = recfindaccount!accountnameeng & "\" & recfindaccount!accountnamearab

End If
recfindaccount.close
End If
Me.caption = "Receipt Voucher  " & cc
End Sub

Private Sub comfromname_Click()
X = comfromname.ListIndex
comreceivedfrom.ListIndex = X

End Sub

Private Sub comfromname_GotFocus()
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(comfromname.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)

End Sub

Private Sub comfromname_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyPageDown
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(comfromname.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Select

End Sub

Private Sub comfromname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtdebitamt.Enabled = True Then
        txtdebitamt.SetFocus
    Else
        txtinvoicenum.SetFocus
    End If
End If

End Sub


Private Sub comreceivedfrom_Click()
X = comreceivedfrom.ListIndex
 comfromname.ListIndex = X

End Sub

Private Sub comreceivedfrom_GotFocus()
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(comreceivedfrom.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)

End Sub

Private Sub comreceivedfrom_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyPageDown
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(comreceivedfrom.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Select

End Sub

Private Sub comreceivedfrom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Trim(comreceivedfrom.Text) <> "" Then
    comreceivedfrom_LostFocus
End If
If KeyAscii = 13 And Trim(comreceivedfrom.Text) = "" Then
    comfromname.SetFocus
End If
End Sub

Private Sub comreceivedfrom_LostFocus()
If Trim(comreceivedfrom.Text) <> "" Then
anu = comreceivedfrom.ListCount
i = 0
If Trim(comreceivedfrom.Text) <> "" Then
Do While i <= anu
    If Trim(comreceivedfrom.Text) = comreceivedfrom.List(i) Then
        doneraja = i
        Exit Do
    End If
    i = i + 1
Loop
'Else
'    doneraja = 0
End If

comfromname.ListIndex = doneraja

    If doneraja <= 0 And Trim(comreceivedfrom.Text) <> "" Then
        MsgBox "Please choose the Currect Account Number", vbInformation, "Invalid Name"
        comreceivedfrom.SetFocus
        Exit Sub
    End If
comfromname.SetFocus
End If
End Sub

Private Sub comsalesman_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyPageDown
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(comsalesman.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Select

End Sub

Private Sub comsalesman_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtdetails.SetFocus
End If


End Sub

Private Sub comsetmode_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyPageDown
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(comsetmode.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Select

End Sub

Private Sub comsetmode_KeyPress(KeyAscii As Integer)
If Trim(comsetmode.Text) = "" Then
    comsetmode.SetFocus
    Exit Sub
End If

If txtinvoicenum.Enabled = True Then
    txtinvoicenum.SetFocus
    Exit Sub
End If
If combdebitaccnum.Enabled = True Then
    combdebitaccnum.SetFocus
    Exit Sub
End If
If KeyAscii = 13 Then
    cmdsave.SetFocus
End If
End Sub

Private Sub comsetmode_LostFocus()
If Val(Trim(Mid(Trim(comsetmode.Text), 1, 6))) = 6 Then
    MsgBox "This is Not Available", vbInformation, "Wrong Choice"
    comsetmode.ListIndex = 0
    comsetmode.SetFocus
    Exit Sub
End If

If Val(Trim(Mid(Trim(comsetmode.Text), 1, 6))) = 1 Then
    lblref.caption = "Inv No."
Else
    lblref.caption = "Ref No."
End If

If Val(Trim(Mid(Trim(comsetmode.Text), 1, 6))) = 2 Or Val(Trim(Mid(Trim(comsetmode.Text), 1, 6))) = 5 Then
    MsgBox "Please This Method is Not Available" & vbCrLf & " åÐå ÇáÎÇÕíå ÛíÑ ãÊÇÍå", vbInformation, "Invalid MOde"
    comsetmode.Text = comsetmode.List(0)
    comsetmode.SetFocus
    Exit Sub
End If

If Val(Trim(Mid(Trim(cmbpaymenttype.Text), 1, 3))) = 9 And Trim(Mid(Trim(comsetmode.Text), 1, 5)) <> "007" Then
    MsgBox "Please Check the Payment Type or Payment Mode", vbInformation, "Invalid Selection"
    On Error Resume Next
    comsetmode.SetFocus
    On Error GoTo 0
    Exit Sub
End If

If Trim(Mid(Trim(comsetmode.Text), 1, 5)) = "007" And Val(Trim(Mid(Trim(cmbpaymenttype.Text), 1, 3))) <> 9 Then
    MsgBox "Please Check the Payment Type or Payment Mode", vbInformation, "Invalid Selection"
    comsetmode.SetFocus
    Exit Sub
End If

End Sub

Private Sub creditamount_GotFocus()
xdecimal = 0
creditamount.Text = Val(lbldebitamount.caption) - creditamountcount
SendKeys "{home}+{end}"
End Sub

Private Sub creditamount_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then
   xdecimal = xdecimal + 1
 End If
If xdecimal = 2 Then
  Beep
  xdecimal = 1
 creditamount.SetFocus
 SendKeys "{Left}+{End}"
 SendKeys "{Delete}"
End If
If KeyAscii = 13 Or KeyAscii = 46 Or KeyAscii >= 48 And KeyAscii <= 57 Then
 c = KeyAscii
 
 
 Else
   If KeyAscii = 46 And xdecimal > 1 Then
       xdecimal = 0
   End If
 If creditamount.Text <> " " Then
        xdecimal = 0

  SendKeys "{End}+{Home}"
  SendKeys "{Delete}"
  creditamount.SetFocus

 End If
End If
  
If KeyAscii = 13 Or KeyAscii = 46 Or KeyAscii >= 48 And KeyAscii <= 57 Then
 strcheck = KeyAscii
 Else
  
  SendKeys "{End}+{Home}"
  SendKeys "{Delete}"
 creditamount.Text = ""
  Beep
End If

'cheking for the dot key

If KeyAscii = 13 Then
    cmdAdd_Click
End If

End Sub

Private Sub creditamount_LostFocus()
xdecimal = 0
End Sub

Private Sub delete_Click()
creditamountcount = creditamountcount - Val(ListView1.SelectedItem.SubItems(2))

deleteindex = ListView1.SelectedItem.Index
ListView1.ListItems.Remove (deleteindex)
'cmdadd.Enabled = True
creditcount = creditcount - 1
End Sub

Private Sub Form_Activate()
lbldate.caption = Format(Date, "dd/mm/yyyy")
dtpinvoicedate.Value = Format(Date, "dd/mm/yyyy")
End Sub

Private Sub Form_Load()
Me.Height = 6195

Timer2.Interval = 1
Label37.caption = Time
Mainform.sbStatusBar.Panels(1).Text = "Status : Receipt Voucher ..."
constring = "dsN=fINANCE;UID=SA;PWD=;"

CON1.Open "dsN=fINANCE;UID=SA;PWD=;"

recpay.Open "Select * from paymode where newcode is not null", CON1, adOpenDynamic, adLockOptimistic, adCmdText

recmod.Open "Select * from setmode", CON1, adOpenDynamic, adLockOptimistic, adCmdText ' the setmod table connection

reccurrency.Open "select * from currencytable where detail = 'Cash' and status='default'", CON1, adOpenKeyset, adLockOptimistic
lblcurrency.caption = reccurrency!currency
reccurrency.close
reccurrency.Open "Select * from currencytable", CON1, adOpenKeyset, adLockOptimistic

Dim recsales As New ADODB.Recordset
recsales.Open "Select * from salesman order by salesmancode", CON1, adOpenKeyset, adLockOptimistic

 Set cn = New ADODB.Connection
 Set recacc = New ADODB.Recordset
 xtable = "Select * from FinanceMaster where active <> '0' order by AccountCode"
 sqltable = True
 xClass.GetTables recacc, cn, xtable, constring, sqltable

'the receipt voucher number table
recnum.Open "select * from reNumber", CON1, adOpenDynamic, adLockOptimistic, adCmdText

' this is for ormaster
recor.Open "select * from vouchers", CON1, adOpenDynamic, adLockOptimistic, adCmdText

'this is for temperary table
rectemp.Open "Select * from tempinvoice", CON1, adOpenDynamic, adLockOptimistic, adCmdText

'this is for temporary2 table
recte.Open "Select * from tempinvoice2 order by autonumber", CON1, adOpenDynamic, adLockOptimistic, adCmdText

'this is for tempagaintsinvoice table
recteinvoice12.Open "Select * from tempagaintsinvoice", CON1, adOpenDynamic, adLockOptimistic, adCmdText

'check for the month and year change if we have to reset the series number.
recset.Open "Select getdate() as currentmoyr", CON1, adOpenKeyset, adLockOptimistic
Dim checkjournaldate As Date
checkjournaldate = Left(Trim(recset!CurrentMoYr), 10)
recset.close
recset.Open "Select * from setup", CON1, adOpenKeyset, adLockOptimistic
recset.MoveFirst
If Trim(recset!CurrentMoYr) <> Format(checkjournaldate, "mmyy") Then
    recset!CurrentMoYr = Format(checkjournaldate, "mmyy")
    recset!nextjn = "0001"
    recset.Update
End If

'end checking the journal no

'this is for add the listview
Dim conlist As New ADODB.Connection
xtable = "Select * from vouchers where deleted = '0' and post ='no' and svoucher='Collections' order by receiptno"
sqltable = True
xClass.GetTables reclist, conlist, xtable, constring, sqltable
    
l2 = 0
If reclist.BOF = False Then
    reclist.MoveFirst
    ListView2.ListItems.clear
        While reclist.EOF = False
            With ListView2
                l2 = l2 + 1
                .ListItems.Add , , Trim(reclist!receiptno)
                .ListItems(l2).ListSubItems.Add , , Trim(reclist!receiptdate)
                .ListItems(l2).ListSubItems.Add , , Trim(reclist!custname)
                .ListItems(l2).ListSubItems.Add , , " " & reclist!paymode 'IsNull(Trim(reclist!paymode), " ", Trim(reclist!paymode))
                .ListItems(l2).ListSubItems.Add , , " " & Trim(reclist!payopt)
                .ListItems(l2).ListSubItems.Add , , Format(Trim(reclist!receiptamount), "############0.#0")
            End With
        reclist.MoveNext
        DoEvents
        Wend
End If

Dim con1111 As New ADODB.Connection
con1111.Mode = adModeShareDenyNone
con1111.Open "Dsn=anufoxpro;uid=sa;pwd=;"

Dim constring1 As String
constring1 = "Dsn=anufoxpro;uid=sa;pwd=;"
xtable = "select cust_code,arabicname,first_name,mid_name,last_name from Marcusfl order by cust_code"
sqltable = True
xClass.GetTables recmarxx, cnmar, xtable, constring1, sqltable

recmarxx.MoveFirst
comreceivedfrom.clear
comfromname.clear
Set Adodc1.Recordset = recmarxx

While Me.Adodc1.Recordset.EOF = False
If Mid(Trim(Adodc1.Recordset!cust_code), 1, 1) = "O" Then
     
comreceivedfrom.AddItem Adodc1.Recordset!cust_code
name1 = ""
        If Trim(Adodc1.Recordset!arabicname) <> "" Then
            name1 = Trim(Adodc1.Recordset!arabicname) & " \ "
        End If
        If Trim(Adodc1.Recordset!first_name) <> "" Then
            name1 = name1 & Trim(Adodc1.Recordset!first_name) & " "
        End If
        If Trim(Adodc1.Recordset!mid_name) <> "" Then
            name1 = name1 & Trim(Adodc1.Recordset!mid_name) & " "
        End If
        If Trim(Adodc1.Recordset!last_name) <> "" Then
            name1 = name1 & Trim(Adodc1.Recordset!last_name)
        End If
        'this is for arabic name
    comfromname.AddItem name1
End If
    Me.Adodc1.Recordset.MoveNext
Wend
' end received from

'this is paymode
recpay.MoveFirst
cmbpaymenttype.clear
cmbpaymenttype.AddItem " "
While recpay.EOF = False
   If Val(Trim(Mid(Trim(recpay!paycode), 1, 3))) <> 9 Then
        cmbpaymenttype.AddItem recpay!newcode & "  " & recpay!neweng
   End If
    recpay.MoveNext
Wend
'end paymode

'for the salesman
comsalesman.clear
comsalesman.AddItem "  "
While recsales.EOF = False
    comsalesman.AddItem recsales!salesmancode & " " & recsales!salesmanname
    recsales.MoveNext
Wend
recsales.close
'for the setmode
recmod.MoveFirst
comsetmode.clear
comsetmode.AddItem " "
While recmod.EOF = False
    If Val(Trim(Mid(Trim(recmod!setcode), 1, 3))) <> 2 And Val(Trim(Mid(Trim(recmod!setcode), 1, 3))) <> 5 And Val(Trim(Mid(Trim(recmod!setcode), 1, 3))) <> 6 And Val(Trim(Mid(Trim(comsetmode.Text), 1, 3))) <= 7 Then
        comsetmode.AddItem recmod!setcode & "     " & recmod!setmode
    End If
    recmod.MoveNext
Wend
cmdsave.Enabled = True
recor.Requery

recacc.MoveFirst
combdebitaccnum.clear
comcreditaccountnumber.clear
cname.clear
ccname.clear
While recacc.EOF = False
    If Mid(Trim(recacc!AccountCode), 1, 5) = "11101" Or Mid(Trim(recacc!AccountCode), 1, 5) = "11102" Or Mid(Trim(recacc!AccountCode), 1, 5) = "11105" Or Mid(Trim(recacc!AccountCode), 1, 5) = "11106" Or Mid(Trim(recacc!AccountCode), 1, 5) = "11107" Then
        combdebitaccnum.AddItem Trim(recacc!AccountCode)
        cname.AddItem Trim(recacc!accountnameeng) & "\" & Trim(recacc!accountnamearab)
    End If
    On Error Resume Next
       comcreditaccountnumber.AddItem Trim(recacc!AccountCode)
       ccname.AddItem Trim(recacc!accountnameeng) & "\" & Trim(recacc!accountnamearab)
    On Error GoTo 0
    recacc.MoveNext
Wend
listcurrency.ListItems.clear
i = 1
While reccurrency.EOF = False
    listcurrency.ListItems.Add , , reccurrency!currency
    listcurrency.ListItems(i).ListSubItems.Add , , Format(reccurrency!rate, "###0.###0")
    listcurrency.ListItems(i).ListSubItems.Add , , reccurrency!latestupdate
    i = i + 1
    reccurrency.MoveNext
Wend
Mainform.sbStatusBar.Panels(1).Text = "Status : Ready."
lbldate.caption = Format(Date, "dd/mm/yyyy")
stopstop1 = 0
Timer2.Interval = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
cmdcancel_Click
On Error Resume Next
CON1.close
CON1.close
conmar.close
conlist.close
On Error GoTo 0
End Sub

Private Sub lblcurrency_Change()
'Dim recfacc As New ADODB.Recordset
'If Trim(cmbpaymenttype.Text) <> "" Then
'    recfacc.Open "Select * from currencytable where currency = '" & Trim(lblcurrency.caption) & "'", CON1, adOpenKeyset, adLockOptimistic
'        combdebitaccnum.Text = Trim(recfacc!accountnumber)
'        combdebitaccnum_Click
'    recfacc.close
'    listcurrency.Visible = False
'Else
'    combdebitaccnum.Text = ""
'    cname.Text = ""
'End If
End Sub

Private Sub lblcurrency_Click()
If listcurrency.Visible = False Then
    listcurrency.Visible = True
    listcurrency.Top = 1440
Else
    listcurrency.Visible = False
End If
End Sub

Private Sub listcurrency_Click()

If listcurrency.ListItems.Count > 0 Then
    lblcurrency.caption = Trim(listcurrency.SelectedItem.Text)
    
    If Val(Trim(txtdebitamt.Text)) > 0 And (Trim(cmbpaymenttype.Text) = "" Or Trim(Mid(Trim(cmbpaymenttype.Text), 1, 3)) = "04" Or Trim(Mid(Trim(cmbpaymenttype.Text), 1, 3)) = "05") Then
        Dim findrate As New ADODB.Recordset
        findrate.Open "Select * from currencytable where currency = '" & Trim(lblcurrency.caption) & "'", CON1, adOpenKeyset, adLockOptimistic
        lbldebitamount.caption = Format(Val(Trim(txtdebitamt.Text)) * Val(findrate!rate), "############0.#0")
        findrate.close
    Else
        lbldebitamount.caption = Format(Val(Trim(txtdebitamt.Text)), "############0.#0")
    End If
    txtdebitamt.SetFocus
End If
    listcurrency.Visible = False
End Sub

Private Sub listcurrency_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
listcurrency.SortKey = ColumnHeader.Index - 1
Me.listcurrency.Sorted = True
End Sub



'Private Sub listcurrency_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'listcurrency.SetFocus
'End Sub

Private Sub ListView1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
    delete_Click
End If
End Sub
Public Sub xxxyyy()
cmdcancel_Click
End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
If ListView1.ListItems.Count > 0 And Button = vbRightButton Then
    PopupMenu file
End If
End Sub

Private Sub ListView2_DblClick()
If CmdEdit.Enabled = True Then
    CMDEDIT_Click
    txtreceiptnumber.Text = ListView2.SelectedItem.Text
    Me.SSTab1.SetFocus
    SendKeys "{Left}"
    txtreceiptnumber_KeyPress 13
End If

End Sub

Private Sub Timer1_Timer()
Dim recreceipt As New ADODB.Recordset
Dim recpayment As New ADODB.Recordset
Dim recopening As New ADODB.Recordset
Dim recclosing As New ADODB.Recordset

recopening.Open "Select * from Balance order by balancedate", CON1, adOpenKeyset, adLockOptimistic
recreceipt.Open "SELECT SUM(debitamount) AS receipttotal From vouchers WHERE (deleted = '0') and okprint = '0' AND svoucher = 'Collections' AND (POST = 'no') and debitamount > 0 AND (SUBSTRING(PAYMODE, 1, 2) = '01') and receiptdate =" & "'" & Format(Date, "mm/dd/yyyy") & "'", CON1, adOpenKeyset, adLockOptimistic
recpayment.Open "SELECT SUM(creditamount) AS paymenttotal From vouchers WHERE (deleted = '0') and okprint = '0' AND svoucher = 'Disbursments' AND (POST = 'no') AND creditamount > 0 AND (SUBSTRING(PAYMODE, 1, 2) = '01') AND receiptDate =" & "'" & Format(Date, "mm/dd/yyyy") & "'", CON1, adOpenKeyset, adLockOptimistic

op.caption = Format(recopening!openingbalance, "##############0.#0")

If recreceipt!receipttotal <> "" Then
re.caption = Format(recreceipt!receipttotal, "##############0.#0")
Else
re.caption = "0.00"
End If

If recpayment!paymenttotal <> "" Then
pa.caption = Format(recpayment!paymenttotal, "##############0.#0")
Else
pa.caption = "0.00"
End If
cl.caption = Format(Val(op.caption) + Val(re.caption) - Val(pa.caption), "############0.#0")
End Sub

Private Sub timaccountnumber_Timer()
Dim recfacc As New ADODB.Recordset
If Trim(cmbpaymenttype.Text) <> "" Then
    recfacc.Open "Select * from currencytable where currency = '" & Trim(lblcurrency.caption) & "'", CON1, adOpenKeyset, adLockOptimistic
        combdebitaccnum.Text = Trim(recfacc!accountnumber)
        combdebitaccnum_Click
    recfacc.close
Else
    combdebitaccnum.Text = ""
    cname.Text = ""
End If

End Sub

Private Sub Timer2_Timer()
Label51.caption = Time
End Sub

Private Sub txtcheckdate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    comsetmode.SetFocus
End If
End Sub

Private Sub txtchecknumber_GotFocus()
SendKeys "{End}+{Home}"
End Sub

Private Sub txtchecknumber_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtcheckdate.Enabled = True Then
        txtcheckdate.SetFocus
    Else
        comsetmode.SetFocus
    End If
End If
End Sub

Private Sub txtdebitamt_GotFocus()
xdecimal = 0
'SendKeys "{End}+{Home}"
End Sub

Private Sub txtdebitamt_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then
   xdecimal = xdecimal + 1
 End If
If xdecimal = 2 Then
  Beep
  xdecimal = 1
 txtdebitamt.SetFocus
 SendKeys "{Left}+{End}"
 SendKeys "{Delete}"
End If
If KeyAscii = 13 Or KeyAscii = 46 Or KeyAscii >= 48 And KeyAscii <= 57 Then
 c = KeyAscii
 
 
 Else
   If KeyAscii = 46 And xdecimal > 1 Then
       xdecimal = 0
   End If
 If txtdebitamt.Text <> " " Then
        xdecimal = 0

  SendKeys "{End}+{Home}"
  SendKeys "{Delete}"
  txtdebitamt.SetFocus

 End If
End If
  
If KeyAscii = 13 Or KeyAscii = 46 Or KeyAscii >= 48 And KeyAscii <= 57 Then
 strcheck = KeyAscii
 Else
  
  SendKeys "{End}+{Home}"
  SendKeys "{Delete}"
 txtdebitamt.Text = ""
  Beep
End If
If KeyAscii = 13 Then
    cmbpaymenttype.SetFocus
End If
'cheking for the dot key
End Sub

Private Sub txtdebitamt_LostFocus()
xdecimal = 0
'Dim AmtInFigure As Currency
'Dim AmtInwords As String
'Dim xAmtinWords As HabitatClass
'Set xAmtinWords = New HabitatClass
'
'On Error Resume Next
'AmtInFigure = IIf(IsNull(txtdebitamt.Text) = True, 0, txtdebitamt.Text)
'On Error GoTo 0
'xAmtinWords.AmountInWords AmtInFigure, AmtInwords
'If AmtInwords <> "" Then
'  txtamountinword = LTrim(AmtInwords)
'  Else
'  txtamountinword = "Can't create amount in words because amount in figure exceeding 21Million" & vbCrLf & "  ÇáãÈáÛ ÃßËÑ ãä 21 ãáíæä áÇ íãßä ÊÍÏíË ÇáÍÞá"
'End If

If Val(Trim(txtdebitamt.Text)) > 0 Then
    Dim findrate As New ADODB.Recordset
    findrate.Open "Select * from currencytable where currency = '" & Trim(lblcurrency.caption) & "'", CON1, adOpenKeyset, adLockOptimistic
    lbldebitamount.caption = Format(Val(Trim(txtdebitamt.Text)) * Val(findrate!rate), "############0.#0")
    findrate.close
End If
'lbldebitamount.Caption = Format(Trim(txtdebitamt.Text), "############0.#0")
End Sub

Private Sub txtdetails_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    comcreditaccountnumber.SetFocus
End If
End Sub



Private Sub txtinvoicenum_GotFocus()
txtinvoicenum.Text = ""
End Sub

Private Sub txtinvoicenum_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    comsalesman.SetFocus
End If
End Sub

Private Sub txtreceiptnumber_KeyPress(KeyAscii As Integer)
sqltable = True
Dim connectionstring As String
Dim recorrr As New ADODB.Recordset

Dim conforsearch As New ADODB.Connection
connectionstring = "dsn=finance;uid=sa;pwd=;"
xtable = "select * from vouchers where svoucher = 'Collections' and receiptno = '" & Trim(txtreceiptnumber.Text) & "'"
myclass.GetTables recorrr, conforsearch, xtable, connectionstring, sqltable

If KeyAscii = 13 Then
    If recorrr.BOF = True Then
        recorrr.close
        conforsearch.close
        MsgBox "This Number is Not Valid", vbInformation, "Roll Back"
        cmdcancel_Click
        Exit Sub
   Else
    If Trim(recorrr!Post) = "Yes" Then
        '********* this is for the display procedure
        txtreceiptnumber.Text = recorrr!receiptno
        lbldate.caption = recorrr!receiptdate
        comreceivedfrom.Text = recorrr!custno
        comfromname.Text = recorrr!custname
        tookcredit = 0
        If IsNull(recorrr!DebitAmount) Then
            tookcredit = 0
        Else
            tookcredit = Trim(recorrr!DebitAmount)
        End If
        
        If tookcredit = 0 Then
            If IsNull(recorrr!checkreceipt) Then
                tookcredit = 0
            Else
                tookcredit = Trim(recorrr!checkreceipt)
            End If
        End If
        
        
        If tookcredit = 0 Then
                tookcredit = Trim(recorrr!doller)
        End If
        
        txtdebitamt.Text = tookcredit
        lbldebitamount.caption = Format(tookcredit, "#########0.#0")
        tookcredit = 0
        comsetmode.Text = recorrr!payopt
        txtinvoicenum.Text = recorrr!optref & " "
        cmbpaymenttype.Text = recorrr!paymode
        If IsNull(recorrr!chkdue) = False Then
        txtcheckdate.Value = recorrr!chkdue
        txtchecknumber.Text = recorrr!moderef
        End If
        txttempreceipt.Text = IIf(IsNull(recorrr!tempinvoice), " ", recorrr!tempinvoice)
        takejournal = IIf(IsNull(recorrr!JournalNumber), " ", recorrr!JournalNumber) ' the journal
        On Error Resume Next
         txtdetails.Text = recorrr!remarks
        '*** this is the display procedure
        Frame1.Enabled = True
        comreceivedfrom.Enabled = False
        comfromname.Enabled = False
        txtdetails.Enabled = False
        txtchecknumber.Enabled = False
        txtdebitamt.Enabled = False
        comsetmode.Enabled = True
        cmdinvoice.Enabled = True
        txtinvoicenum.Enabled = False
        cmbpaymenttype.Enabled = True
        txtchecknumber.Enabled = True
        txtcheckdate.Enabled = True
        Exit Sub
    End If
        
   End If
   If Trim(recorrr!ncount) = "101" And Trim(recorrr!Post) = "no" Then
        If MsgBox("Are Your Sure Your Want to Analyse Again ?", vbYesNo + vbQuestion + vbDefaultButton2, "Conformation") = vbYes Then
            'this cording is exist in saving for accountant
        ElseIf Trim(Mid(Trim(recorrr!payopt), 1, 4)) = "001" And Trim(Mid(Trim(recorrr!paymode), 1, 3)) = "01" Then
            If MsgBox("Do You Want to Pay for Invoices", vbYesNo + vbQuestion + vbDefaultButton2, "Conformation") = vbYes Then
                '********* this is for the display procedure
                txtreceiptnumber.Text = recorrr!receiptno
                lbldate.caption = recorrr!receiptdate
                comreceivedfrom.Text = recorrr!custno
                comfromname.Text = recorrr!custname
                tookcredit = 0
                If IsNull(recorrr!DebitAmount) Then
                    tookcredit = 0
                Else
                    tookcredit = Trim(recorrr!DebitAmount)
                End If
                
                If tookcredit = 0 Then
                    If IsNull(recorrr!checkreceipt) Then
                        tookcredit = 0
                    Else
                        tookcredit = Trim(recorrr!checkreceipt)
                    End If
                End If
                
                If tookcredit = 0 Then
                        tookcredit = Trim(recorrr!doller)
                        Rate1 = 0
                        Do While Rate1 <= 0
                            Rate1 = Val(InputBox("Please Enter The Dollar Rate", "Dollar Rate", "5.50"))
                            If Rate1 <= 0 Then
                                MsgBox "Plese Enter the Correct Rate to Convert to L.E", vbInformation, "Invalid Rate"
                            End If
                        Loop
                        tookcredit = tookcredit * Rate1
                End If
        
                txtdebitamt.Text = tookcredit
                lbldebitamount.caption = Format(tookcredit, "#########0.#0")
                tookcredit = 0
                comsetmode.Text = recorrr!payopt
                txtinvoicenum.Text = recorrr!optref & " "
                cmbpaymenttype.Text = recorrr!paymode
                If IsNull(recorrr!chkdue) = False Then
                txtcheckdate.Value = recorrr!chkdue
                txtchecknumber.Text = recorrr!moderef
                End If
                txttempreceipt.Text = IIf(IsNull(recorrr!tempinvoice), " ", recorrr!tempinvoice)
                takejournal = recorrr!JournalNumber ' the journal
                On Error Resume Next
                 comprepared.Text = recorrr!cashier
                 comchecked.Text = recorrr!empcheck
                 comapproved.Text = recorrr!empapp
                 txtdetails.Text = recorrr!remarks
                '*** this is the display procedure
                Frame1.Enabled = True
                comreceivedfrom.Enabled = False
                comfromname.Enabled = False
                txtdetails.Enabled = False
                txtchecknumber.Enabled = False
                txtdebitamt.Enabled = False
                comsetmode.Enabled = True
                cmdinvoice.Enabled = True
                txtinvoicenum.Enabled = False
                cmbpaymenttype.Enabled = True
                txtchecknumber.Enabled = True
                txtcheckdate.Enabled = True
                comprepared.Enabled = False
                comchecked.Enabled = False
                comapproved.Enabled = False
                txtattachment.Enabled = False
                Exit Sub
            Else
            recorrr.close
            conforsearch.close
            cmdcancel_Click
            Exit Sub
            End If
        Else
            recorrr.close
            conforsearch.close
            cmdcancel_Click
            Exit Sub
        End If
   End If
                Frame1.Enabled = True
                comreceivedfrom.Enabled = False
                comfromname.Enabled = False
                txtdetails.Enabled = False
                txtchecknumber.Enabled = False
                txtdebitamt.Enabled = False
                comsetmode.Enabled = True
                cmdinvoice.Enabled = True
                txtinvoicenum.Enabled = False
                cmbpaymenttype.Enabled = True
                txtchecknumber.Enabled = True
                txtcheckdate.Enabled = True
                comprepared.Enabled = False
                comchecked.Enabled = False
                comapproved.Enabled = False
                '********* this is for the display procedure
                txtreceiptnumber.Text = recorrr!receiptno
                lbldate.caption = recorrr!receiptdate
                comreceivedfrom.Text = recorrr!custno
                comfromname.Text = recorrr!custname
                tookcredit = 0
                If IsNull(recorrr!DebitAmount) Then
                    tookcredit = 0
                Else
                    tookcredit = Trim(recorrr!DebitAmount)
                End If
                
                If tookcredit = 0 Then
                    If IsNull(recorrr!checkreceipt) Then
                        tookcredit = 0
                    Else
                        tookcredit = Trim(recorrr!checkreceipt)
                    End If
                End If
                
                If tookcredit = 0 Then
                        tookcredit = Trim(recorrr!doller)
                        Rate1 = 0
                        Do While Rate1 <= 0
                            Rate1 = Val(InputBox("Please Enter The Dollar Rate", "Dollar Rate", "5.50"))
                            If Rate1 <= 0 Then
                                MsgBox "Plese Enter the Correct Rate to Convert to L.E", vbInformation, "Invalid Rate"
                            End If
                        Loop
                        tookcredit = tookcredit * Rate1
                End If

                
                txtdebitamt.Text = tookcredit
                lbldebitamount.caption = Format(tookcredit, "#########0.#0")
                tookcredit = 0
                comsetmode.Text = recorrr!payopt
                txtinvoicenum.Text = IIf(IsNull(recorrr!optref), " ", recorrr!optref)
                cmbpaymenttype.Text = recorrr!paymode
                If IsNull(recorrr!chkdue) = False Then
                txtcheckdate.Value = recorrr!chkdue
                txtchecknumber.Text = IIf(IsNull(recorrr!moderef), "", recorrr!moderef)
                End If
                txttempreceipt.Text = IIf(IsNull(recorrr!tempinvoice), " ", recorrr!tempinvoice)
                takejournal = recorrr!JournalNumber ' the journal
                On Error Resume Next
                 comprepared.Text = recorrr!cashier
                 comchecked.Text = recorrr!empcheck
                 comapproved.Text = recorrr!empapp
                 txtdetails.Text = recorrr!remarks
                '*** this is the display procedure
                'Frame2.Enabled = True
                'cname.Enabled = True
                ccname.Enabled = True
                'combdebitaccnum.Enabled = True
                comcreditaccountnumber.Enabled = True
        
        ' this is to identify the accountnumber for the client code
        Dim conmarcus As New ADODB.Connection
        conmarcus.Mode = adModeShareDenyNone
        conmarcus.Open "Dsn=anufoxpro;uid=sa;pwd=;"
        Dim recmacmac12 As New ADODB.Recordset
        'this is for client code when main code is null
        
        recmacmac12.Open "Select * from marcusfl where LEFT(TRIM(cust_code), 1) = 'O' AND cust_code = '" & Trim(comreceivedfrom.Text) & "'", conmarcus, adOpenKeyset, adLockOptimistic
                    If recmacmac12.BOF = False Then
                        If Trim(recmacmac12!mcustcode) <> "" Then ' if there is main code then
                            takemaincode = Trim(recmacmac12!mcustcode)
                        Else
                            If Trim(recmacmac12!acctNo) <> "" Then
                                comcreditaccountnumber.Text = recmacmac12!acctNo
                                findaccountnumber = 1
                            End If
                        End If
                    End If
       
       On Error Resume Next
       recmacmac12.close
       On Error GoTo 0
        
        'this is for the main client code
        recmacmac12.Open "Select * from marcusfl where LEFT(TRIM(cust_code), 1) = 'O' AND cust_code = '" & Trim(takemaincode) & "'", conmarcus, adOpenKeyset, adLockOptimistic
            If recmacmac12.BOF = False Then
                comcreditaccountnumber.Text = recmacmac12!acctNo
                findaccountnumber = 1
                recmacmac12.close
            End If
            
       On Error Resume Next
       recmacmac12.close
       On Error GoTo 0
       
       'if mainclient code and accountcode also null then this will check
       If findaccountnumber <> 1 Then
            recmacmac12.Open "Select * from marcusfl where LEFT(TRIM(cust_code), 1) = 'O' AND cust_code = " & "'" & Trim(comreceivedfrom.Text) & "'", conmarcus, adOpenKeyset, adLockOptimistic
              If recmacmac12.BOF = False Then
                If UCase(Trim(recmacmac12!Grp)) = "OS" Then
                    comcreditaccountnumber.Text = "111041001001"
                Else
                    comcreditaccountnumber.Text = "111041102001"
                End If
              End If
            findaccountnumber = 2
        End If
        On Error Resume Next
        recmacmac12.close
        On Error GoTo 0
        conmarcus.close
        ' end client code
                    On Error GoTo 0
                    comverified.Enabled = True
                    comenterd.Enabled = True
                    txtdebitamt.Enabled = True
                    ListView1.Enabled = True
                    creditamount.Enabled = True
                    'cmdadd.Enabled = True
                    combdebitaccnum.Text = "111010101001"
                    combdebitaccnum.SetFocus
                    Exit Sub
End If
End Sub

Public Sub prcdisplay()
txtreceiptnumber.Text = recor!receiptno
lbldate.caption = recor!receiptdate
comreceivedfrom.Text = recor!custno
comfromname.Text = recor!custname
tookcredit = 0
If IsNull(recor!DebitAmount) Then
    tookcredit = 0
Else
    tookcredit = Trim(recor!DebitAmount)
End If

If tookcredit = 0 Then
    If IsNull(recor!checkreceipt) Then
        tookcredit = 0
    Else
        tookcredit = Trim(recor!checkreceipt)
    End If
End If

If tookcredit = 0 Then
        tookcredit = Trim(recor!doller)
End If
txtdebitamt.Text = tookcredit
lbldebitamount.caption = Format(tookcredit, "#########0.#0")
tookcredit = 0
comsetmode.Text = recor!payopt
txtinvoicenum.Text = recor!optref & " "
cmbpaymenttype.Text = recor!paymode
txtchecknumber.Text = recor!moderef
If IsNull(recor!chkdue) = False Then
txtcheckdate.Value = recor!chkdue
End If
txttempreceipt.Text = IIf(IsNull(recor!tempinvoice), " ", recor!tempinvoice)
takejournal = recor!JournalNumber ' the journal
On Error Resume Next
 comprepared.Text = recor!cashier
 comchecked.Text = recor!empcheck
 comapproved.Text = recor!empapp
 txtdetails.Text = recor!remarks
End Sub

Public Sub prcclear()

lbldate.caption = Format(Date, "mm/dd/yyyy")

comsetmode.ListIndex = 0
cmbpaymenttype.ListIndex = 0
comsalesman.ListIndex = 0
comreceivedfrom.Text = ""
comfromname.Text = ""

txtdebitamt.Text = ""
txtinvoicenum.Text = ""
txtchecknumber.Text = ""
txtcheckdate.Value = Null
txttempreceipt.Text = " "
txtdetails.Text = ""
txtallinvoice.Text = ""

combdebitaccnum.Text = ""
comcreditaccountnumber.Text = ""
ccname.Text = ""
cname.Text = ""
lbldebitamount.caption = ""
ListView1.ListItems.clear
creditcount = 0
creditamount = 0
creditamountcount = 0
'cmdadd.Enabled = True
End Sub

Private Sub txttempreceipt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        comreceivedfrom.SetFocus
End If
End Sub

Private Sub txttempreceipt_LostFocus()
    
If Trim(txttempreceipt.Text) <> "" Then
recchecktemp.Open "Select * from vouchers where tempinvoice =" & "'" & Trim(txttempreceipt.Text) & "'", CON1, adOpenKeyset, adLockOptimistic
    If recchecktemp.BOF = True Then
        recchecktemp.close
        'comreceivedfrom.SetFocus
    Else
        MsgBox "Please check Your Temporary Voucher Number is Repeating" & vbCrLf & "  åÐÇ ÇáÑÞã Êã ÇÏÎÇáå ãä ÞÈá ÇÎÊÑ ÑÞã ÇÎÑ", vbInformation, "Repeating Number"
        recchecktemp.close
        txttempreceipt.Text = ""
        txttempreceipt.SetFocus
        Exit Sub
    End If
'Else
'    comreceivedfrom.SetFocus
End If
End Sub



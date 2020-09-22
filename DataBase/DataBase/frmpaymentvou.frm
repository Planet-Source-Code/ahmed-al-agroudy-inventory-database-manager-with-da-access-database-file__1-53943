VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmpaymentvou 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payment Voucher  ÓäÏ ÇáÕÑÝ"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11415
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arabic Transparent"
      Size            =   9
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmpaymentvou.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   11415
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timaccountcode 
      Interval        =   500
      Left            =   4080
      Top             =   6240
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
      Left            =   360
      TabIndex        =   54
      Top             =   5760
      Width           =   855
   End
   Begin MSComctlLib.ListView listcurrency 
      Height          =   1455
      Left            =   6960
      TabIndex        =   53
      Top             =   5400
      Visible         =   0   'False
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   2566
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
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
   Begin VB.Timer Timer2 
      Left            =   2280
      Top             =   5640
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   9128
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
      TabPicture(0)   =   "frmpaymentvou.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "framebutton"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "langopt"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "List of View ÞÇÆãÉ "
      TabPicture(1)   =   "frmpaymentvou.frx":045E
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
         Picture         =   "frmpaymentvou.frx":047A
         TabIndex        =   51
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   2385
      End
      Begin VB.Frame framebutton 
         Height          =   2655
         Left            =   -64920
         TabIndex        =   43
         Top             =   2280
         Width           =   1095
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
            TabIndex        =   48
            Top             =   1200
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
            TabIndex        =   47
            Top             =   720
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
            TabIndex        =   46
            Top             =   1680
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
            TabIndex        =   45
            Top             =   240
            Width           =   855
         End
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
            TabIndex        =   44
            Top             =   2160
            Width           =   855
         End
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   4575
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   8070
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
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Payment To"
            Object.Width           =   7937
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Pay-Type"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Pay-Option"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Amount"
            Object.Width           =   2540
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
         Height          =   4695
         Left            =   -74880
         TabIndex        =   1
         Top             =   360
         Width           =   11175
         Begin VB.ComboBox txtattachment 
            Height          =   330
            Left            =   1200
            Style           =   1  'Simple Combo
            TabIndex        =   56
            Top             =   1560
            Width           =   9015
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
            TabIndex        =   50
            Top             =   840
            Width           =   2655
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
            TabIndex        =   49
            Top             =   1200
            Width           =   2655
         End
         Begin VB.ComboBox cname 
            Enabled         =   0   'False
            Height          =   330
            Left            =   3120
            Sorted          =   -1  'True
            TabIndex        =   29
            Top             =   2160
            Width           =   4935
         End
         Begin VB.ComboBox comcreditaccountnumber 
            Height          =   330
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   28
            Top             =   2760
            Width           =   2895
         End
         Begin VB.ComboBox ccname 
            Height          =   330
            Left            =   3120
            Sorted          =   -1  'True
            TabIndex        =   27
            Top             =   2760
            Width           =   4935
         End
         Begin VB.ComboBox combdebitaccnum 
            Enabled         =   0   'False
            Height          =   330
            ItemData        =   "frmpaymentvou.frx":1344
            Left            =   120
            List            =   "frmpaymentvou.frx":1346
            Sorted          =   -1  'True
            TabIndex        =   26
            Top             =   2160
            Width           =   2895
         End
         Begin VB.TextBox creditamount 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   8160
            TabIndex        =   25
            Top             =   2760
            Width           =   1575
         End
         Begin VB.TextBox txtdebitamt 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   8950
            TabIndex        =   4
            Top             =   480
            Width           =   1250
         End
         Begin VB.ComboBox comreceivedfrom 
            Height          =   330
            Left            =   1200
            Sorted          =   -1  'True
            TabIndex        =   3
            Top             =   480
            Width           =   5775
         End
         Begin VB.TextBox txtchecknumber 
            Enabled         =   0   'False
            Height          =   315
            Left            =   5520
            MaxLength       =   15
            TabIndex        =   2
            Top             =   840
            Width           =   1455
         End
         Begin MSComCtl2.DTPicker txtcheckdate 
            Height          =   345
            Left            =   8400
            TabIndex        =   5
            Top             =   840
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   609
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   63176705
            CurrentDate     =   37548
         End
         Begin MSMask.MaskEdBox txtinvoicenum 
            Height          =   315
            Left            =   5520
            TabIndex        =   6
            Top             =   1200
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
            MaxLength       =   14
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arabic Transparent"
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
            Height          =   1500
            Left            =   120
            TabIndex        =   52
            Top             =   3120
            Width           =   9615
            _ExtentX        =   16960
            _ExtentY        =   2646
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
         Begin MSMask.MaskEdBox txtreceiptnumber 
            Height          =   315
            Left            =   1200
            TabIndex        =   57
            Top             =   120
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
            TabIndex        =   62
            Top             =   2520
            Width           =   345
         End
         Begin VB.Label Label11 
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
            TabIndex        =   61
            Top             =   1920
            Width           =   345
         End
         Begin VB.Label Label6 
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
            TabIndex        =   60
            Top             =   480
            Width           =   345
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
            Left            =   3960
            TabIndex        =   59
            Top             =   840
            Width           =   720
         End
         Begin VB.Label Label5 
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
            Left            =   4320
            TabIndex        =   58
            Top             =   1200
            Width           =   315
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
            MouseIcon       =   "frmpaymentvou.frx":1348
            MousePointer    =   99  'Custom
            TabIndex        =   55
            Top             =   480
            Width           =   570
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Credit Account "
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
            TabIndex        =   40
            Top             =   1920
            Width           =   1695
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
            Left            =   8160
            TabIndex        =   39
            Top             =   2160
            Width           =   1575
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
            Left            =   8160
            TabIndex        =   38
            Top             =   1920
            Width           =   540
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
            Left            =   3240
            TabIndex        =   37
            Top             =   1920
            Width           =   1455
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Debit Account "
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
            TabIndex        =   36
            Top             =   2520
            Width           =   1065
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
            Left            =   8160
            TabIndex        =   35
            Top             =   2520
            Width           =   540
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
            Left            =   3240
            TabIndex        =   34
            Top             =   2520
            Width           =   1455
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            Caption         =   "ÑÞã ÍÓÇÈ ÇáãÏíä "
            BeginProperty Font 
               Name            =   "Arabic Transparent"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1800
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   2520
            Width           =   1140
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "ÇÓã ÇáÍÓÇÈ "
            BeginProperty Font 
               Name            =   "Arabic Transparent"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   6840
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   2520
            Width           =   810
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "ÑÞã  ÍÓÜÇÈ ÇáÏÇÆä"
            BeginProperty Font 
               Name            =   "Arabic Transparent"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1800
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   1920
            Width           =   1230
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            Caption         =   "ÇÓã ÇáÍÓÇÈ "
            BeginProperty Font 
               Name            =   "Arabic Transparent"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   6840
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   1920
            Width           =   810
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Payment # :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   120
            Width           =   1020
         End
         Begin VB.Label Label2 
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
            Height          =   255
            Left            =   4200
            TabIndex        =   23
            Top             =   120
            Width           =   1215
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
            Left            =   5640
            TabIndex        =   22
            Top             =   120
            Width           =   1200
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " ÑÞã ÇáÇÓÊáÇã "
            BeginProperty Font 
               Name            =   "Arabic Transparent"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   3000
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "ãÑÝÞÇÊ "
            BeginProperty Font 
               Name            =   "Arabic Transparent"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   10440
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   1560
            Width           =   495
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "ÊÇÑíÎ ÇáÔíß "
            BeginProperty Font 
               Name            =   "Arabic Transparent"
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
            TabIndex        =   19
            Top             =   840
            Width           =   810
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "ÑÞã ÇáÓíß "
            BeginProperty Font 
               Name            =   "Arabic Transparent"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   6960
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   840
            Width           =   675
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "ãÑÌÚ Èå "
            BeginProperty Font 
               Name            =   "Arabic Transparent"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   7080
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Top             =   1200
            Width           =   600
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ÏÝÚ Çáí"
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
            Left            =   7080
            TabIndex        =   16
            Top             =   480
            Width           =   555
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Date"
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
            Left            =   7800
            TabIndex        =   14
            Top             =   840
            Width           =   345
         End
         Begin VB.Label lblref 
            AutoSize        =   -1  'True
            Caption         =   "Ref. No"
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
            Left            =   4800
            TabIndex        =   13
            Top             =   1200
            Width           =   555
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Chq. No"
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
            Left            =   4800
            TabIndex        =   12
            Top             =   840
            Width           =   585
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Payment for"
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
            Left            =   120
            TabIndex        =   11
            Top             =   1200
            Width           =   840
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Payment Mode"
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
            Left            =   120
            TabIndex        =   10
            Top             =   840
            Width           =   1065
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Description"
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
            Left            =   120
            TabIndex        =   9
            Top             =   1560
            Width           =   795
         End
         Begin VB.Label Label7 
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
            Left            =   7800
            TabIndex        =   8
            Top             =   480
            Width           =   540
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Pay to"
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
            Left            =   120
            TabIndex        =   7
            Top             =   480
            Width           =   450
         End
      End
   End
   Begin VB.Label Label51 
      Caption         =   "Label51"
      Height          =   255
      Left            =   4920
      TabIndex        =   42
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Label Label33 
      Caption         =   "Label37"
      Height          =   255
      Left            =   2880
      TabIndex        =   41
      Top             =   5760
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
Attribute VB_Name = "frmpaymentvou"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim recor As New ADODB.Recordset
Dim creditcount As Integer        ' this is for add button
Dim creditamountcount As Currency ' thsi is for count the balance
Dim recnum As New ADODB.Recordset
Dim recacc As New ADODB.Recordset
Dim recemp As New ADODB.Recordset
Dim recmar As New ADODB.Recordset
Dim recche As New ADODB.Recordset
Dim recpayable As New ADODB.Recordset
Dim reclist As New ADODB.Recordset
Dim recset As New ADODB.Recordset
Dim CON1 As New ADODB.Connection
Dim conpay As New ADODB.Connection
Dim myclass As New HabitatClass
Dim okcontinue As Integer
Dim takejournal As String
Dim sqltable As Boolean
Dim xtable As String
Dim xdecimal As Integer
Dim recpay As New ADODB.Recordset ' this is for paymode
Dim recmod As New ADODB.Recordset
Public fromwho As String
Public wrongcount As Integer
Public bringinvoice As String 'this takes the
Public bankname As String
Public stopstop1 As Integer
Public checknumberanddate As String
Dim constring As String

Private Sub ccname_Click()
Dim recfindacc As New ADODB.Recordset
namenamenum = InStr(1, Trim(ccname.Text), "\", vbTextCompare)

If namenamenum > 0 Then
namename = Mid(Trim(ccname.Text), 1, (namenamenum - 1))
'namename = Trim(ccname.Text)
Else
namename = Trim(ccname.Text)
End If
recfindacc.Open "Select * from financemaster where active <> '0' and  accountnameeng = " & "'" & namename & "'", CON1, adOpenKeyset, adLockOptimistic
If recfindacc.BOF = False Then
    comcreditaccountnumber.Text = recfindacc!AccountCode
End If
 recfindacc.close
 'this is for the mother name
    getmothername nonumber, namename, cc
Me.caption = "Disbursement Voucher  " & cc
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
recfindacc.Open "Select * from financemaster where active <> '0' and  accountnameeng = " & "'" & namename & "'", CON1, adOpenKeyset, adLockOptimistic
If recfindacc.BOF = True Then
    recfindacc.close
    MsgBox "Please choose the Correct Account Number", vbInformation, "Invalid Account Number"
    ccname.SetFocus
    Exit Sub
Else
    comcreditaccountnumber.Text = recfindacc!AccountCode
End If
End Sub
Private Sub cmbpaymenttype_Click()
Dim reccurrency As New ADODB.Recordset 'this is for add currency
reccurrency.Open "select * from currencytable", CON1, adOpenKeyset, adLockOptimistic
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
'end add currency optionsIf Trim(cmbpaymenttype.Text) <> "" Then
    If Val(Trim(txtdebitamt.Text)) > 0 And (Trim(Mid(Trim(cmbpaymenttype.Text), 1, 3)) = "04" Or Trim(Mid(Trim(cmbpaymenttype.Text), 1, 3)) = "05") Then
        Dim findrate As New ADODB.Recordset
        findrate.Open "Select * from currencytable where currency = '" & Trim(lblcurrency.caption) & "'", CON1, adOpenKeyset, adLockOptimistic
        lbldebitamount.caption = Format(Val(Trim(txtdebitamt.Text)) * Val(findrate!rate), "############0.#0")
        findrate.close
    Else
        lbldebitamount.caption = Format(Trim(txtdebitamt.Text), "############0.#0")
    End If



If Val(Mid(Trim(cmbpaymenttype.Text), 1, 2)) <> 3 And Val(Mid(Trim(cmbpaymenttype.Text), 1, 2)) <> 4 Then
    txtchecknumber.Enabled = False
    txtcheckdate.Enabled = False
Else
    'txtchecknumber.Enabled = True
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
    listcurrency.Top = 1200
   
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
Else
    'lbldebitamount.Caption = ""
End If

'If Trim(Mid(Trim(cmbpaymenttype.Text), 1, 2)) = "10" And Trim(comreceivedfrom.Text) <> "O900002" Then
'    MsgBox "Please Check Your Client Code or Payment Mode", vbInformation, "Invalid Choice"
'    cmbpaymenttype.SetFocus
'    Exit Sub
'End If

If Val(Trim(Mid(Trim(comsetmode.Text), 1, 3))) = 7 And Trim(comsetmode.Text) <> "" And Val(Trim(Mid(Trim(cmbpaymenttype.Text), 1, 3))) <> 9 Then
    MsgBox "Please Check Your Payment Mode", vbInformation, "Invalid Choice"
    cmbpaymenttype.SetFocus
    Exit Sub
End If

If Trim(Mid(Trim(cmbpaymenttype.Text), 1, 2)) = "03" Then
    frmpromptforcheck.Show 1
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

If Val(lbldebitamount.caption) = Val(creditamountcount) Then
    'cmdadd.Enabled = False
    cmdsave.SetFocus
Else
    comcreditaccountnumber.SetFocus
End If

'to check all the values

creditcount = creditcount + 1
ListView1.ListItems.Add , , Trim(comcreditaccountnumber.Text)
ListView1.ListItems(creditcount).ListSubItems.Add , , Trim(ccname.Text)
ListView1.ListItems(creditcount).ListSubItems.Add , , Format(Trim(creditamount.Text), "############0.00#")
If creditcount = 29 Then
    'cmdadd.Enabled = False
    cmdsave.SetFocus
End If
End Sub

Private Sub cmdcancel_Click()
Dim reca As New ADODB.Recordset
reca.Open "delete from checkdeposittemp where paymentno = '" & Trim(txtreceiptnumber.Text) & "'", CON1, adOpenKeyset, adLockOptimistic
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
reccurrencychange.Open "select * from currencytable where status='default'", CON1, adOpenKeyset, adLockOptimistic
lblcurrency.caption = reccurrencychange!currency
reccurrencychange.close

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
txtdebitamt.Enabled = True
comsetmode.Enabled = True
txtinvoicenum.Enabled = True
cmbpaymenttype.Enabled = True
txtattachment.Enabled = True
cmdsave.Enabled = True
cmdnewrecord.Enabled = False
cmdprint.Enabled = False
recnum.Requery
txtreceiptnumber.Text = Val(recnum!paymentnumber)
txtreceiptnumber.Enabled = False
comreceivedfrom.SetFocus
Call prcclear
txtinvoicenum.Text = ""
End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub CMDEDIT_Click()
creditcount = 0
creditamount = 0
creditamountcount = 0
cmdnewrecord.Enabled = False
CmdEdit.Enabled = False
'cmdadd.Enabled = True
cmdshow.Enabled = True
cmdsave.Enabled = True
cmdcancel.Enabled = True
If UCase(cLogUser) = UCase("Cashier") Then
    cmdprint.Enabled = True
End If
Call prcclear
txtreceiptnumber.Enabled = True
txtreceiptnumber.Text = " "
SendKeys "{home}+{end}"
txtreceiptnumber.SetFocus
End Sub

Private Sub cmdPrint_Click()

'frmlanguagemessage.Show 1

If MsgBox("Select the Language to Print      ÇÎÊÇÑ áÛÉ ÇáØÈÇÚÉ  " & vbCrLf & "Yes :- For Arabic                                       ÈÇáÚÑÈí " & vbCrLf & "No  :- For English                                  ÈÇáÇäÌáíÒí", vbInformation + vbYesNo, "Language Option  ÇÎÊíÇÑ ÇááÛÉ") = vbYes Then
    frmpaymentvou.langopt.Value = 1
Else
     frmpaymentvou.langopt.Value = 0
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
Printer.FontName = "ARABIC TRANSPARENT"
Printer.FontItalic = False
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
'this is for different menues
    Printer.Print Tab(22); "Disbursement Voucher No.:-  "; addzero & Trim(txtreceiptnumber.Text); "  ÓäÏ ÕÑÝ  "
    
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
Printer.Print ; Tab(70); "Amount";
Printer.FontBold = False
Printer.Print Tab(110 - Len(amount & " " & Trim(lblcurrency.caption) & " ")); Format(Trim(txtdebitamt.Text), "###,###,###,###0.#0"); " "; Trim(lblcurrency.caption);
Printer.FontBold = True
Printer.Print Tab(132 - Len("..... ÇáãÈáÛ")); "..... ÇáãÈáÛ"
Printer.FontBold = False
Printer.FontSize = 2
Printer.Print ""
Printer.FontSize = 10
Printer.FontBold = True
Printer.Print "Pay To";
Printer.FontBold = False
'MsgBox langopt.Caption
If langopt.Value = 1 Then
    Printer.RightToLeft = True
    Printer.Print Tab(21); Trim(comreceivedfrom.Text);
    Printer.RightToLeft = False
Else
    Printer.Print Tab(25); Trim(comreceivedfrom.Text);
End If

Printer.FontBold = True
Printer.Print Tab(132 - Len("..... ÇÓÊáãäÇ ãä")); "..... ÇÓÊáãäÇ ãä"
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
Printer.Print Tab(133 - Len("..... ÇáãÈáÛ ÈÇáÍÑæÝ")); "..... ÇáãÈáÛ ÈÇáÍÑæÝ"
Printer.FontSize = 2
Printer.Print ""


Printer.FontSize = 10
Printer.Print "Payment Mode";
Printer.FontBold = False
'temporary
namenamenum = InStr(1, Trim(cmbpaymenttype.Text), "~", vbTextCompare)
nameara = Trim(Mid(Trim(cmbpaymenttype.Text), (namenamenum + 1), 30))
If namenamenum > 0 Then
NameEng = Mid(Trim(cmbpaymenttype.Text), 1, (namenamenum - 1))
Else
NameEng = Trim(cmbpaymenttype.Text)
End If
'end
If langopt.Value = 0 Then
    Printer.Print Tab(25); Trim(Mid(NameEng, 3, 30));
Else
    Printer.RightToLeft = True
    Printer.Print Tab(21); Trim(nameara);
    Printer.RightToLeft = False
End If
Printer.FontBold = True
Printer.Print Tab(132 - Len("..... äæÚ ÇáÏÝÚ")); "..... äæÚ ÇáÏÝÚ"
Printer.FontSize = 2
Printer.Print ""


Printer.FontBold = True
Printer.FontSize = 10
    If (Trim(Mid(Trim(cmbpaymenttype.Text), 1, 3)) = "03" Or Trim(Mid(Trim(cmbpaymenttype.Text), 1, 3)) = "02") And (Trim(Mid(Trim(comsetmode.Text), 1, 3)) <> "006" And Trim(Mid(Trim(comsetmode.Text), 1, 3)) <> "500") Then
        Printer.Print "Check/Ref No";
        Printer.FontBold = False
        Printer.Print Tab(25); Trim(txtchecknumber.Text);
        Printer.FontBold = True
        Printer.Print Tab(51); "Due Date";
        Printer.FontBold = False
        Printer.Print Tab(70); Format(txtcheckdate.Value, "dd/mm/yyyy");
        Printer.FontBold = True
        Printer.Print Tab(133 - Len("..... Çáíæã ãÓÊÍÞ ÇáÏÝÚ")); "..... Çáíæã ãÓÊÍÞ ÇáÏÝÚ"
    Else
        If (Trim(Mid(Trim(cmbpaymenttype.Text), 1, 3)) = "03" Or Trim(Mid(Trim(cmbpaymenttype.Text), 1, 3)) = "09" Or Trim(Mid(Trim(cmbpaymenttype.Text), 1, 3)) = "10") And (Trim(Mid(Trim(comsetmode.Text), 1, 3)) = "006" Or Trim(Mid(Trim(comsetmode.Text), 1, 3)) = "007" Or Trim(Mid(Trim(comsetmode.Text), 1, 3)) = "500") Then
            Printer.Print "Check No & Date ";
            Printer.FontBold = False
            Printer.Print Tab(25);
            
            Printer.FontSize = 8
            Printer.Print Trim(Mid(checknumberanddate, 1, 106))
            
            If Trim(Mid(checknumberanddate, 106, 211)) <> "" Then
                Printer.FontSize = 10
                Printer.Print Tab(25);
                
                Printer.FontSize = 8
                Printer.Print Trim(Mid(checknumberanddate, 107, 211))
                
            End If
            If Trim(Mid(checknumberanddate, 212, 317)) <> "" Then
                Printer.FontSize = 10
                Printer.Print Tab(25);
                
                Printer.FontSize = 8
                Printer.Print Trim(Mid(checknumberanddate, 212, 317))
                
            End If
        End If
    End If
Printer.FontSize = 2
Printer.Print " "


Printer.FontSize = 10
Printer.FontBold = True
Printer.Print "Payment Againts";
Printer.FontBold = False
'temporary
namenamenum = InStr(1, Trim(comsetmode.Text), "~", vbTextCompare)
nameara = Trim(Mid(Trim(comsetmode.Text), (namenamenum + 1), 30))
If namenamenum > 0 Then
NameEng = Mid(Trim(comsetmode.Text), 1, (namenamenum - 1))
Else
NameEng = Trim(comsetmode.Text)
End If
'end
If langopt.Value = 0 Then
    Printer.Print Tab(25); Trim(Trim(Mid(NameEng, 4, 30)));
Else
    Printer.RightToLeft = True
    Printer.Print Tab(21); Trim(nameara);
    Printer.RightToLeft = False
End If
Printer.FontBold = True
Printer.Print Tab(133 - Len("..... ãÞÇÈá ÇáÏÝÚ")); "..... ãÞÇÈá ÇáÏÝÚ"
Printer.FontBold = False
Printer.Print ""
Printer.FontBold = True
Printer.Print "Description ";
Printer.FontBold = False
If langopt.Value = 0 Then
    Printer.Print Tab(25); Trim(txtattachment.Text);
Else
    'delete this if it is not working
    Printer.RightToLeft = True
    Printer.Print Tab(21); Trim(txtattachment.Text);
    Printer.RightToLeft = False
    'end delete
    'Printer.Print Tab(118 - Len(Trim(Trim(comreceivedfrom.Text) & "  " & UCase(Trim(comfromname.Text))))); Trim(Trim(comreceivedfrom.Text) & "  " & UCase(Trim(comfromname.Text)));
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
Printer.Print Tab(13); "   (Payee/ÇáãÏÝæÚ áåþ)    "; Tab(53); "   (Manager/ÇáãÏíÑ)    "; Tab(92); "     (Cashier/Ããíä ÇáÕäÏæÞ)    "
On Error GoTo er
Printer.EndDoc
MsgBox "Voucher Printed Successfully  Êã ÇáØÈÚ ÈäÌÇÍ", vbInformation, "Print Conformation"

er:
If Err.Number = 482 Then
    If MsgBox("Please check your Printer And Turn ON And Press Yes." & vbCrLf & "  ÑÇÌÚ ÇáØÇÈÚå . ÇÝÊÍåÇ æÇÖÛØ ãæÇÝÞ", vbYesNo, "Not Ready") = vbYes Then
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

If Val(txtdebitamt.Text) <= 0 Then
    MsgBox "Please Check Your Receipt Amount", vbInformation, "Amount Missing"
    txtdebitamt.SetFocus
    Exit Sub
End If
If comreceivedfrom.Text = "" And comreceivedfrom.Enabled = True Then
MsgBox "Please check your payment to", vbInformation, "Invalid Data"
    comreceivedfrom.SetFocus
    Exit Sub
End If

If Trim(cmbpaymenttype.Text) = "" Then
    MsgBox "Plese check Your Payment Type", vbInformation, "Invalid PaymentType"
    cmbpaymenttype.SetFocus
    Exit Sub
End If

If Trim(comsetmode.Text) = "" Then
    MsgBox "Plese check Your Payment For", vbInformation, "Invalid PaymentFor"
    comsetmode.SetFocus
    Exit Sub
End If

If Trim(Mid(Trim(cmbpaymenttype.Text), 1, 3)) = "03" Then
    If Trim(checknumberanddate) = "" And Trim(txtchecknumber.Text) = "" Then
        MsgBox "Your check Number That is Missing", vbInformation, "missing Check Number"
        txtchecknumber.SetFocus
        Exit Sub
    End If
    
    If Trim(checknumberanddate) = "" And IsNull(txtcheckdate.Value) = True Then
        MsgBox "Your Check Due Date That is Missing", vbInformation, "Missing check Date"
        Exit Sub
    End If
    
    If Trim(txtchecknumber.Text) <> "" And IsNull(txtcheckdate.Value) = True Then
        MsgBox "Your Check Due Date That is Missing", vbInformation, "Missing check Date"
        Exit Sub
    End If
End If

If Trim(combdebitaccnum.Text) = "" Or Trim(cname.Text) = "" Then
    MsgBox "Plese check Debit Account Number or Debit Account Name", vbInformation, "Invalid PaymentMode"
    combdebitaccnum.SetFocus
    Exit Sub
End If
    
If Trim(Mid(Trim(comsetmode.Text), 1, 4)) = "006" Then ' for bank deposit
    If Trim(Mid(Trim(cmbpaymenttype.Text), 1, 3)) <> "03" And Trim(Mid(Trim(cmbpaymenttype.Text), 1, 3)) <> "01" And Trim(Mid(Trim(cmbpaymenttype.Text), 1, 3)) <> "10" Then
        MsgBox "Please Choose The Correct Payment Mode", vbInformation, "Invalid Data"
        comsetmode.SetFocus
        Exit Sub
    End If
End If

If Trim(Mid(Trim(cmbpaymenttype.Text), 1, 3)) <> "03" And Trim(Mid(Trim(comsetmode.Text), 1, 3)) = "006" And Trim(checknumberanddate) = "" Then
    MsgBox "Please Choose the Check to Deposit", vbInformation, "Empty Numbers"
    cmbpaymenttype.SetFocus
    Exit Sub
End If
    
If Trim(Mid(Trim(cmbpaymenttype.Text), 1, 3)) = "10" And Trim(checknumberanddate) = "" Then
        MsgBox "Your check Number is Missing For Deposit", vbInformation, "Check Number is missing"
        comsetmode.SetFocus
        Exit Sub
End If

ListCount = ListView1.ListItems.Count
For i = 1 To ListCount
    checkallamount = checkallamount + Val(Trim(ListView1.ListItems(i).ListSubItems(2).Text))
Next

If Val(checkallamount) <> Val(Trim(lbldebitamount.caption)) Then
    MsgBox "Please Check Credit Amount it is not Equal to Debit Amount", vbInformation, "Amount Overflow"
    checkallamount = 0
    Exit Sub
End If
    checkallamount = 0

frmconformpassword.Show 1

If stopstop1 = 1 Then
    stopstop1 = 0
    Exit Sub
End If

'this is for cashier
txtdebitamt.Enabled = True
recor.addnew

findrate.Open "Select * from currencytable where currency = '" & Trim(lblcurrency.caption) & "'", CON1, adOpenKeyset, adLockOptimistic
    recor!currencyrate = Val(findrate!rate)
    recor!currencymark = Trim(lblcurrency.caption)
    If findrate!Status = "default" Then
        recor!currencydefault = 1
    End If
findrate.close

recor!receiptno = Trim(txtreceiptnumber.Text)
recor!receiptdate = Format(Date, "mm/dd/yyyy")
'recor!custno = " "
recor!custname = Trim(comreceivedfrom.Text)
recor!deleted = "0"
recor!okprint = "0"
'recor!dollarcheck = "0"
If bankname <> "" Then
recor!bankname = bankname

Dim recfindbankcode As New ADODB.Recordset
recfindbankcode.Open "select accountcode from financemaster where rtrim(accountnameeng) = '" & Trim(bankname) & "'", CON1, adOpenKeyset, adLockOptimistic
recor!banknumber = recfindbankcode!AccountCode
recfindbankcode.close

bankokok = InStr(1, bankname, "$", vbTextCompare)
    If bankokok > 0 Then
        'recor!dollarcheck = "1"
    End If
End If
'If Trim(Mid(Trim(cmbpaymenttype.Text), 1, 3)) = "03" Or Trim(Mid(Trim(cmbpaymenttype.Text), 1, 3)) = "10" Or Trim(Mid(Trim(cmbpaymenttype.Text), 1, 3)) = "09" Then
'    recor!checkreceipt = Trim(txtdebitamt.Text)
'    recor!creditamount = "0"
'    recor!doller = "0"
'    recor!DebitAmount = "0"
'Else
'    If Trim(Mid(Trim(cmbpaymenttype.Text), 1, 3)) = "07" Then
'        recor!doller = Trim(txtdebitamt.Text)
'        recor!creditamount = "0"
'        recor!checkreceipt = "0"
'        recor!DebitAmount = "0"
'    Else
'       recor!creditamount = Trim(txtdebitamt.Text)
'       recor!doller = "0"
'       recor!checkreceipt = "0"
'       recor!DebitAmount = "0"
'    End If
'End If
recor!receiptamount = Val(txtdebitamt.Text)
recor!payopt = Trim(comsetmode.Text)
recor!optref = Trim(txtinvoicenum.Text) & " "
recor!paymode = Trim(cmbpaymenttype.Text)

If (Trim(Mid(Trim(cmbpaymenttype.Text), 1, 3)) = "03" Or Trim(Mid(Trim(cmbpaymenttype.Text), 1, 3)) = "10") And Trim(Mid(Trim(comsetmode.Text), 1, 3)) = "006" Then
    recor!moderef = "Checks For Deposit"
ElseIf Mid(Trim(cmbpaymenttype.Text), 1, 3) = "09" And Trim(Mid(Trim(comsetmode.Text), 1, 3)) = "007" Then
    recor!moderef = "Returned to Clients"
End If

If (Trim(Mid(Trim(cmbpaymenttype.Text), 1, 3)) = "03" Or Trim(Mid(Trim(cmbpaymenttype.Text), 1, 3)) = "10") And Trim(txtchecknumber.Text) <> "" Then
    recor!moderef = Trim(txtchecknumber.Text)
    If IsNull(txtcheckdate.Value) <> True Then
        recor!chkdue = Format(txtcheckdate.Value, "mm/dd/yyyy")
    End If
End If
'recor!ticketnumber = 1
'recor!ncount = "909"
recor!Post = "no"
recor!svoucher = "Disbursements"
'this is for take the journal number
recset.Requery
totaljournal = "CSP-" & recset!CurrentMoYr & "-" & Trim(recset!nextjn)
recor!JournalNumber = totaljournal
totaljournal = "0000" & (Val(Trim(recset!nextjn)) + 1)
    recset!nextjn = Right(totaljournal, 5)
    recset.Update
recor!remarks = Trim(txtattachment.Text)
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
recor!accountnumber = Trim(combdebitaccnum.Text)
recor.Update

recnum!paymentnumber = Val(Val(Trim(txtreceiptnumber.Text)) + 1)
recnum.Update
                          
'this is for payable setup
    If okcontinue = 1 Then
        recpayable!Paidmark = 1
        recpayable.Update
        recpayable.close
        
        'after that this is for payablejournal
        recpayable.Open "Select * from payjournal where serno = '" & Trim(txtinvoicenum.Text) & "'", CON1, adOpenKeyset, adLockOptimistic
        If recpayable.BOF = False Then
            While recpayable.EOF = False
                natheerdetails = "PYB No " & Trim(recpayable!serno) & " \" & IIf(IsNull(recpayable!InvNo), " ", " I.No " & Trim(recpayable!InvNo) & " \") & IIf(IsNull(recpayable!SENumber), " ", " S.E.No " & Trim(recpayable!SENumber) & " \") & " CSP No " & Trim(txtreceiptnumber.Text) & " \" & Format(Date, "mmm/dd/yyyy")
                recpayable!particulars = natheerdetails
                natheerjournal = Trim(recpayable!SerialNo)
                recpayable.Update
                recpayable.MoveNext
            Wend
        End If
        recpayable.close
        
        'after that this is for the glmaster
        If Trim(natheerjournal) <> "" Then
            recpayable.Open "Select * from glmaster where Journalno = '" & natheerjournal & "'", CON1, adOpenKeyset, adLockOptimistic
            If recpayable.BOF = False Then
                While recpayable.EOF = False
                    recpayable!particulars = natheerdetails
                    recpayable.Update
                    recpayable.MoveNext
                Wend
            End If
            recpayable.close
            natheerjournal = ""
            natheerdetails = ""
        End If
    End If
      okcontinue = 2 ' see previous lines
      
If Mid(Trim(cmbpaymenttype.Text), 1, 2) = "03" And Trim(txtchecknumber.Text) <> "" Then
'this is for checkregister for update the cheques
recche.Open "select * from checkregister where checknumber = " & "'" & (txtchecknumber.Text) & "'" & " and bank = " & "'" & bankname & "'", conpay, adOpenKeyset, adLockOptimistic
      If recche.BOF = False Then
        With recche
        !receiptno = Trim(txtreceiptnumber.Text)
        !issuedto = Trim(comreceivedfrom.Text)
        !valuedate = txtcheckdate.Value
        .Update
        End With
      End If
recche.close
End If

'this all for checkdeposittemp table
    Dim mystring As String
    Dim recdeldeposit As New ADODB.Recordset
    Dim condeldeposit As New ADODB.Connection
    Dim constringdel As String
    Dim recdelfind As New ADODB.Recordset
    Dim condelfind As New ADODB.Connection
    
    constringdel = "dsn=finance;uid=sa;pwd=;"
    
If (Trim(Mid(Trim(cmbpaymenttype.Text), 1, 3)) = "03" Or Mid(Trim(cmbpaymenttype.Text), 1, 2) = "10" Or Mid(Trim(cmbpaymenttype.Text), 1, 2) = "09") And (Trim(Mid(Trim(comsetmode.Text), 1, 3)) = "006" Or Trim(Mid(Trim(comsetmode.Text), 1, 3)) = "007" Or Trim(Mid(Trim(comsetmode.Text), 1, 3)) = "500") Then
    sqltable = True
    xtable = "delete from checkdeposittemp where paymentno = '" & Trim(txtreceiptnumber.Text) & "' and selected is null"
    myclass.GetTables recdeldeposit, condeldeposit, xtable, constringdel, sqltable
    condeldeposit.close
    
    'this is to update the voucher table in collection checks
    sqltable = True
    xtable = "select * from checkdeposittemp where paymentno = '" & Trim(txtreceiptnumber.Text) & "' and selected = '1'"
    myclass.GetTables recdeldeposit, condeldeposit, xtable, constringdel, sqltable
    'MsgBox recdeldeposit.RecordCount
        If recdeldeposit.BOF = False Then
            While recdeldeposit.EOF = False
                    sqltable = True
                    xtable = "select * from vouchers where (substring(paymode,1,2) = '03' or substring(paymode,1,2) = '10') and " & _
                    "deleted <> '1' and deposit = '0' and svoucher = 'Collections' and receiptno = '" & recdeldeposit!receiptno & "'"
                    myclass.GetTables recdelfind, condelfind, xtable, constringdel, sqltable
                    recdelfind!deposit = "1"
                    recdelfind!banknumber = Trim(comreceivedfrom.Text)
                    If Trim(Mid(Trim(comsetmode.Text), 1, 3)) = "007" And (Val(Mid(Trim(cmbpaymenttype.Text), 1, 2)) = 9) Then
                        recdelfind!paymode = "09     Cashed in Check"
                        recdelfind!deposit = 9
                    End If
                    recdelfind.Update
                    recdelfind.close
                    condelfind.close
                recdeldeposit.MoveNext
            Wend
        End If
    recdeldeposit.close
    condeldeposit.close
'this is for check deposit report

'    On Error Resume Next
'    dataanu.rscom_checkdeposit_Grouping.Close
'    On Error GoTo 0
'    dataanu.com_checkdeposit_Grouping Trim(txtreceiptnumber.Text)
'    re_cashier_checkdeposit.Show 1
Else

'this is to delete the checkdeposittemp table
    sqltable = True
    xtable = "delete checkdeposittemp where paymentno = '" & Trim(txtreceiptnumber.Text) & "'"
    myclass.GetTables recdeldeposit, condeldeposit, xtable, constringdel, sqltable
    condeldeposit.close
End If
     Dim deletedata As New ADODB.Recordset
     
     'to insert the first ticket to cashjournal
     If Trim(Mid(Trim(cmbpaymenttype.Text), 1, 3)) = "07" Then
         deletedata.Open "insert into cashjournal SELECT JournalNumber, '1', '0', Accountnumber, AccountName, accountnamearab, mothername," & _
         "convert(nvarchar(10),vouchers.receiptno) + ' \ ' + isnull(vouchers.remarks,'') + '\'+ ltrim(substring(vouchers.paymode,4,20)) + '\' + isnull(vouchers.moderef,'') + '-' + isnull(convert(nvarchar(11),vouchers.chkdue),'')" & _
         ",againtsname+' \ '+'" & lastuser & "',ReceiptNo,'" & Format(Date, "mm/dd/yyyy") & "',0,receiptamount,NULL,'UnPosted','P' FROM vouchers where receiptno = '" & Trim(txtreceiptnumber.Text) & "' and svoucher = 'Disbursements'", CON1, adOpenKeyset, adLockOptimistic
     Else
         deletedata.Open "insert into cashjournal SELECT JournalNumber, '1', '0', Accountnumber, AccountName, accountnamearab, mothername," & _
         "convert(nvarchar(10),vouchers.receiptno) + ' \ ' + isnull(vouchers.remarks,'') + '\'+ ltrim(substring(vouchers.paymode,4,20)) + '\' + isnull(vouchers.moderef,'') + '-' + isnull(convert(nvarchar(11),vouchers.chkdue),'')" & _
         ",againtsname+' \ '+'" & lastuser & "',ReceiptNo,'" & Format(Date, "mm/dd/yyyy") & "',0,receiptamount,NULL,'UnPosted','P' FROM vouchers where receiptno = '" & Trim(txtreceiptnumber.Text) & "' and svoucher = 'Disbursements'", CON1, adOpenKeyset, adLockOptimistic
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
            !DebitAmount = Trim(ListView1.ListItems(xxx - 10).ListSubItems(2).Text)
            !creditamount = 0
            !Status = "UnPosted"
            !Trantype = "P"
            .Update
            addParticulars = ""
    End With
 Next

Me.caption = "Disbursement Voucher"
MsgBox "Record Updated Successfully" & vbCrLf & " Êã ÇáÍÝÙ", vbInformation, "Data Saved"
cmdcancel.Enabled = False
creditcount = 0
creditamountcount = 0
Frame1.Enabled = False
fromwho = "none"
cmdprint.Enabled = True
bankname = ""
cmdsave.Enabled = False
cmdnewrecord.Enabled = True
cmdPrint_Click

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
         .ListItems(l2).ListSubItems.Add , , Trim(reclist!paymode)
         .ListItems(l2).ListSubItems.Add , , Trim(reclist!payopt)
         .ListItems(l2).ListSubItems.Add , , Format(Trim(reclist!receiptamount), "############0.#0")
     End With
 reclist.MoveNext
 DoEvents
 Wend
End If
End Sub

'
'Private Sub cmdshowinvoice_Click()
'If Trim(Mid(Trim(comreceivedfrom.Text), 1, 10)) <> "" And Val(Trim(Mid(Trim(comsetmode.Text), 1, 3))) = 1 And Trim(txtdebitamt.Text) <> "" Then
'frminvoice.activeform = 0
'frminvoice.anucustcode = Trim(Mid(Trim(comreceivedfrom.Text), 1, 10))
'On Error GoTo er:
'frminvoice.Label5.Visible = False
'frminvoice.Label6.Visible = False
'frminvoice.ListView1.Enabled = False
'frminvoice.cmdcancel.Visible = False
'frminvoice.cmdclose.Caption = "Close"
'frminvoice.Timer1.Interval = 0
'frminvoice.Show 1
'End If
'
'er:
'If Err.Number = 360 Then
'End If
'
'End Sub

'Private Sub cmdshow_Click()
'frmcheckdeposit.showclick = 1
'frmcheckdeposit.Show 1
'End Sub

Private Sub cname_Click()
Dim recfindacc As New ADODB.Recordset
namenamenum = InStr(1, Trim(cname.Text), "\", vbTextCompare)
If namenamenum > 0 Then
namename = Mid(Trim(cname.Text), 1, (namenamenum - 1))
Else
namename = Trim(cname.Text)
End If
recfindacc.Open "Select * from financemaster where active <> '0' and  accountnameeng = " & "'" & namename & "'", CON1, adOpenKeyset, adLockOptimistic
If recfindacc.BOF = False Then
    combdebitaccnum.Text = recfindacc!AccountCode
End If
 recfindacc.close
 'this is for the mother name
    getmothername nonumber, namename, cc
    Me.caption = "Disbursement Voucher " & cc
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
Else
namename = Trim(cname.Text)
End If
recfindacc.Open "Select * from financemaster where active <> '0' and  accountnameeng = " & "'" & namename & "'", CON1, adOpenKeyset, adLockOptimistic
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
Dim recfindaccount As New ADODB.Recordset
recfindaccount.Open "select * from financemaster where active <> '0' and  accountcode = " & "'" & Trim(combdebitaccnum.Text) & "'", CON1, adOpenKeyset, adLockOptimistic
If recfindaccount.BOF = False Then
    cname.Text = Trim(recfindaccount!accountnameeng) & "\" & Trim(recfindaccount!accountnamearab)
End If
recfindaccount.close
    getmothername Trim(combdebitaccnum.Text), noname, cc
   Me.caption = "Disbursement Voucher " & cc
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
If KeyAscii = 13 Then
    cname.SetFocus
End If
End Sub

Private Sub combdebitaccnum_LostFocus()
Dim recfindaccount As New ADODB.Recordset
If Trim(combdebitaccnum.Text) <> "" Then
recfindaccount.Open "select * from financemaster where active <> '0' and  accountcode = " & "'" & Trim(combdebitaccnum.Text) & "'", CON1, adOpenKeyset, adLockOptimistic
If recfindaccount.BOF = True Then
    MsgBox "Please Choose the Correct Account", vbInformation, "Invalid AccountNumber"
    combdebitaccnum.SetFocus
    Exit Sub
Else
    cname.Text = recfindaccount!accountnameeng & "\" & recfindaccount!accountnamearab
End If
recfindaccount.close
End If

End Sub
Private Sub comcreditaccountnumber_Click()
Dim recfindaccount As New ADODB.Recordset
recfindaccount.Open "select * from financemaster where active <> '0' and accountcode = " & "'" & Trim(comcreditaccountnumber.Text) & "'", CON1, adOpenKeyset, adLockOptimistic
If recfindaccount.BOF = False Then
    ccname.Text = Trim(recfindaccount!accountnameeng) & "\" & Trim(recfindaccount!accountnamearab)
End If
recfindaccount.close

    getmothername Trim(comcreditaccountnumber.Text), noname, cc
    Me.caption = "Disbursement Voucher  " & cc
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
If KeyAscii = 13 Then
    ccname.SetFocus
End If
End Sub
Private Sub comcreditaccountnumber_LostFocus()
Dim recfindaccount As New ADODB.Recordset
If Trim(comcreditaccountnumber.Text) <> "" Then
recfindaccount.Open "select * from financemaster where active <> '0' and accountcode = " & "'" & Trim(comcreditaccountnumber.Text) & "'", CON1, adOpenKeyset, adLockOptimistic
If recfindaccount.BOF = True Then
    recfindaccount.close
    MsgBox "Please choose the Correct Account Number", vbInformation, "Invalid Account Number"
    ccname.SetFocus
    Exit Sub
Else
    ccname.Text = recfindaccount!accountnameeng & "\" & recfindaccount!accountnamearab
End If
recfindaccount.close
End If
Me.caption = "Disbursement Voucher  " & cc
End Sub

Private Sub Command1_Click()
cmdPrint_Click
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
If KeyAscii = 13 Then
    txtdebitamt.SetFocus
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

If KeyAscii = 13 Then
    If txtattachment.Enabled = False = True Then
        combdebitaccnum.SetFocus
        Exit Sub
   ElseIf txtattachment.Enabled = True Then
        txtinvoicenum.SetFocus
    ElseIf Trim(comsetmode.Text) = "" Then
        comsetmode.SetFocus
    Else
    cmdsave.SetFocus
    End If
End If
End Sub

Private Sub comsetmode_LostFocus()

If Trim(Mid(Trim(comsetmode.Text), 1, 3)) = "006" And Trim(Mid(Trim(cmbpaymenttype.Text), 1, 2)) = "03" And Trim(txtchecknumber.Text) <> "" Then
   MsgBox "You Cannot Deposit the Company Check", vbInformation, "Not Valid"
   comsetmode.ListIndex = 6
   Exit Sub
End If

If Trim(comsetmode.Text) <> "" Then
    If Val(Trim(Mid(Trim(cmbpaymenttype.Text), 1, 3))) = 9 And Trim(Mid(Trim(comsetmode.Text), 1, 3)) <> "007" Then
        MsgBox "Please Check the Payment Type or Payment Mode", vbInformation, "Invalid Selection"
        On Error Resume Next
        comsetmode.SetFocus
        On Error GoTo 0
        Exit Sub
    End If
    On Error Resume Next
    If Trim(Mid(Trim(comsetmode.Text), 1, 3)) = "007" And Val(Trim(Mid(Trim(cmbpaymenttype.Text), 1, 3))) <> 9 Then
         MsgBox "Please Check the Payment Type or Payment Mode", vbInformation, "Invalid Selection"
        comsetmode.SetFocus
        Exit Sub
    End If
    On Error GoTo 0
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

If KeyAscii = 13 Then 'And cmdadd.Enabled = True Then
    cmdAdd_Click
End If
'If KeyAscii = 13 And cmdadd.Enabled = False Then
'cmdsave.SetFocus
'End If

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
End Sub

Private Sub Form_Load()
Me.Height = 5835

Timer2.Interval = 1
Label33.caption = Time
Mainform.sbStatusBar.Panels(1).Text = "Status : Disbursement Voucher ..."
constring = "dsN=fINANCE;UID=SA;PWD=;"
lbldate.caption = Format(Date, "mm/dd/yyyy")

xtable = "Select * from vendor" ' the marcusfl table connections
sqltable = True
myclass.GetTables recmar, CON1, xtable, constring, sqltable

conpay.Open "dsN=fINANCE;UID=SA;PWD=;" 'the payee table connections
recpay.Open "Select * from paymode where newcode is not null", conpay, adOpenKeyset, adLockOptimistic

recmod.Open "Select * from setmode", conpay, adOpenKeyset, adLockOptimistic ' the setmod table connection


Dim reccurrency As New ADODB.Recordset
reccurrency.Open "select * from currencytable where detail = 'Cash' and status='default'", conpay, adOpenKeyset, adLockOptimistic
lblcurrency.caption = reccurrency!currency
reccurrency.close
reccurrency.Open "Select * from currencytable", conpay, adOpenKeyset, adLockOptimistic

'list currency
listcurrency.ListItems.clear
i = 1
While reccurrency.EOF = False
    If Trim(reccurrency!detail) = "Cash" Then
        listcurrency.ListItems.Add , , reccurrency!currency
        listcurrency.ListItems(i).ListSubItems.Add , , Format(reccurrency!rate, "###0.###0")
        listcurrency.ListItems(i).ListSubItems.Add , , reccurrency!latestupdate
        i = i + 1
    End If
    reccurrency.MoveNext
Wend
'end if

Dim conrecacc As New ADODB.Connection ' the financemaster table
xtable = "select * from financemaster where active <> '0' "
sqltable = True
myclass.GetTables recacc, conrecacc, xtable, constring, sqltable


recnum.Open "select * from reNumber", conpay, adOpenKeyset, adLockOptimistic 'the receipt voucher number table

recor.Open "select * from vouchers", conpay, adOpenKeyset, adLockOptimistic

'check for the month and year change if we have to reset the series number.
recset.Open "Select getdate() as currentmoyr", CON1, adOpenKeyset, adLockOptimistic
Dim checkjournaldate As Date
checkjournaldate = Left(Trim(recset!CurrentMoYr), 10)
recset.close
recset.Open "Select * from setup", conpay, adOpenKeyset, adLockOptimistic
recset.MoveFirst
If Trim(recset!CurrentMoYr) <> Format(checkjournaldate, "mmyy") Then
    recset!CurrentMoYr = Format(checkjournaldate, "mmyy")
    recset!nextjn = "0001"
    recset.Update
End If
'end checking the journal no

Dim conlist As New ADODB.Connection
xtable = "Select * from vouchers where deleted = '0' and post ='no' and svoucher ='Disbursements' order by receiptno"
sqltable = True
myclass.GetTables reclist, conlist, xtable, constring, sqltable
    
comreceivedfrom.clear
Dim recpayee As New ADODB.Recordset
xtable = "select * from financemaster where active <> '0' and substring(accountcode,1,3) = '131'"
sqltable = True
Dim conrecpayee As New ADODB.Connection
myclass.GetTables recpayee, conrecpayee, xtable, constring, sqltable

While recpayee.EOF = False
If Trim(recpayee!accountnamearab) = "" Then
    anu = recpayee!accountnameeng
Else
    anu = recpayee!accountnamearab & "\" & recpayee!accountnameeng
End If
    comreceivedfrom.AddItem anu
    recpayee.MoveNext
Wend
recpayee.close
conrecpayee.close

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
                'hellosay = reclist!ncount
                .ListItems(l2).ListSubItems.Add , , Trim(reclist!paymode)
                .ListItems(l2).ListSubItems.Add , , Trim(reclist!payopt)
                .ListItems(l2).ListSubItems.Add , , Format(Trim(reclist!receiptamount), "############0.#0")
            End With
        reclist.MoveNext
        DoEvents
        Wend
End If

xtable = "select * from financemaster where active <> '0' and substring(accountcode,1,5) = '11102'"
sqltable = True
myclass.GetTables recpayee, conrecpayee, xtable, constring, sqltable

While recpayee.EOF = False
If Trim(recpayee!accountnamearab) = "" Then
    anu = recpayee!accountnameeng
Else
    anu = recpayee!accountnamearab & "\" & recpayee!accountnameeng & " - " & recpayee!AccountCode
End If
    comreceivedfrom.AddItem anu
    recpayee.MoveNext
Wend
recpayee.close
conrecpayee.close

xtable = "select * from payee"
sqltable = True
myclass.GetTables recpayee, conrecpayee, xtable, constring, sqltable
While recpayee.EOF = False
anu = recpayee!payeecode & "     " & recpayee!payeenameeng
    comreceivedfrom.AddItem anu
    recpayee.MoveNext
Wend
recpayee.close
conrecpayee.close


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
stopstop1 = 0

combdebitaccnum.clear
cname.clear
ccname.clear
comcreditaccountnumber.clear

recacc.MoveFirst
While recacc.EOF = False
    If Mid(Trim(recacc!AccountCode), 1, 5) = "11101" Or Mid(Trim(recacc!AccountCode), 1, 5) = "11102" Or Mid(Trim(recacc!AccountCode), 1, 5) = "11105" Or Mid(Trim(recacc!AccountCode), 1, 5) = "11106" Or Mid(Trim(recacc!AccountCode), 1, 5) = "11107" Then
        combdebitaccnum.AddItem Trim(recacc!AccountCode)
        cname.AddItem Trim(recacc!accountnameeng) & "\" & Trim(recacc!accountnamearab)
    End If
        comcreditaccountnumber.AddItem Trim(recacc!AccountCode)
        ccname.AddItem Trim(recacc!accountnameeng) & "\" & Trim(recacc!accountnamearab)
    recacc.MoveNext
Wend
ListView1.ListItems.clear
lbldebitamount.caption = " "
creditamount.Text = "  "

Mainform.sbStatusBar.Panels(1).Text = "Status : Ready."
Timer2.Interval = 0
lbldate.caption = Format(Date, "dd/mm/yyyy")
End Sub

Private Sub Form_Unload(Cancel As Integer)
cmdcancel_Click
conpay.close
End Sub

Private Sub langopt_Click()

If langopt.Value = 1 Then
    langopt.caption = "ÇÎÊíÇÑ ÇáØÈÇÚÉ (Arabic)"
    comreceivedfrom.RightToLeft = True
    txtattachment.RightToLeft = True
Else
    langopt.caption = "ÇÎÊíÇÑ ÇáØÈÇÚÉ (English)"
    comreceivedfrom.RightToLeft = False
    txtattachment.RightToLeft = False
End If
If comreceivedfrom.Enabled = True Then
    comreceivedfrom.SetFocus
End If
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
    listcurrency.Top = 1200
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

Private Sub ListView1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
    delete_Click
End If
End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
If ListView1.ListItems.Count > 0 And Button = vbRightButton Then
    PopupMenu file
End If
End Sub

Private Sub Timer1_Timer()
Dim recreceipt As New ADODB.Recordset
Dim recpayment As New ADODB.Recordset
Dim recopening As New ADODB.Recordset
Dim recclosing As New ADODB.Recordset

recopening.Open "Select * from Balance order by balancedate", conpay, adOpenKeyset, adLockOptimistic
recreceipt.Open "SELECT SUM(debitamount) AS receipttotal From vouchers WHERE (deleted = '0') and okprint = '0' AND svoucher ='Collections' AND (POST = 'no') AND (debitamount > 0) AND (SUBSTRING(PAYMODE, 1, 2) = '01') AND receiptDate = " & "'" & Format(Date, "mm/dd/yyyy") & "'", conpay, adOpenKeyset, adLockOptimistic
recpayment.Open "SELECT SUM(creditamount) AS paymenttotal From vouchers WHERE (deleted = '0') and okprint = '0' AND svoucher ='Disbursements' AND (POST = 'no') AND (creditamount > 0) AND (SUBSTRING(PAYMODE, 1, 2) = '01') AND receiptDate = " & "'" & Format(Date, "mm/dd/yyyy") & "'", conpay, adOpenKeyset, adLockOptimistic

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

Private Sub timaccountcode_Timer()
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

Private Sub txtattachment_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    comcreditaccountnumber.SetFocus
End If

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

Public Sub Changeamount()
        Dim findrate As New ADODB.Recordset
        findrate.Open "Select * from currencytable where currency = '" & Trim(lblcurrency.caption) & "'", CON1, adOpenKeyset, adLockOptimistic
        Dim allamount As Currency
        allamount = Val(Trim(txtdebitamt.Text)) * Val(findrate!rate)
        lbldebitamount.caption = Format(allamount, "############0.#0")
        findrate.close
End Sub

Private Sub txtdebitamt_GotFocus()
xdecimal = 0
'SendKeys "{End}+{Home}"
End Sub

Private Sub txtdebitamt_KeyPress(KeyAscii As Integer)
'cheking for the dot key
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
End Sub

Private Sub txtdebitamt_LostFocus()
xdecimal = 0
End Sub

Private Sub txtdetails_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtattachment.SetFocus
End If
End Sub

Private Sub txtinvoicenum_GotFocus()
txtinvoicenum.Text = ""
comreceivedfrom.Enabled = True
txtdebitamt.Enabled = True
End Sub

Private Sub txtinvoicenum_KeyPress(KeyAscii As Integer)

'this is for payablesetup
On Error Resume Next
recpayable.close
On Error GoTo 0

recpayable.Open "SELECT * FROM PayableSetup WHERE (CancelledMark = '0') AND (ConfirmedMark = '1') AND (Paidmark <> '1') AND (DeleteMark <> '1') and post = 'No' and serialno ='" & Trim(txtinvoicenum.Text) & "'", conpay, adOpenKeyset, adLockOptimistic

If KeyAscii = 13 Then
If Val(Trim(Mid(Trim(comsetmode.Text), 1, 6))) = 2 Or Val(Trim(Mid(Trim(comsetmode.Text), 1, 6))) = 5 Then
    If recpayable.BOF = True Then
        MsgBox "No Any Confirmed Payables Available " & vbCrLf & "  áÇ ÊæÌÏ ÍÑßå ÈÇáãæÇÝÞÇÊ ááÕÑÝ . ÇÊÕá ÈÇáÍÓÇÈÇÊ", vbInformation, "Invalid Number"
        txtinvoicenum.Text = ""
        comsetmode.ListIndex = 5
        comsetmode.SetFocus
        Exit Sub
    Else
        recpayable.MoveFirst
            If Trim(recpayable!SerialNo) = Trim(txtinvoicenum.Text) Then
                okcontinue = 1
                If Val(Mid(Trim(cmbpaymenttype.Text), 1, 2)) = 7 Then
                    txtdebitamt.Text = Format(Trim(recpayable!FcAmount), "############0.#0")
                Else
                    txtdebitamt.Text = Format(Trim(recpayable!amtreqested), "############0.#0")
                End If
                comreceivedfrom.Text = Trim(recpayable!Payee)
                txtdebitamt.Enabled = False
                comreceivedfrom.Enabled = False
            End If
    End If
    
    If okcontinue <> 1 Then
        MsgBox "This Number is Not Approved " & vbCrLf & " åÐÇ ÇáÑÞã áã íæÇÝÞ Úáíå ", vbInformation, "Invalid Number"
        comsetmode.ListIndex = comsetmode.ListCount - 1
        txtinvoicenum.Text = ""
        comsetmode.SetFocus
        Exit Sub
    End If
End If
    txtattachment.SetFocus
End If

End Sub

Private Sub txtreceiptnumber_KeyPress(KeyAscii As Integer)
sqltable = True
Dim connectionstring As String
Dim recorrr As New ADODB.Recordset

Dim conforsearch As New ADODB.Connection
connectionstring = "dsn=finance;uid=sa;pwd=;"
xtable = "select * from vouchers where svoucher ='Disbursements' and receiptno = '" & Trim(txtreceiptnumber.Text) & "'"
myclass.GetTables recorrr, conforsearch, xtable, connectionstring, sqltable
If KeyAscii = 13 Then
    If recorrr.BOF = True Then
        recorrr.close
        conforsearch.close
        MsgBox "This Number is Not Valid", vbInformation, "Roll Back"
        Exit Sub
    End If
    If Trim(recorrr!Post) = "Yes" Then
        'this is for the procedure
txtreceiptnumber.Text = recorrr!receiptno
lbldate.caption = recorrr!receiptdate
comreceivedfrom.Text = recorrr!custno & "     " & recorrr!custname
tookcredit = 0

       If IsNull(recorrr!creditamount) Then
            tookcredit = 0
        Else
            tookcredit = Trim(recorrr!creditamount)
        End If
        
        If tookcredit = 0 Then
            If IsNull(recorrr!checkreceipt) Then
                tookcredit = 0
            Else
                tookcredit = Trim(recorrr!checkreceipt)
            End If
        End If
        
        If tookcredit = 0 Then
            If Val(Trim(Mid(Trim(comsetmode.Text), 1, 6))) = 2 Or Val(Trim(Mid(Trim(comsetmode.Text), 1, 6))) = 5 Then
                Dim rectakeamount As New ADODB.Recordset
                rectakeamount.Open "SELECT * FROM PayableSetup WHERE (CancelledMark = '0') AND (ConfirmedMark = '1') AND (Paidmark = '1') AND (DeleteMark <> '1') and serialno ='" & Trim(recorrr!optref) & "'", CON1, adOpenKeyset, adLockOptimistic
                tookcredit = Format(Trim(rectakeamount!amtreqested), "############0.#0")
                rectakeamount.close
            Else
                tookcredit = Trim(recorrr!doller)
                tookcredit = tookcredit * Rate1
            End If
        End If

txtdebitamt.Text = tookcredit
lbldebitamount.caption = Format(tookcredit, "#########0.#0")
tookcredit = 0
comsetmode.Text = recorrr!payopt
txtinvoicenum.Text = recorrr!optref & " "
cmbpaymenttype.Text = recorrr!paymode
takejournal = recorrr!JournalNumber ' journal number
If IsNull(recorrr!chkdue) = False Then
txtcheckdate.Value = recorrr!chkdue
txtchecknumber.Text = recorrr!moderef

End If

comprepared.Text = recorrr!cashier
comchecked.Text = recorrr!empcheck
comapproved.Text = recorrr!empapp
txtattachment.Text = recorrr!remarks
'end procedure
        Frame1.Enabled = True
        comreceivedfrom.Enabled = False
        txtdebitamt.Enabled = False
        comsetmode.Enabled = True
        txtinvoicenum.Enabled = False
        cmbpaymenttype.Enabled = True
        txtchecknumber.Enabled = True
        txtcheckdate.Enabled = True
        comprepared.Enabled = False
        comchecked.Enabled = False
        comapproved.Enabled = False
        txtattachment.Enabled = False
        Exit Sub
    End If
   If Trim(recorrr!ncount) = "909" Then
        If MsgBox("Are Your Sure Your Want to Analyse Again ?", vbYesNo + vbQuestion + vbDefaultButton3, "Conformation") = vbYes Then
            'this cording is exist in saving for accountant
        Else
            recorrr.close
            conforsearch.close
            cmdcancel_Click
            Exit Sub
        End If
   End If
        Frame1.Enabled = True
        comreceivedfrom.Enabled = False
        txtdebitamt.Enabled = False
        comsetmode.Enabled = True
        txtinvoicenum.Enabled = False
        cmbpaymenttype.Enabled = True
        txtchecknumber.Enabled = True
        txtcheckdate.Enabled = True
        comprepared.Enabled = False
        comchecked.Enabled = False
        comapproved.Enabled = False
        txtattachment.Enabled = False
'this is for the procedure
txtreceiptnumber.Text = recorrr!receiptno
lbldate.caption = recorrr!receiptdate
comreceivedfrom.Text = recorrr!custno & "     " & recorrr!custname
tookcredit = 0
        If IsNull(recorrr!creditamount) Then
            tookcredit = 0
        Else
            tookcredit = Trim(recorrr!creditamount)
        End If
        
        If tookcredit = 0 Then
            If IsNull(recorrr!checkreceipt) Then
                tookcredit = 0
            Else
                tookcredit = Trim(recorrr!checkreceipt)
            End If
        End If
        
        If tookcredit = 0 Then
            If Val(Trim(Mid(Trim(comsetmode.Text), 1, 6))) = 2 Or Val(Trim(Mid(Trim(comsetmode.Text), 1, 6))) = 5 Then
                rectakeamount.Open "SELECT * FROM PayableSetup WHERE (CancelledMark = '0') AND (ConfirmedMark = '1') AND (Paidmark = '1') AND (DeleteMark <> '1') and serialno ='" & Trim(recorrr!optref) & "'", CON1, adOpenKeyset, adLockOptimistic
                tookcredit = Format(Trim(rectakeamount!amtreqested), "############0.#0")
                rectakeamount.close
            Else
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
        End If
        
txtdebitamt.Text = tookcredit
lbldebitamount.caption = Format(tookcredit, "#########0.#0")
tookcredit = 0
comsetmode.Text = recorrr!payopt
txtinvoicenum.Text = recorrr!optref & " "
cmbpaymenttype.Text = recorrr!paymode
takejournal = recorrr!JournalNumber ' journal number
If IsNull(recorrr!chkdue) = False Then
txtcheckdate.Value = recorrr!chkdue
'txtchecknumber.Text = recorrr!moderef
txtchecknumber.Text = IIf(IsNull(recorrr!moderef), "", recorrr!moderef)

End If

comprepared.Text = recorrr!cashier
comchecked.Text = recorrr!empcheck
comapproved.Text = recorrr!empapp
txtattachment.Text = recorrr!remarks
'end procedure
            Frame2.Enabled = True
            combdebitaccnum.Text = "111010101001"
                If Trim(Mid(Trim(comsetmode.Text), 1, 3)) = "005" Then
                Dim recpayjor As New ADODB.Recordset ' this is to take accountnumber
                recpayjor.Open "select * from payjournal where serno = '" & Trim(txtinvoicenum.Text) & "'", CON1, adOpenKeyset, adLockOptimistic
                    If recpayjor.BOF = False Then
                        While recpayjor.EOF = False
                            If Val(Trim(recpayjor!DBamount)) > 0 Then
                                'combdebitaccnum.Text = recpayjor!Accno
                            Else
                                comcreditaccountnumber.Text = recpayjor!AccNo
                            End If
                            recpayjor.MoveNext
                        Wend
                    End If
                  recpayjor.close
                End If
            combdebitaccnum.SetFocus
End If
recorrr.close
conforsearch.close
End Sub
Public Sub prcdisplay()
txtreceiptnumber.Text = recor!receiptno
lbldate.caption = recor!receiptdate
comreceivedfrom.Text = recor!custno & "     " & recor!custname
tookcredit = 0
If IsNull(recor!creditamount) Then
    tookcredit = 0
Else
    tookcredit = Trim(recor!creditamount)
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
takejournal = recor!JournalNumber ' journal number
If IsNull(recor!chkdue) = False Then
txtcheckdate.Value = recor!chkdue
End If

comprepared.Text = recor!cashier
comchecked.Text = recor!empcheck
comapproved.Text = recor!empapp
txtattachment.Text = recor!remarks

End Sub

Public Sub prcclear()
'this is for clear all control in frame1
lbldate.caption = Format(Date, "dd/mm/yyyy")
comreceivedfrom.Text = ""
txtdebitamt.Text = ""
lbldebitamount.caption = ""
comsetmode.ListIndex = 0
'txtinvoicenum.Text = ""
cmbpaymenttype.ListIndex = 0
txtchecknumber.Text = ""
txtcheckdate.Value = Null
txtattachment.Text = ""
comreceivedfrom.Text = ""
comreceivedfrom.Text = ""
txtinvoicenum.Text = ""
txtdebitamt.Text = ""
ListView1.ListItems.clear
creditcount = 0
creditamount = 0
creditamountcount = 0
'cmdadd.Enabled = True
' end clear
End Sub

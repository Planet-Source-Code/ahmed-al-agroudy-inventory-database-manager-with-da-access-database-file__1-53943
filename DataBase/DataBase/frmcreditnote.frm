VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmcreditnote 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Credit Note  ÇÔÚÇÑ ÏÇÆä "
   ClientHeight    =   9285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11775
   Icon            =   "frmcreditnote.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9285
   ScaleWidth      =   11775
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   2535
      Left            =   10560
      TabIndex        =   60
      Top             =   3720
      Width           =   1095
      Begin VB.CommandButton cmdshowinvoice 
         Caption         =   "S&how"
         Height          =   375
         Left            =   120
         TabIndex        =   66
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdnewrecord 
         Caption         =   "&New"
         Height          =   375
         Left            =   120
         TabIndex        =   65
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton cmdcancel 
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   350
         Left            =   120
         TabIndex        =   64
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton cmdsave 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   350
         Left            =   120
         TabIndex        =   63
         Top             =   1680
         Width           =   855
      End
      Begin VB.CommandButton cmdedit 
         Caption         =   "&Edit"
         Height          =   350
         Left            =   120
         TabIndex        =   62
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton cmdclose 
         Caption         =   "Cl&ose"
         Height          =   350
         Left            =   120
         TabIndex        =   61
         Top             =   2040
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Payment Analysis"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   3000
      Width           =   11775
      Begin VB.CommandButton cmdadd2 
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
         Left            =   9480
         TabIndex        =   57
         Top             =   480
         Width           =   855
      End
      Begin VB.ComboBox ccname 
         Height          =   315
         Left            =   2520
         Sorted          =   -1  'True
         TabIndex        =   24
         Top             =   2400
         Width           =   4815
      End
      Begin VB.ComboBox combdebitaccnum 
         Height          =   315
         ItemData        =   "frmcreditnote.frx":0442
         Left            =   120
         List            =   "frmcreditnote.frx":0444
         Sorted          =   -1  'True
         TabIndex        =   21
         Top             =   480
         Width           =   2295
      End
      Begin VB.ComboBox cname 
         Height          =   315
         Left            =   2520
         Sorted          =   -1  'True
         TabIndex        =   22
         Top             =   480
         Width           =   4815
      End
      Begin VB.ComboBox comcreditaccountnumber 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   23
         Top             =   2400
         Width           =   2295
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
         Left            =   9480
         TabIndex        =   26
         Top             =   2400
         Width           =   855
      End
      Begin MSMask.MaskEdBox creditamount 
         Height          =   315
         Left            =   7440
         TabIndex        =   25
         Top             =   2400
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "0.00"
         PromptChar      =   " "
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   135
         Left            =   120
         TabIndex        =   45
         Top             =   4080
         Visible         =   0   'False
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   238
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   1260
         Left            =   120
         TabIndex        =   27
         Top             =   2760
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   2223
         View            =   3
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
            Name            =   "Arabic Transparent"
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
            Object.Width           =   3087
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Account Name"
            Object.Width           =   9349
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Amount"
            Object.Width           =   3175
         EndProperty
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   1260
         Left            =   120
         TabIndex        =   58
         Top             =   840
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   2223
         View            =   3
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
            Name            =   "Arabic Transparent"
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
            Object.Width           =   3087
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Account Name"
            Object.Width           =   9349
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Amount"
            Object.Width           =   3175
         EndProperty
      End
      Begin MSMask.MaskEdBox txtdebitamount 
         Height          =   315
         Left            =   7440
         TabIndex        =   59
         Top             =   480
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "0.00"
         PromptChar      =   " "
      End
      Begin VB.Label lblshow 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Please Wait ........."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   9720
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   3960
         Visible         =   0   'False
         Width           =   1650
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Application  æËíÞÉ ÇáÏÝÚ "
         Height          =   195
         Left            =   9360
         TabIndex        =   43
         Top             =   240
         Width           =   2340
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
         Left            =   1200
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   240
         Width           =   1140
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         Caption         =   "ÑÞã ÍÓÇÈ ÇáÏÇÆä "
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
         Left            =   1200
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   2160
         Width           =   1110
      End
      Begin VB.Label Label45 
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
         Left            =   6480
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   2160
         Width           =   810
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         Caption         =   "ÇáÃÌãÇáí "
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
         Left            =   8640
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   2160
         Width           =   600
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "ÇáÃÌãÇáí "
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
         Left            =   8640
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   240
         Width           =   600
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
         Left            =   6480
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   240
         Width           =   810
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         Height          =   195
         Left            =   7440
         TabIndex        =   32
         Top             =   240
         Width           =   540
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Account Name"
         Height          =   255
         Left            =   2520
         TabIndex        =   31
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Debit Number"
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
         TabIndex        =   30
         Top             =   2160
         Width           =   960
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         Height          =   195
         Left            =   7440
         TabIndex        =   29
         Top             =   2160
         Width           =   540
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Account Name"
         Height          =   255
         Left            =   2520
         TabIndex        =   28
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Credit Number"
         Height          =   195
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Width           =   1005
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   0
      TabIndex        =   1
      Top             =   -120
      Width           =   11775
      Begin VB.Timer Timer1 
         Left            =   8520
         Top             =   840
      End
      Begin VB.ComboBox comno 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1080
         TabIndex        =   3
         Top             =   150
         Width           =   1575
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2055
         Left            =   120
         TabIndex        =   33
         Top             =   720
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   3625
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Invoice No"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Record Date     Due Date"
            Object.Width           =   4939
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Amount"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Unpaid"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Applied"
            Object.Width           =   2822
         EndProperty
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "( Tax Credit )"
         Height          =   195
         Left            =   9360
         TabIndex        =   56
         Top             =   2160
         Width           =   900
      End
      Begin VB.Label lbltaxcredit 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label9"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   10320
         TabIndex        =   55
         Top             =   2160
         Width           =   1350
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sur Tax"
         Height          =   195
         Left            =   9360
         TabIndex        =   54
         Top             =   1920
         Width           =   555
      End
      Begin VB.Label lblsurtax 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label9"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   10320
         TabIndex        =   53
         Top             =   1920
         Width           =   1350
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Net Amount "
         Height          =   195
         Left            =   9360
         TabIndex        =   52
         Top             =   1440
         Width           =   885
      End
      Begin VB.Label lblnet 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label9"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   10320
         TabIndex        =   51
         Top             =   1440
         Width           =   1350
      End
      Begin VB.Label lbltransport 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "transport"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   10320
         TabIndex        =   50
         Top             =   1200
         Width           =   1350
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transport"
         Height          =   195
         Left            =   9360
         TabIndex        =   49
         Top             =   1200
         Width           =   675
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice Number:"
         Height          =   195
         Left            =   9360
         TabIndex        =   48
         Top             =   2640
         Width           =   1170
      End
      Begin VB.Label lblinvoice 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label9"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   10320
         TabIndex        =   47
         Top             =   2640
         Width           =   1350
      End
      Begin VB.Label lblname 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label9"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1320
         TabIndex        =   36
         Top             =   480
         Width           =   540
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client Name"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   35
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblamount 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label9"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   10320
         TabIndex        =   16
         Top             =   2400
         Width           =   1350
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount:"
         Height          =   195
         Left            =   9360
         TabIndex        =   15
         Top             =   2400
         Width           =   990
      End
      Begin VB.Label lblvat 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label9"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   10320
         TabIndex        =   14
         Top             =   1680
         Width           =   1350
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "V. A . T"
         Height          =   195
         Left            =   9360
         TabIndex        =   13
         Top             =   1680
         Width           =   540
      End
      Begin VB.Label lbltradediscount 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label9"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   10320
         TabIndex        =   12
         Top             =   960
         Width           =   1350
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "( Trade Discount )"
         Height          =   195
         Left            =   9360
         TabIndex        =   11
         Top             =   960
         Width           =   1275
      End
      Begin VB.Label lbltotalamount 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label9"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   10320
         TabIndex        =   10
         Top             =   720
         Width           =   1350
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Allowances :"
         Height          =   195
         Left            =   9360
         TabIndex        =   9
         Top             =   720
         Width           =   900
      End
      Begin VB.Label lblremarks 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label9"
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
         Left            =   5520
         TabIndex        =   6
         Top             =   480
         Width           =   465
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reason  :"
         Height          =   195
         Index           =   1
         Left            =   4560
         TabIndex        =   34
         Top             =   480
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Credit Note"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date  :"
         Height          =   195
         Left            =   4560
         TabIndex        =   4
         Top             =   150
         Width           =   480
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client Code  :"
         Height          =   195
         Left            =   9360
         TabIndex        =   7
         Top             =   150
         Width           =   945
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Unpaid Amount"
         Height          =   195
         Left            =   4200
         TabIndex        =   19
         Top             =   2800
         Width           =   1500
      End
      Begin VB.Label lbltotal 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "lbltotamount"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6360
         TabIndex        =   20
         Top             =   2800
         Width           =   1710
      End
      Begin VB.Label lblunapplied 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label9"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1680
         TabIndex        =   18
         Top             =   2800
         Width           =   1470
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UnApplied Amount"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   2800
         Width           =   1320
      End
      Begin VB.Label lbldate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label9"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5760
         TabIndex        =   5
         Top             =   150
         Width           =   555
      End
      Begin VB.Label lblcode 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label9"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   10680
         TabIndex        =   8
         Top             =   150
         Width           =   690
      End
   End
   Begin VB.Menu xfile 
      Caption         =   "file"
      Visible         =   0   'False
      Begin VB.Menu edit 
         Caption         =   "Edit"
      End
      Begin VB.Menu Dash 
         Caption         =   "-"
      End
      Begin VB.Menu delete 
         Caption         =   "Delete"
      End
   End
   Begin VB.Menu full 
      Caption         =   "full"
      Visible         =   0   'False
      Begin VB.Menu deletefull 
         Caption         =   "Delete"
      End
   End
End
Attribute VB_Name = "frmcreditnote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim recacc As New ADODB.Recordset  ' this is for the finance master
Dim recinv As New ADODB.Recordset ' this is for the creditmain table
Dim reccust As New ADODB.Recordset
Dim con As New ADODB.Connection
Dim con2 As New ADODB.Connection
Dim con3 As New ADODB.Connection
Public transferamount As Currency
Dim selectclick As Integer
Public validatedebitamount As Currency
Public totalvalidatedebitamount As Currency
Dim creditamountcount As Currency ' thsi is for count the balance
Dim creditcount As Integer
Dim debitamountcount As Currency
Dim debitcount As Integer

Dim amountofcreditnote As Currency
Dim fromwho As String


Private Sub ccname_Click()
Dim recfindacc As New ADODB.Recordset
namenamenum = InStr(1, Trim(ccname.Text), "\", vbTextCompare)
If namenamenum > 0 Then
namename = Mid(Trim(ccname.Text), 1, (namenamenum - 1))
End If
recfindacc.Open "Select * from financemaster where accountnameeng = " & "'" & namename & "'", con, adOpenKeyset, adLockOptimistic
If recfindacc.BOF = False Then
    comcreditaccountnumber.Text = recfindacc!AccountCode
End If
 recfindacc.Close

Dim cc As String
    getmothername nonumber, namename, cc
Me.caption = "Credit Note  " & cc
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
If KeyAscii = 13 And Trim(ccname.Text) <> "" Then
    creditamount.SetFocus
Else
    ccname.SetFocus
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
recfindacc.Open "Select * from financemaster where accountnameeng = " & "'" & namename & "'", con, adOpenKeyset, adLockOptimistic
If recfindacc.BOF = True Then
    recfindacc.Close
    MsgBox "Please choose the Correct Account Number", vbInformation, "Invalid Account Number"
    ccname.SetFocus
    Exit Sub
Else
    comcreditaccountnumber.Text = Trim(recfindacc!AccountCode)
End If

End Sub

Private Sub cmdAdd_Click()
If Trim(comcreditaccountnumber.Text) = "" Or Trim(ccname.Text) = "" Or Val(Trim(creditamount.Text)) <= "0" Then
    MsgBox "Please check the Credit Account Number and Credit Amount", vbInformation, "Empty Data"
    comcreditaccountnumber.SetFocus
    Exit Sub
End If
    
creditamountcount = Val(creditamountcount) + Val(creditamount.Text) 'this is text box
If Val(amountofcreditnote) < Val(creditamountcount) Then
    MsgBox "Please Check Your Payment Amount", vbInformation, "Invalid Amount"
    creditamountcount = Val(creditamountcount) - Val(creditamount.Text)
    creditamount.SetFocus
    Exit Sub
End If

If Val(amountofcreditnote) = Val(creditamountcount) Then
    cmdadd.Enabled = False
    cmdsave.SetFocus
Else
    comcreditaccountnumber.SetFocus
End If
'to check all the values
creditcount = creditcount + 1
ListView2.ListItems.Add , , Trim(comcreditaccountnumber.Text)
ListView2.ListItems(creditcount).ListSubItems.Add , , Trim(ccname.Text)
ListView2.ListItems(creditcount).ListSubItems.Add , , Format(Trim(creditamount.Text), "############0.00#")
If creditcount = 10 Then
    cmdadd.Enabled = False
    cmdsave.SetFocus
End If
End Sub

Private Sub cmdadd2_Click()
If Trim(combdebitaccnum.Text) = "" Or Trim(cname.Text) = "" Or Val(Trim(txtdebitamount.Text)) <= "0" Then
    MsgBox "Please check the Debit Account Number and Credit Amount", vbInformation, "Empty Data"
    combdebitaccnum.SetFocus
    Exit Sub
End If
    
debitamountcount = Val(debitamountcount) + Val(txtdebitamount.Text)  'this is text box
If Val(amountofcreditnote) < Val(debitamountcount) Then
    MsgBox "Please Check Your Payment Amount", vbInformation, "Invalid Amount"
    debitamountcount = Val(debitamountcount) - Val(txtdebitamount.Text)
    txtdebitamount.SetFocus
    Exit Sub
End If

If Val(amountofcreditnote) = Val(debitamountcount) Then
    cmdadd2.Enabled = False
    comcreditaccountnumber.SetFocus
Else
    combdebitaccnum.SetFocus
End If

'to check all the values

debitcount = debitcount + 1
ListView3.ListItems.Add , , Trim(combdebitaccnum.Text)
ListView3.ListItems(debitcount).ListSubItems.Add , , Trim(cname.Text)
ListView3.ListItems(debitcount).ListSubItems.Add , , Format(Trim(txtdebitamount.Text), "############0.00#")
If creditcount = 10 Then
    cmdadd2.Enabled = False
    comcreditaccountnumber.SetFocus
End If

End Sub

Private Sub cmdcancel_Click()
On Error Resume Next
Frame1.Enabled = False
Frame2.Enabled = False
cmdedit.Enabled = True
cmdsave.Enabled = False
cmdprint.Enabled = False
cmdclose.Enabled = True
cmdnewrecord.Enabled = True
cmdcancel.Enabled = False
ListView1.ListItems.Clear
ListView2.ListItems.Clear
ListView3.ListItems.Clear

For Each Control In Me
    If TypeOf Control Is ComboBox Then
        Control.Text = " "
    End If
Next
txtdebitamount.Text = " "
lbltradediscount.caption = " "
lblinvoice.caption = " "
lbltransport.caption = "  "
amountofcreditnote = 0
creditamount.Text = " "
lblcode.caption = " "
lblamount.caption = " "
lblremarks.caption = " "
lblname.caption = " "
lblvat.caption = " "
lbltotalamount.caption = " "
validatedebitamount = 0
Timer1.Interval = 0
lblunapplied.caption = " "
lblnet.caption = " "
lblsurtax.caption = " "
lbltaxcredit.caption = " "
comno.Enabled = False
Frame1.Enabled = False
Frame2.Enabled = False
Me.caption = "Credit Note"
End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub CMDEDIT_Click()
comno.Enabled = True
fromwho = "a"
ListView1.ListItems.Clear
ListView2.ListItems.Clear
ListView3.ListItems.Clear
cmdedit.Enabled = False
cmdadd.Enabled = True
cmdadd2.Enabled = True
cmdsave.Enabled = True
cmdclose.Enabled = False
cmdnewrecord.Enabled = False
Frame1.Enabled = True
comno.Enabled = True
creditcount = 0
creditamount = 0
creditamountcount = 0
cmdcancel.Enabled = True
comno.SetFocus
Call prcclear
End Sub

Private Sub cmdnewrecord_Click()
comno.Enabled = True
Frame1.Enabled = True
Frame2.Enabled = False
ListView1.ListItems.Clear
ListView2.ListItems.Clear
ListView3.ListItems.Clear
comno.Text = " "
validatedebitamount = 0
cmdcancel.Enabled = True
comno.Enabled = True
cmdsave.Enabled = True
cmdnewrecord.Enabled = False
cmdedit.Enabled = False
cmdclose.Enabled = False
fromwho = "c"
comno.SetFocus
Call prcclear
End Sub

Private Sub cmdsave_Click()
ProgressBar1.Max = ListView1.ListItems.Count + ListView2.ListItems.Count + ListView3.ListItems.Count + 1

'this is for payment analysis
If fromwho = "a" Then ' ListView2.ListItems.Count > 0 And Trim(comno.Text) <> "" Then
    
ProgressBar1.Visible = True
lblshow.Visible = True

If ListView2.Enabled = True Then
    ListCount = ListView2.ListItems.Count
    i = 1
    checkallamount = 0
    For i = 1 To ListCount
        checkallamount = checkallamount + Val(Trim(ListView2.ListItems(i).ListSubItems(2).Text))
    Next
    i = 1
    checkallamount2 = 0
    For i = 1 To ListView3.ListItems.Count
        checkallamount2 = checkallamount2 + Val(Trim(ListView3.ListItems(i).ListSubItems(2).Text))
    Next
    
    If Val(checkallamount) <> Val(Trim(amountofcreditnote)) Or Val(checkallamount2) <> Val(Trim(amountofcreditnote)) Then
        MsgBox "Please Check Credit Amount it is not Equal to Debit Amount", vbInformation, "Amount Overflow"
        checkallamount = 0
        checkallamount2 = 0
        Exit Sub
    End If
    
    'this is for clear the creditnote table and cashjournal table *****************
    Dim deletedata As New ADODB.Recordset
    deletedata.Open "delete from creditnote where status = 'UnPosted' and creditnoteno = '" & Trim(comno.Text) & "'", con, adOpenKeyset, adLockOptimistic

    Dim recsales As New ADODB.Recordset ' this is for salesreturnjounal
    recsales.Open "select * from creditnote", con, adOpenKeyset, adLockOptimistic
       
   'check for the month and year change if we have to reset the series number.
    Dim recset As New ADODB.Recordset
    Dim checkjournaldate As Date
    
    recset.Open "Select getdate() as currentmoyr", con, adOpenKeyset, adLockOptimistic
    checkjournaldate = Left(Trim(recset!CurrentMoYr), 10)
    recset.Close
    
    recset.Open "Select * from setup", con, adOpenKeyset, adLockOptimistic
    recset.MoveFirst
    If Trim(recset!CurrentMoYr) <> Format(checkjournaldate, "mmyy") Then
        recset!CurrentMoYr = Format(checkjournaldate, "mmyy")
        recset!nextjn = "0001"
        recset.Update
    End If
    recset.Requery
    
    takejournal = "SRL-" & recset!CurrentMoYr & "-" & Trim(recset!nextjn)
    totaljournal = "0000" & (Val(Trim(recset!nextjn)) + 1)
    recset!nextjn = Right(totaljournal, 5)
    recset.Update
    'end journal number

    ' this is for listview
        xxx3 = 1
        For xxx3 = 1 To ListView3.ListItems.Count
            recsales.AddNew
            recsales!SerialNo = takejournal
            recsales!creditnoteno = Trim(comno.Text)
            recsales!ticket = xxx3
            recsales!deletemark = "0"
            recsales!accountnumber = Trim(ListView3.ListItems(xxx3).Text)
                namenamenum = InStr(1, Trim(ListView3.ListItems(xxx3).ListSubItems(1).Text), "\", vbTextCompare)
                If namenamenum > 0 Then
                namename = Mid(Trim(ListView3.ListItems(xxx3).ListSubItems(1).Text), 1, (namenamenum - 1))
                Else
                namename = Trim(ListView3.ListItems(xxx3).ListSubItems(1).Text)
                End If
            recsales!accountname = Trim(namename)
            recsales!accountnamearab = Mid(Trim(ListView3.ListItems(xxx3).ListSubItems(1).Text), (namenamenum + 1), Val(Len(Trim(ListView3.ListItems(xxx3).ListSubItems(1).Text)) - Val(namenamenum)))
            'this is for find the mother name
            getmothername Trim(ListView3.ListItems(xxx3).Text), noname, cc
            Me.caption = "Credit Note  " & cc
            recsales!mothername = cc
            recsales!Description = Trim(ListView2.ListItems(1).Text)  'againts name
            recsales!TRansDate = Format(Trim(lbldate.caption), "mm/dd/yyyy")
            recsales!DebitAmount = 0
            recsales!creditamount = FormatNumber(Trim(ListView3.ListItems(xxx3).ListSubItems(2).Text), 4)
            recsales!reasons = Trim(lblremarks.caption)
            recsales!Status = "UnPosted"
            recsales!Trantype = "SRL"
            recsales.Update
            
            ProgressBar1.Visible = True
            lblshow.Visible = True
            ProgressBar1.Min = 0
            ProgressBar1.Max = ListView3.ListItems.Count
            If ProgressBar1.Value <> ProgressBar1.Max Then
                ProgressBar1.Value = ProgressBar1.Value + 1
            Else
                ProgressBar1.Value = 0
            End If
        Next
    'end listview3
    
    ' this is for listview2 ==== debit amount
        For xxx = 1 To ListView2.ListItems.Count
            recsales.AddNew
            recsales!SerialNo = takejournal
            recsales!creditnoteno = Trim(comno.Text)
            recsales!ticket = xxx + (xxx3 - 1) ' the for loop add one more value to the xx3 so minus one
            recsales!deletemark = "0"
            recsales!accountnumber = Trim(ListView2.ListItems(xxx).Text)
                namenamenum = InStr(1, Trim(ListView2.ListItems(xxx).ListSubItems(1).Text), "\", vbTextCompare)
                If namenamenum > 0 Then
                namename = Mid(Trim(ListView2.ListItems(xxx).ListSubItems(1).Text), 1, (namenamenum - 1))
                Else
                namename = Trim(ListView2.ListItems(xxx).ListSubItems(1).Text)
                End If
            recsales!accountname = Trim(namename)
            recsales!accountnamearab = Mid(Trim(ListView2.ListItems(xxx).ListSubItems(1).Text), (namenamenum + 1), Val(Len(Trim(ListView2.ListItems(xxx).ListSubItems(1).Text)) - Val(namenamenum)))
            'this is for find the mother name
            getmothername Trim(combdebitaccnum.Text), noname, cc
            Me.caption = "Credit Note  " & cc
            recsales!mothername = cc
            recsales!Description = Trim(ListView3.ListItems(1).Text)  'againts name
            recsales!TRansDate = Format(Trim(lbldate.caption), "mm/dd/yyyy")
            recsales!DebitAmount = FormatNumber(Trim(ListView2.ListItems(xxx).ListSubItems(2).Text), 4)
            recsales!creditamount = 0
            recsales!reasons = Trim(lblremarks.caption)
            recsales!Status = "UnPosted"
            recsales!Trantype = "SRL"
            recsales.Update
            
            ProgressBar1.Visible = True
            lblshow.Visible = True
            ProgressBar1.Min = 0
            ProgressBar1.Max = ListView2.ListItems.Count
            If ProgressBar1.Value <> ProgressBar1.Max Then
                ProgressBar1.Value = ProgressBar1.Value + 1
            Else
                ProgressBar1.Value = 0
            End If
        Next
    End If
On Error Resume Next
dataanu.rscom_creditnote_byjounal_Grouping.Close
On Error GoTo 0
dataanu.com_creditnote_byjounal_Grouping takejournal
re_creditnote_byjournal.Show
re_creditnote_byjournal.PrintReport False, rptRangeAllPages
End If

'**************************************'this is for new data
If fromwho = "a" Or fromwho = "c" Then ' ListView1.ListItems.Count > 0 Then
Dim recte As New ADODB.Recordset
recte.Open "Select * from tempagaintsinvoice where receiptno = " & "'" & Trim(comno.Text) & "'", con, adOpenKeyset, adLockOptimistic

Dim recteinvoice12 As New ADODB.Recordset 'this is for add tempagaintsinvoice
Dim rectee As New ADODB.Recordset
rectee.Open "tempinvoice2", con, adOpenKeyset, adLockOptimistic ' tempinvoice2 table
recteinvoice12.Open "tempagaintsinvoice", con, adOpenKeyset, adLockOptimistic
i = 1
While ListView1.ListItems.Count >= i
        If Val(Trim(ListView1.ListItems(i).ListSubItems(4).Text)) > 0 Then
            recte.Requery
            If recte.BOF = True Then 'this is to add menu
                recteinvoice12.AddNew
                    recteinvoice12!receiptno = Trim(comno.Text)
                    recteinvoice12!invoiceno = "Invoice Number"
                    recteinvoice12!Menu = "Amount"
                    recteinvoice12!svoucher = "CreditNote"
                    recteinvoice12!custno = Trim(lblcode.caption)
                recteinvoice12.Update
            End If
                 recteinvoice12.AddNew ' this is for tempagaintsinvoice
                     recteinvoice12!receiptno = Trim(comno.Text)
                     recteinvoice12!invoiceno = Trim(ListView1.ListItems(i).Text)
                     recteinvoice12!Applied = Trim(ListView1.ListItems(i).ListSubItems(4).Text)
                     unappliedbalance = Trim(lblunapplied.caption)
                     recteinvoice12!svoucher = "CreditNote"
                     recteinvoice12!display = Trim(Right(Trim(ListView1.ListItems(i).ListSubItems(1).Text), 2))
                     recteinvoice12!invoicedate = Mid(Trim(ListView1.ListItems(i).ListSubItems(1).Text), 1, 12)
                     recteinvoice12!custno = Trim(lblcode.caption)
                     notequla = 2 ' this is for unapplied amount
                recteinvoice12.Update
            'this is for tempinvoice2
                rectee.AddNew
                    rectee!receiptno = Trim(comno.Text)
                    rectee!custid = Trim(lblcode.caption)
                    rectee!invoiceno = Trim(ListView1.ListItems(i).Text)
                    rectee!receiptdate = Trim(ListView1.ListItems(i).ListSubItems(1).Text)
                    rectee!amount = Val(lblamount.caption)
                    rectee!unpaid = (lblunapplied.caption)
                    rectee!Applied = Trim(ListView1.ListItems(i).ListSubItems(4).Text)
                rectee.Update
            ' this is to update the sjmaster table
                 Dim recsjmaster As New ADODB.Recordset
                 recsjmaster.Open "select * from sjmaster where invc_no = '" & Trim(ListView1.ListItems(i).Text) & "'", con2, adOpenKeyset, adLockOptimistic
                 If recsjmaster.BOF = False Then
                     recsjmaster!unpaidamt = Trim(ListView1.ListItems(i).ListSubItems(3).Text)
                     recsjmaster!paidamt = Val(Trim(recsjmaster!paidamt)) + Val(Trim(ListView1.ListItems(i).ListSubItems(4).Text))
                     recsjmaster.Update
                 End If
                recsjmaster.Close
           'end update the foxpro sjmaster
        End If
        ProgressBar1.Visible = True
        lblshow.Visible = True
        ProgressBar1.Min = 0
        ProgressBar1.Max = ListView1.ListItems.Count
        If ProgressBar1.Value <> ProgressBar1.Max Then
            ProgressBar1.Value = ProgressBar1.Value + 1
        Else
            ProgressBar1.Value = 0
        End If
i = i + 1
Wend
rectee.Close
Dim recclear As New ADODB.Recordset
recclear.Open "delete from tempagaintsinvoice where receiptno = " & "'" & Trim(comno.Text) & "'" & " and svoucher = 'CreditNote' and (invoiceNO='Invoice SubTotal' or invoiceno = 'Un Applied Amount')", con, adOpenKeyset, adLockOptimistic

    recteinvoice12.AddNew  ' to record total amount
        recteinvoice12!receiptno = Trim(comno.Text)
        recteinvoice12!invoiceno = "Invoice SubTotal"
        recteinvoice12!Applied = Val(lblamount.caption) - Val(unappliedbalance)
        recteinvoice12!svoucher = "CreditNote"
        recteinvoice12!custno = Trim(lblcode.caption)
        recteinvoice12.Update
        
    If unappliedbalance > 0 Then
         recteinvoice12.AddNew ' to record un applied amount
            recteinvoice12!receiptno = Trim(comno.Text)
            recteinvoice12!invoiceno = "Un Applied Amount"
            recteinvoice12!Applied = unappliedbalance
            recteinvoice12!svoucher = "CreditNote"
            recteinvoice12!custno = Trim(lblcode.caption)
            notequla = 1
        recteinvoice12.Update
        
    recteinvoice12.Close
    End If
End If  ' if listview1.listitem greater than the zero

'this is the common things
MsgBox "Your Data Saved Successfully", vbInformation, "Save Conformation"
Frame1.Enabled = False
Frame2.Enabled = False
cmdsave.Enabled = False
cmdcancel.Enabled = False
cmdclose.Enabled = True
cmdedit.Enabled = True
debitamountcount = 0
creditamount = 0
creditamountcount = 0
debitcount = 0
cmdnewrecord.Enabled = True
ProgressBar1.Visible = False
lblshow.Visible = False
txtdebitamount.Text = " "
creditamount.Text = " "
Me.caption = "Credit Note"
comno.Enabled = False
End Sub

Private Sub cmdshowinvoice_Click()
If amountofcreditnote > 0 Then
    frmcreditinvoice.anucustcode = lblcode.caption
    frmcreditinvoice.anucustname = lblname.caption
    frmcreditinvoice.transferamount = FormatNumber(amountofcreditnote, "############0.#0")
    frmcreditinvoice.receiptno = Trim(comno.Text)
    frmcreditinvoice.Show 1
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
recfindacc.Open "Select * from financemaster where accountnameeng = " & "'" & namename & "'", con, adOpenKeyset, adLockOptimistic
If recfindacc.BOF = False Then
    combdebitaccnum.Text = recfindacc!AccountCode
End If
 recfindacc.Close
  
   getmothername nonumber, namename, cc
Me.caption = "Credit Note  " & cc
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
    txtdebitamount.SetFocus
End If
End Sub

Private Sub cname_LostFocus()
Dim recfindacc As New ADODB.Recordset
namenamenum = InStr(1, Trim(cname.Text), "\", vbTextCompare)
If namenamenum > 0 Then
namename = Mid(Trim(cname.Text), 1, (namenamenum - 1))
End If
recfindacc.Open "Select * from financemaster where accountnameeng = " & "'" & namename & "'", con, adOpenKeyset, adLockOptimistic
If recfindacc.BOF = True Then
    MsgBox "Please choose the Currect Account Number", vbInformation, "Invalid accountNumber"
    Exit Sub
Else
    combdebitaccnum.Text = recfindacc!AccountCode
End If
recfindacc.Close

End Sub

Private Sub combdebitaccnum_Click()
'this is for find the account name
Dim recfindaccount As New ADODB.Recordset
recfindaccount.Open "select * from financemaster where accountcode = " & "'" & Trim(combdebitaccnum.Text) & "'", con, adOpenKeyset, adLockOptimistic
If recfindaccount.BOF = False Then
    cname.Text = Trim(recfindaccount!accountnameeng) & "\" & Trim(recfindaccount!accountnamearab)
End If
recfindaccount.Close
'end find the account name

 Dim cc As String
    Dim ccc As String
    getmothername Trim(combdebitaccnum.Text), ccc, cc
Me.caption = "Credit Note  " & cc
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
recfindaccount.Open "select * from financemaster where accountcode = " & "'" & Trim(combdebitaccnum.Text) & "'", con, adOpenKeyset, adLockOptimistic
If recfindaccount.BOF = True Then
    MsgBox "Please Choose the Sub Account", vbInformation, "Invalid AccountNumber"
    combdebitaccnum.SetFocus
    Exit Sub
Else
        cname.Text = Trim(recfindaccount!accountnameeng) & "\" & Trim(recfindaccount!accountnamearab)
End If
recfindaccount.Close
End If
End Sub

Private Sub comcreditaccountnumber_Click()
Dim recfindaccount As New ADODB.Recordset
recfindaccount.Open "select * from financemaster where accountcode = " & "'" & Trim(comcreditaccountnumber.Text) & "'", con, adOpenKeyset, adLockOptimistic
If recfindaccount.BOF = False Then
    ccname.Text = Trim(recfindaccount!accountnameeng) & "\" & Trim(recfindaccount!accountnamearab)
End If
recfindaccount.Close
    getmothername Trim(comcreditaccountnumber.Text), noname, cc
Me.caption = "Credit Note  " & cc

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
recfindaccount.Open "select * from financemaster where accountcode = " & "'" & Trim(comcreditaccountnumber.Text) & "'", con, adOpenKeyset, adLockOptimistic
If recfindaccount.BOF = True Then
    MsgBox "Please Choose the Sub Account", vbInformation, "Invalid AccountNumber"
    comcreditaccountnumber.SetFocus
    Exit Sub
Else
        ccname.Text = Trim(recfindaccount!accountnameeng) & "\" & Trim(recfindaccount!accountnamearab)
End If
recfindaccount.Close
End If
End Sub


Private Sub creditamount_GotFocus()
   creditamount.Text = " "
End Sub

Private Sub creditamount_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And cmdadd.Enabled = True Then
    cmdadd.SetFocus
End If
If KeyAscii = 13 And cmdadd.Enabled = False Then
cmdsave.SetFocus
End If

End Sub

Private Sub txtdebitamount_GotFocus()
   txtdebitamount.Text = " "
End Sub

Private Sub txtdebitamount_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And cmdadd2.Enabled = True Then
    cmdadd2.SetFocus
End If
If KeyAscii = 13 And cmdadd2.Enabled = False Then
comcreditaccountnumber.SetFocus
End If

End Sub

Private Sub deletefull_Click()
On Error GoTo er
If ListView2.SelectedItem.Index > 0 Then
    creditamountcount = creditamountcount - Val(ListView2.SelectedItem.SubItems(2))
    deleteindex = ListView2.SelectedItem.Index
    ListView2.ListItems.Remove (deleteindex)
    cmdadd.Enabled = True
    creditcount = creditcount - 1
End If

er:
On Error GoTo X
If ListView3.SelectedItem.Index > 0 Then
    debitamountcount = debitamountcount - Val(ListView3.SelectedItem.SubItems(2))
    deleteindex = ListView3.SelectedItem.Index
    ListView3.ListItems.Remove (deleteindex)
    cmdadd2.Enabled = True
    debitcount = debitcount - 1
End If
X:
End Sub

Private Sub edit_Click()
ListView1_DblClick
End Sub

Private Sub Form_Activate()
On Error Resume Next
If ListView1.Enabled = False Then
cmdnewrecord.SetFocus
End If
On Error GoTo 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
con.Close
con2.Close
con3.Close
End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = vbRightButton Then
    ListView1_Click
    If Val(ListView1.ListItems(selectclick).ListSubItems(4).Text) > 0 Then
        Delete.Enabled = True
    End If
    PopupMenu xfile, vbAlignRight, (ListView1.SelectedItem.Left + 6000), (ListView1.SelectedItem.Top + ListView1.Top + 250)
End If
End Sub
Private Sub comno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Trim(comno.Text) <> "" Then
    Dim reccheckinjournal As New ADODB.Recordset
    reccheckinjournal.Open "select * from creditnote where CreditNoteno = '" & Trim(comno.Text) & "'", con, adOpenKeyset, adLockOptimistic
    'if payment analyze is found then
    If reccheckinjournal.BOF = False And fromwho = "a" Then
        If Trim(reccheckinjournal!Status) = "Posted" Then
            MsgBox "This Transaction is Posted", vbInformation, "Roll Back"
            cmdcancel_Click
            Frame1.Enabled = False
            Frame2.Enabled = False
            comno.Enabled = False
            Exit Sub ' stop here
        End If
         If MsgBox("Are Your Sure Your Want to Analyse Again ?", vbYesNo + vbQuestion + vbDefaultButton2, "Conformation") = vbYes Then
            Frame2.Enabled = True
            'this is mean he like to analyse again
         Else
            cmdcancel_Click
            Frame1.Enabled = False
            Frame2.Enabled = False
            comno.Enabled = False
            Exit Sub ' stop here
        End If
    End If
    reccheckinjournal.Close
    On Error Resume Next
    recinv.Close
    On Error GoTo 0
    
    recinv.Open "Select * from credmain where invc_no = " & "'" & Trim(comno.Text) & "'" & " order by invc_date", con3, adOpenKeyset, adLockOptimistic
        
        If recinv.BOF = True Then  'if it is empty
            MsgBox "No Any Credit Notes to Pay", vbInformation, "Empty"
            Exit Sub
        End If
    Frame1.Enabled = True
    On Error Resume Next
    lbldate.caption = recinv!Trans_dt
    lblcode.caption = recinv!cust_code
    lblamount.caption = Format(recinv!tot_amt, "############0.#0")
    lbltransport.caption = FormatNumber(recinv!transchg, 2)
    lblinvoice.caption = Trim(recinv!oinvc_no)
    lblremarks.caption = recinv!rem1
    lbltotalamount = FormatNumber(recinv!Sub_Amt, 2)
    lbltradediscount = FormatNumber(recinv!Discount, 2)
    lblvat.caption = FormatNumber(recinv!Tot_Vat, 2)
    lblnet.caption = FormatNumber(recinv!tot_amt - recinv!Discount + recinv!transchg, 2)
    lblsurtax.caption = FormatNumber(recinv!surcharge, 2)
    lbltaxcredit.caption = "0.00"
    anucustcode = Trim(recinv!cust_code)
    On Error GoTo 0
    
    Dim rectempinvoice As New ADODB.Recordset 'to check whether there is unapplied amount or not
    rectempinvoice.Open "Select * from tempagaintsinvoice where invoiceno = 'Un Applied Amount' and receiptno= " & "'" & Trim(comno.Text) & "'", con, adOpenKeyset, adLockOptimistic
    Dim rectempinvoicefeedetails As New ADODB.Recordset
    rectempinvoicefeedetails.Open "Select * from tempagaintsinvoice where receiptno = " & "'" & Trim(comno.Text) & "'", con, adOpenKeyset, adLockOptimistic
    
    If rectempinvoice.BOF = True And rectempinvoicefeedetails.BOF = True Then
        validatedebitamount = Val(Format(recinv!tot_amt, "############0.#0")) - Val(recinv!amt_paid)
    Else
        If rectempinvoice.BOF = False Then
            validatedebitamount = rectempinvoice!Applied
        Else
            validatedebitamount = 0
        End If
    End If
    rectempinvoicefeedetails.Close
    rectempinvoice.Close
    'end tempagaintsinvoice table to search the un applied amounts
    
    reccust.Open "select * from marcusfl where cust_code = " & "'" & anucustcode & "'", con3, adOpenKeyset, adLockOptimistic
    lblname.caption = reccust!first_name
    anumaincode = Trim(reccust!mcustcode)
    reccust.Close
    
    'this is to add the invoice details
    Dim recdelfin As New ADODB.Recordset ' this is from sjmaster data
    If anumaincode <> "" Then
         recdelfin.Open "select * from SJMASTER where mcustcode = " & "'" & anumaincode & "'" & " and unpaidamt > 0 order by delv_date,invc_date", con2, adOpenKeyset, adLockOptimistic
    Else
        recdelfin.Open "select * from SJMASTER where mcustcode = " & "'" & anucustcode & "'" & " and unpaidamt > 0 order by delv_date,invc_date", con2, adOpenKeyset, adLockOptimistic
    End If
    If recdelfin.BOF = False Then
        comno.Enabled = False
        recdelfin.MoveFirst
        Timer1.Interval = 600
        ListView1.Enabled = True
        ListView1.ListItems.Clear
    End If
    xx = 0
    While recdelfin.EOF = False
        If Val(Trim(recdelfin!unpaidamt)) > 0 Then
            xx = xx + 1
            ListView1.ListItems.Add , , Trim(recdelfin!invc_no)
            Dim invc_date As Date, delv_date
            invc_date = Trim(recdelfin!invc_date)
            delv_date = Trim(recdelfin!delv_date)
            If recdelfin!LDIsp = False Then
                idisp = "F"
            Else
                idisp = "T"
            End If
            ListView1.ListItems(xx).ListSubItems.Add , , Format(invc_date, "dd/mm/yyyy") & "          " & Format(delv_date, "dd/mm/yyyy") & "             " & idisp
            ListView1.ListItems(xx).ListSubItems.Add , , Format(Trim(recdelfin!tot_amt), "###########0.#0")
            ListView1.ListItems(xx).ListSubItems.Add , , Format(Trim(recdelfin!unpaidamt), "###########0.#0")
            ListView1.ListItems(xx).ListSubItems.Add , , " " ' this is for applied value
        End If
        recdelfin.MoveNext
        Wend
    recdelfin.Close
    cmdsave.Enabled = True
    cmdcancel.Enabled = True
    If fromwho = "a" Then
        Frame2.Enabled = True
        amountofcreditnote = Format(recinv!tot_amt, "############0.#0")
    End If
    On Error Resume Next
    ListView1.SetFocus
    On Error GoTo 0
    recinv.Close
End If
 If fromwho = "a" Then
       ' this is to identify the accountnumber for the client code
     Dim conmarcus As New ADODB.Connection
     conmarcus.Mode = adModeShareDenyNone
     conmarcus.Open "Dsn=anufoxpro;uid=sa;pwd=;"
     Dim recmacmac12 As New ADODB.Recordset
     'this is for client code when main code is null
     recmacmac12.Open "Select * from marcusfl where cust_code = " & "'" & Trim(lblcode.caption) & "'", conmarcus, adOpenKeyset, adLockOptimistic
                 If recmacmac12.BOF = False Then
                     If Trim(recmacmac12!mcustcode) <> "" Then ' if there is main code then
                         takemaincode = Trim(recmacmac12!mcustcode)
                     Else
                         If Trim(recmacmac12!acctNo) <> "" Then
                             combdebitaccnum.Text = recmacmac12!acctNo
                             findaccountnumber = 1
                         End If
                     End If
                 End If
    On Error Resume Next
    recmacmac12.Close
    On Error GoTo 0
    'end take account number
 combdebitaccnum.SetFocus
 End If

End Sub

Private Sub Form_Load()
Me.Height = 7815
con.Open "Dsn=Finance;Uid=Sa;Pwd=;"
con2.Mode = adModeShareDenyNone
con2.Open "Dsn=anufoxpro;uid=sa;pwd=;"

con3.Open "Dsn=anufoxpro;uid=sa;pwd=;"

recinv.Open "Select * from credmain where trim(left(invc_no,3)) = 'OCN' order by invc_date", con3, adOpenKeyset, adLockOptimistic
    While recinv.EOF = False
        comno.AddItem Trim(recinv!invc_no)
        recinv.MoveNext
    Wend
recinv.Close
lbldate.caption = " "
Call prcclear
Dim constring As String
Dim myclass As New HabitatClass
Dim sqltable As Boolean
Dim xtable As String
Dim conrecacc As New ADODB.Connection

constring = "Dsn=Finance;Uid=Sa;Pwd=;"
xtable = "select * from financemaster where active <> '0'"
sqltable = True
myclass.GetTables recacc, conrecacc, xtable, constring, sqltable

    'check for the month and year change if we have to reset the series number.
    Dim recset As New ADODB.Recordset
    recset.Open "Select getdate() as currentmoyr", con, adOpenKeyset, adLockOptimistic
    Dim checkjournaldate As Date
    checkjournaldate = Left(Trim(recset!CurrentMoYr), 10)
    recset.Close
    recset.Open "Select * from setup", con, adOpenKeyset, adLockOptimistic
    recset.MoveFirst
    If Trim(recset!CurrentMoYr) <> Format(checkjournaldate, "mmyy") Then
        recset!CurrentMoYr = Format(checkjournaldate, "mmyy")
        recset!nextjn = "0001"
        recset.Update
    End If
    'end journal number

recacc.MoveFirst
combdebitaccnum.Clear
comcreditaccountnumber.Clear
cname.Clear
ccname.Clear

Dim recaccadd As New ADODB.Recordset
'this is for financemaster
While recacc.EOF = False
    If Mid(Trim(recacc!AccountCode), 1, 3) = "111" Or Mid(Trim(recacc!AccountCode), 1, 3) = "113" Or Mid(Trim(recacc!AccountCode), 1, 3) = "116" Or Mid(Trim(recacc!AccountCode), 1, 3) = "117" Or Mid(Trim(recacc!AccountCode), 1, 3) = "112" Then
        combdebitaccnum.AddItem Trim(recacc!AccountCode)
        cname.AddItem Trim(recacc!accountnameeng) & "\" & Trim(recacc!accountnamearab)
    End If
        comcreditaccountnumber.AddItem Trim(recacc!AccountCode)
        ccname.AddItem Trim(recacc!accountnameeng) & "\" & Trim(recacc!accountnamearab)
    recacc.MoveNext
Wend
recacc.Close
conrecacc.Close
End Sub
Private Sub ListView1_DblClick()
If ListView1.ListItems.Count > 0 Then
frmpaymentforcredit.Show 1
If transferamount > Val(ListView1.ListItems(selectclick).ListSubItems(3).Text) Then
    MsgBox "Your Payment Amount Can not More Than the Unpaid Amount", vbInformation, "Invalid Amount"
    Exit Sub
End If

' totalvalidatedebitamount is to calculate the total debit when you press the enter
totalvalidatedebitamount = totalvalidatedebitamount + transferamount
'validatedebitamount is the total amount from cashier
If totalvalidatedebitamount > validatedebitamount Then ' validatedebitamount  is the first amount
    MsgBox "Please check Your Amount is More the Applicable Amount", vbInformation, "Amount Over flow"
    totalvalidatedebitamount = totalvalidatedebitamount - transferamount
    Exit Sub
End If
If transferamount > 0 Then 'this amount directly from the payment table
    ListView1.ListItems(selectclick).ListSubItems(4).Text = Format(Val(transferamount) + Val(ListView1.ListItems(selectclick).ListSubItems(4).Text), "###########0.#0")
    ListView1.ListItems(selectclick).ListSubItems(3).Text = Format(Val(ListView1.ListItems(selectclick).ListSubItems(3).Text) - Val(transferamount), "###########0.#0")
    'ListView1.ListItems(selectclick).ListSubItems(5).Text = Format(Val(ListView1.ListItems(selectclick).ListSubItems(3).Text), "###########0.#0")
End If
End If
End Sub

Private Sub delete_Click()
totalvalidatedebitamount = totalvalidatedebitamount - Val(Trim(ListView1.ListItems(selectclick).ListSubItems(4).Text))
ListView1.ListItems(selectclick).ListSubItems(3).Text = Format(Val(ListView1.ListItems(selectclick).ListSubItems(3).Text) + Val(Trim(ListView1.ListItems(selectclick).ListSubItems(4).Text)))
ListView1.ListItems(selectclick).ListSubItems(4).Text = " "
'ListView1.ListItems(selectclick).ListSubItems(5).Text = Format(Val(ListView1.ListItems(selectclick).ListSubItems(3).Text), "###########0.#0")

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
End Sub
Private Sub ListView1_Click()
If ListView1.ListItems.Count > 0 Then
If ListView1.ListItems.Count > 0 Then
selectclick = ListView1.SelectedItem.Index
End If
End If
End Sub

Private Sub ListView1_KeyUp(KeyCode As Integer, Shift As Integer)
If ListView1.ListItems.Count > 0 Then
ListView1_Click

If KeyCode = 45 Then
    ListView1_DblClick
End If
If KeyCode = 46 And Val(ListView1.ListItems(selectclick).ListSubItems(4).Text) > 0 Then
    delete_Click
End If
End If
End Sub

Private Sub ListView2_KeyUp(KeyCode As Integer, Shift As Integer)
If ListView2.ListItems.Count > 0 Then
If KeyCode = 46 Then
    deletefull_Click
End If
End If
End Sub

Private Sub ListView2_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
If ListView2.ListItems.Count > 0 And Button = vbRightButton Then
    PopupMenu full
End If

End Sub
Private Sub ListView3_KeyUp(KeyCode As Integer, Shift As Integer)
If ListView3.ListItems.Count > 0 Then
If KeyCode = 46 Then
    deletefull_Click
End If
End If
End Sub

Private Sub ListView3_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
If ListView3.ListItems.Count > 0 And Button = vbRightButton Then
    PopupMenu full
End If
End Sub
Private Sub Timer1_Timer()
totallistcount = 0
totallistcount = ListView1.ListItems.Count
totalvalidatedebitamount = 0
For ListCount = 1 To totallistcount
    totalbalance = totalbalance + Val(Trim(ListView1.ListItems(ListCount).ListSubItems(3).Text))
    totalvalidatedebitamount = totalvalidatedebitamount + Val(Trim(ListView1.ListItems(ListCount).ListSubItems(4).Text))
Next
lbltotal.caption = Format(totalbalance, "###,###,###,##0.#0")


lblunapplied.caption = Format(Val(validatedebitamount - totalvalidatedebitamount), "###########0.#0")
totalbalance = 0
'Timer1.Interval = 0
End Sub

Private Sub prcclear()
lblinvoice.caption = " "
lbltransport.caption = " "
lblamount.caption = " "
lblcode.caption = " "
lblname.caption = " "
lbltotal.caption = " "
lblunapplied.caption = " "
lblremarks.caption = " "
lblnet.caption = " "
lblsurtax.caption = " "
lbltaxcredit.caption = " "
lblvat.caption = " "
lbltradediscount.caption = " "
lbltotalamount.caption = " "
combdebitaccnum.Text = " "
comcreditaccountnumber.Text = " "
cname.Text = " "
ccname.Text = "  "
End Sub

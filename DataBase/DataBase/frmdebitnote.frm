VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmdebitnote 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Debit Note"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10215
   Icon            =   "frmdebitnote.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   10215
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdcancel 
      Caption         =   "&Cancel"
      Enabled         =   0   'False
      Height          =   345
      Left            =   9300
      TabIndex        =   34
      Top             =   4560
      Width           =   855
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   345
      Left            =   9300
      TabIndex        =   33
      Top             =   4080
      Width           =   855
   End
   Begin VB.CommandButton cmdedit 
      Caption         =   "&Edit"
      Height          =   345
      Left            =   9300
      TabIndex        =   32
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Cl&ose"
      Height          =   345
      Left            =   9300
      TabIndex        =   31
      Top             =   5040
      Width           =   855
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
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   2040
      Width           =   10215
      Begin VB.ComboBox ccname 
         Height          =   315
         Left            =   2520
         Sorted          =   -1  'True
         TabIndex        =   20
         Top             =   1080
         Width           =   4815
      End
      Begin VB.ComboBox combdebitaccnum 
         Height          =   315
         ItemData        =   "frmdebitnote.frx":0442
         Left            =   120
         List            =   "frmdebitnote.frx":0444
         Sorted          =   -1  'True
         TabIndex        =   17
         Top             =   480
         Width           =   2295
      End
      Begin VB.ComboBox cname 
         Height          =   315
         Left            =   2520
         Sorted          =   -1  'True
         TabIndex        =   18
         Top             =   480
         Width           =   4815
      End
      Begin VB.ComboBox comcreditaccountnumber 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   19
         Top             =   1080
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
         Left            =   9360
         TabIndex        =   22
         Top             =   1080
         Width           =   735
      End
      Begin MSMask.MaskEdBox creditamount 
         Height          =   315
         Left            =   7440
         TabIndex        =   21
         Top             =   1080
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
         TabIndex        =   46
         Top             =   3480
         Visible         =   0   'False
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   238
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   1980
         Left            =   120
         TabIndex        =   23
         Top             =   1440
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   3493
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
            Object.Width           =   3351
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
         TabIndex        =   58
         Top             =   240
         Width           =   960
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
         Left            =   8400
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   3440
         Visible         =   0   'False
         Width           =   1650
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
         TabIndex        =   44
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
         TabIndex        =   43
         Top             =   840
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
         TabIndex        =   42
         Top             =   840
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
         TabIndex        =   41
         Top             =   840
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
         TabIndex        =   40
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
         TabIndex        =   39
         Top             =   240
         Width           =   810
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         Caption         =   "ÇÏÎÇá ÈæÇÓØÉ "
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
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   2400
         Width           =   960
      End
      Begin VB.Label lbldebitamount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   7440
         TabIndex        =   30
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         Height          =   195
         Left            =   7440
         TabIndex        =   29
         Top             =   240
         Width           =   540
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enterd By :"
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
         TabIndex        =   28
         Top             =   2400
         Width           =   795
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Verified By :"
         Height          =   195
         Left            =   5760
         TabIndex        =   27
         Top             =   2400
         Width           =   840
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Account Name"
         Height          =   255
         Left            =   2520
         TabIndex        =   26
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         Height          =   195
         Left            =   7440
         TabIndex        =   25
         Top             =   840
         Width           =   540
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Account Name"
         Height          =   255
         Left            =   2520
         TabIndex        =   24
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Credit Number"
         Height          =   195
         Left            =   120
         TabIndex        =   45
         Top             =   840
         Width           =   1005
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10215
      Begin VB.ComboBox comno 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   3480
         TabIndex        =   60
         Top             =   840
         Width           =   3255
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   120
         TabIndex        =   59
         Top             =   840
         Width           =   3255
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "( Tax Credit ) :"
         Height          =   195
         Left            =   6960
         TabIndex        =   57
         Top             =   960
         Width           =   990
      End
      Begin VB.Label lbltaxcredit 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
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
         Left            =   8640
         TabIndex        =   56
         Top             =   960
         Width           =   1350
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sur Tax :"
         Height          =   195
         Left            =   3600
         TabIndex        =   55
         Top             =   1680
         Width           =   645
      End
      Begin VB.Label lblsurtax 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
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
         Left            =   5280
         TabIndex        =   54
         Top             =   1680
         Width           =   1350
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Net Amount :"
         Height          =   195
         Left            =   3600
         TabIndex        =   53
         Top             =   960
         Width           =   930
      End
      Begin VB.Label lblnet 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
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
         Left            =   5280
         TabIndex        =   52
         Top             =   960
         Width           =   1350
      End
      Begin VB.Label lbltransport 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
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
         Left            =   1920
         TabIndex        =   51
         Top             =   1680
         Width           =   1350
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transport :"
         Height          =   195
         Left            =   240
         TabIndex        =   50
         Top             =   1680
         Width           =   765
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice Number :"
         Height          =   195
         Left            =   6960
         TabIndex        =   49
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label lblinvoice 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
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
         Left            =   8640
         TabIndex        =   48
         Top             =   1680
         Width           =   1350
      End
      Begin VB.Label lblname 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lblname"
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
         TabIndex        =   38
         Top             =   600
         Width           =   630
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client Name"
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
         Index           =   0
         Left            =   120
         TabIndex        =   37
         Top             =   600
         Width           =   1035
      End
      Begin VB.Label lblamount 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
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
         Left            =   8640
         TabIndex        =   16
         Top             =   1320
         Width           =   1350
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount :"
         Height          =   195
         Left            =   6960
         TabIndex        =   15
         Top             =   1320
         Width           =   1035
      End
      Begin VB.Label lblvat 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
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
         Left            =   5280
         TabIndex        =   14
         Top             =   1320
         Width           =   1350
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "V. A . T :"
         Height          =   195
         Left            =   3600
         TabIndex        =   13
         Top             =   1320
         Width           =   630
      End
      Begin VB.Label lbltradediscount 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
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
         Left            =   1920
         TabIndex        =   12
         Top             =   1320
         Width           =   1350
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "( Trade Discount ) :"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   1320
         Width           =   1365
      End
      Begin VB.Label lbltotalamount 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
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
         Left            =   1920
         TabIndex        =   10
         Top             =   960
         Width           =   1350
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Allowances :"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   900
      End
      Begin VB.Label lblremarks 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lblremarks"
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
         Left            =   4320
         TabIndex        =   6
         Top             =   600
         Width           =   720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reason  :"
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
         Index           =   1
         Left            =   3240
         TabIndex        =   35
         Top             =   600
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Debit Note"
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
         TabIndex        =   2
         Top             =   240
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date  :"
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
         Left            =   3240
         TabIndex        =   4
         Top             =   240
         Width           =   600
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client Code  :"
         Height          =   195
         Left            =   6840
         TabIndex        =   7
         Top             =   240
         Width           =   945
      End
      Begin VB.Label lbldate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lbldate"
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
         Left            =   4320
         TabIndex        =   5
         Top             =   240
         Width           =   585
      End
      Begin VB.Label lblcode 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblcode"
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
         Height          =   270
         Left            =   8760
         TabIndex        =   8
         Top             =   240
         Width           =   1305
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   6840
         TabIndex        =   61
         Top             =   840
         Width           =   3255
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
Attribute VB_Name = "frmdebitnote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim recemp As New ADODB.Recordset ' this is for the employer
Dim recacc As New ADODB.Recordset  ' this is for the finance master
Dim recinv As New ADODB.Recordset ' this is for the creditmain table
Dim reccust As New ADODB.Recordset
Dim con As New ADODB.Connection
Dim con2 As New ADODB.Connection
Public transferamount As Currency
Dim selectclick As Integer
Public validatedebitamount As Currency
Public totalvalidatedebitamount As Currency
Dim creditamountcount As Currency ' thsi is for count the balance
Dim creditcount As Integer
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
Me.caption = "Debit Note  " & cc
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
If Val(lbldebitamount.caption) < Val(creditamountcount) Then
    MsgBox "Please Check Your Payment Amount", vbInformation, "Invalid Amount"
    creditamountcount = Val(creditamountcount) - Val(creditamount.Text)
    creditamount.SetFocus
    Exit Sub
End If
If Val(lbldebitamount.caption) = Val(creditamountcount) Then
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

Private Sub cmdcancel_Click()
On Error Resume Next
Frame2.Enabled = False
cmdedit.Enabled = True
cmdsave.Enabled = False
cmdprint.Enabled = False
cmdclose.Enabled = True
cmdcancel.Enabled = False
ListView2.ListItems.Clear
For Each Control In Me
    If TypeOf Control Is ComboBox Then
        Control.Text = " "
    End If
Next

lbltradediscount.caption = " "
lblinvoice.caption = " "
lbltransport.caption = "  "
lbldebitamount.caption = " "
creditamount.Text = " "
lblcode.caption = " "
lblamount.caption = " "
lblremarks.caption = " "
lblname.caption = " "
lblvat.caption = " "
lbltotalamount.caption = " "
validatedebitamount = 0
lblunapplied.caption = " "
lblnet.caption = " "
lblsurtax.caption = " "
lbltaxcredit.caption = " "
comno.Enabled = False
Frame2.Enabled = False
Me.caption = "Debit Note"
End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub CMDEDIT_Click()
comno.Enabled = True
fromwho = "a"
ListView2.ListItems.Clear
cmdedit.Enabled = False
cmdadd.Enabled = True
cmdsave.Enabled = True
cmdclose.Enabled = False
Frame1.Enabled = True
comno.Enabled = True
creditcount = 0
creditamount = 0
creditamountcount = 0
cmdcancel.Enabled = True
Frame2.Enabled = True
comno.SetFocus
Call prcclear
End Sub

Private Sub cmdsave_Click()
'this is for payment analysis
    If Trim(comcreditaccountnumber.Text) = "" Or Trim(ccname.Text) = "" Then
        MsgBox "Please check the Credit Account Number", vbInformation, "Empty Data"
        Exit Sub
    End If
    If Trim(combdebitaccnum.Text) = "" Or Trim(cname.Text) = "" Then
        MsgBox "Please Check the Debit Account Number", vbInformation, "Empty Field"
        combdebitaccnum.SetFocus
        Exit Sub
    End If

If ListView2.ListItems.Count > 0 And Trim(comno.Text) <> "" Then
If ListView2.Enabled = True Then
    ListCount = ListView2.ListItems.Count
    i = 1
    For i = 1 To ListCount
        checkallamount = checkallamount + Val(Trim(ListView2.ListItems(i).ListSubItems(2).Text))
    Next
    
    If Val(checkallamount) <> Val(Trim(lbldebitamount.caption)) Then
        MsgBox "Please Check Credit Amount it is not Equal to Debit Amount", vbInformation, "Amount Overflow"
        checkallamount = 0
        Exit Sub
    End If
    'this is for clear the debitnote table and cashjournal table *****************
    Dim deletedata As New ADODB.Recordset
    deletedata.Open "delete from debitnote where status = 'UnPosted' and creditnoteno = '" & Trim(comno.Text) & "'", con, adOpenKeyset, adLockOptimistic
    '***************************************************************************

Dim recsales As New ADODB.Recordset ' this is for debitnote
recsales.Open "debitnote", con, adOpenKeyset, adLockOptimistic
    'this is for take the journal number
    Dim recset As New ADODB.Recordset
    recset.Open "Select * from setup", con, adOpenKeyset, adLockOptimistic
    If Trim(recset!CurrentMoYr) <> Format(Date, "mmyy") Then
        recset!CurrentMoYr = Format(Date, "mmyy")
        recset!nextjn = "00001"
        recset.Update
    End If
    recset.Requery
    totaljournal = "SPA-" & recset!CurrentMoYr & "-" & Right(Trim(recset!nextjn), 5)
    totaljournal = "0000" & (Val(Trim(recset!nextjn)) + 1)
    recset!nextjn = Right(totaljournal, 5)
    recset.Update
    
' this is for debit side
        recsales.AddNew
        recsales!SerialNo = totaljournal
        recsales!creditnoteno = Trim(comno.Text)
        recsales!ticket = "1"
        recsales!deletemark = "0"
        recsales!accountnumber = Trim(combdebitaccnum.Text)
            namenamenum = InStr(1, Trim(cname.Text), "\", vbTextCompare)
            If namenamenum > 0 Then
            namename = Mid(Trim(cname.Text), 1, (namenamenum - 1))
            Else
            namename = Trim(cname.Text)
            End If
        recsales!accountname = namename
        recsales!accountnamearab = Mid(Trim(cname.Text), (namenamenum + 1), Val(Len(Trim(cname.Text))) - Val(namenamenum))
            'get the mohter name
            getmothername Trim(combdebitaccnum.Text), noname, cc
        recsales!mothername = cc
        recsales!Description = Trim(ListView2.ListItems(1).ListSubItems(1).Text)
        recsales!TRansDate = Format(Date, "mm/dd/yyyy")
        recsales!creditamount = FormatNumber(lbldebitamount.caption, 4)
        recsales!DebitAmount = "0"
        recsales!reasons = Trim(lblremarks.caption)
        recsales!Status = "UnPosted"
        recsales!Trantype = "SPA"
        recsales.Update
' this is for listview

For xxx = 1 To ListView2.ListItems.Count
        recsales.AddNew
        recsales!SerialNo = totaljournal
        recsales!creditnoteno = Trim(comno.Text)
        recsales!ticket = xxx + 1
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
        recsales!mothername = cc
        recsales!Description = Trim(cname.Text)  'againts name
        recsales!TRansDate = Format(Date, "mm/dd/yyyy")
        recsales!DebitAmount = FormatNumber(Trim(ListView2.ListItems(xxx).ListSubItems(2).Text), 4)
        recsales!creditamount = "0"
        recsales!reasons = Trim(lblremarks.caption)
        recsales!Status = "UnPosted"
        recsales!Trantype = "SPA"
        recsales.Update
Next
End If
On Error Resume Next
dataanu.rscom_debitnote_byJournal_Grouping.Close
On Error GoTo 0
dataanu.com_debitnote_byJournal_Grouping totaljournal
re_debitnote_byjournal.PrintReport False, rptRangeAllPages

End If
'this is the common things
MsgBox "Your Data Saved Successfully", vbInformation, "Save Conformation"
cmdsave.Enabled = False
cmdcancel.Enabled = False
cmdclose.Enabled = True
cmdedit.Enabled = True
comno.Enabled = False
Me.caption = "Debit Note"
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
Me.caption = "Debit Note  " & cc
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
recfindacc.Open "Select * from financemaster where accountnameeng = " & "'" & namename & "'", con, adOpenKeyset, adLockOptimistic
If recfindacc.BOF = True Then
    MsgBox "Please choose the Currect Account Number", vbInformation, "Invalid accountNumber"
    cname.SetFocus
    Exit Sub
Else
    combdebitaccnum.Text = recfindacc!AccountCode
End If
recfindacc.Close
End Sub

Private Sub combdebitaccnum_Click()
Dim recfindaccount As New ADODB.Recordset
recfindaccount.Open "select * from financemaster where accountcode = " & "'" & Trim(combdebitaccnum.Text) & "'", con, adOpenKeyset, adLockOptimistic
If recfindaccount.BOF = False Then
    cname.Text = Trim(recfindaccount!accountnameeng) & "\" & Trim(recfindaccount!accountnamearab)
End If
recfindaccount.Close
 Dim cc As String
    Dim ccc As String
    getmothername Trim(combdebitaccnum.Text), ccc, cc
Me.caption = "Debit Note  " & cc
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
    cname.Text = recfindaccount!accountnameeng & "\" & recfindaccount!accountnamearab
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
Me.caption = "Debit Note  " & cc
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
recfindaccount.Open "select * from financemaster where accountcode = " & "'" & Trim(comcreditaccountnumber.Text) & "'", con, adOpenKeyset, adLockOptimistic
If recfindaccount.BOF = True Then
    MsgBox "Please Choose the Sub Account", vbInformation, "Invalid AccountNumber"
    comcreditaccountnumber.SetFocus
    Exit Sub
Else
    ccname.Text = recfindaccount!accountnameeng & "\" & recfindaccount!accountnamearab
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
comenterd.SetFocus
End If

End Sub

Private Sub deletefull_Click()
creditamountcount = creditamountcount - Val(ListView2.SelectedItem.SubItems(2))
deleteindex = ListView2.SelectedItem.Index
ListView2.ListItems.Remove (deleteindex)
cmdadd.Enabled = True
creditcount = creditcount - 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
con.Close
con2.Close
End Sub
Private Sub comno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Trim(comno.Text) <> "" Then
    Dim reccheckinjournal As New ADODB.Recordset
    reccheckinjournal.Open "select * from debitnote where CreditNoteno = '" & Trim(comno.Text) & "'", con, adOpenKeyset, adLockOptimistic
    
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
            'this is mean he like to analyse again
         Else
            cmdcancel_Click
            Frame2.Enabled = False
            comno.Enabled = False
            Exit Sub ' stop here
         End If
    End If
    reccheckinjournal.Close
    On Error Resume Next
    recinv.Close
    On Error GoTo 0
    
    recinv.Open "Select * from credmain where invc_no = " & "'" & Trim(comno.Text) & "'" & " order by invc_date", con2, adOpenKeyset, adLockOptimistic
        
        If recinv.BOF = True Then  'if it is empty
            MsgBox "No Any Credit Notes to Pay", vbInformation, "Empty"
            Exit Sub
        End If
    
    On Error Resume Next
    lblcode.caption = recinv!cust_code
    lblamount.caption = FormatNumber(recinv!tot_amt, 2)
    lbltransport.caption = Format(recinv!transchg, "############0.##0")
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
    
    reccust.Open "select * from marcusfl where cust_code = " & "'" & anucustcode & "'", con2, adOpenKeyset, adLockOptimistic
    lblname.caption = reccust!first_name
    anumaincode = Trim(reccust!mcustcode)
    reccust.Close
    
    cmdsave.Enabled = True
    cmdcancel.Enabled = True
    If fromwho = "a" Then
        lbldebitamount.caption = recinv!tot_amt
    End If
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
con.Open "Dsn=Finance;Uid=Sa;Pwd=;"
con2.Mode = adModeShareDenyNone
con2.Open "Dsn=anufoxpro;uid=sa;pwd=;"
recinv.Open "Select * from credmain where trim(left(invc_no,3)) = 'ODN' order by invc_date", con2, adOpenKeyset, adLockOptimistic
    While recinv.EOF = False
        comno.AddItem Trim(recinv!invc_no)
        recinv.MoveNext
    Wend
recinv.Close
lbldate.caption = Format(Date, "dd/mm/yyyy")
Call prcclear
recemp.Open "Select * from newlog order by userid", con, adOpenKeyset, adLockOptimistic

Dim constring As String
Dim myclass As New HabitatClass
Dim sqltable As Boolean
Dim xtable As String
Dim conrecacc As New ADODB.Connection

constring = "Dsn=Finance;Uid=Sa;Pwd=;"
xtable = "select * from financemaster where active <> '0'"
sqltable = True
myclass.GetTables recacc, conrecacc, xtable, constring, sqltable

'recacc.Open "select * from financemaster where active <> '0'", con, adOpenKeyset, adLockOptimistic
recemp.MoveFirst
While recemp.EOF = False
On Error Resume Next
    comverified.AddItem recemp!Userid
    comenterd.AddItem recemp!Userid
    recemp.MoveNext
Wend
On Error GoTo 0
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
recemp.Close
recacc.Close
conrecacc.Close
End Sub

Private Sub ListView2_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
If ListView2.ListItems.Count > 0 And Button = vbRightButton Then
    PopupMenu full
End If
End Sub

Private Sub prcclear()
lblinvoice.caption = " "
lbltransport.caption = " "
lblamount.caption = " "
lblcode.caption = " "
lblname.caption = " "
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

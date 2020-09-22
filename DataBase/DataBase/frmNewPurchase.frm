VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmNewPurchse 
   Caption         =   "Purchase Setup"
   ClientHeight    =   8310
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11790
   Icon            =   "frmNewPurchase.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   11790
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   8295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11775
      Begin VB.TextBox txtPriorPayments 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   78
         Top             =   6720
         Width           =   1335
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   325
         Left            =   9360
         TabIndex        =   75
         Top             =   7800
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtLastAmt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4920
         TabIndex        =   54
         Top             =   7440
         Width           =   1335
      End
      Begin VB.TextBox txtTotalInvoice 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4920
         TabIndex        =   51
         Top             =   7080
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         Caption         =   "     With Rate"
         Height          =   375
         Left            =   240
         TabIndex        =   50
         Top             =   7560
         Width           =   2415
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Culculate Çáå ÍÇÓÈÉ "
         Height          =   325
         Left            =   8160
         TabIndex        =   46
         Top             =   7440
         Width           =   3495
      End
      Begin VB.CommandButton Command2 
         Caption         =   "E&xit"
         Height          =   325
         Left            =   10560
         TabIndex        =   45
         Top             =   7800
         Width           =   1095
      End
      Begin VB.CommandButton cmdGoPay 
         Caption         =   "Go Back ÇáÑÌæÚ ááÎáÝ "
         Height          =   325
         Left            =   8160
         TabIndex        =   44
         Top             =   7080
         Width           =   3495
      End
      Begin VB.Frame Frame3 
         Height          =   2655
         Left            =   120
         TabIndex        =   3
         Top             =   1560
         Width           =   11535
         Begin VB.TextBox txtCuDut 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   8880
            TabIndex        =   33
            Top             =   2160
            Width           =   1335
         End
         Begin VB.TextBox txtSuT 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   8880
            TabIndex        =   32
            Top             =   1800
            Width           =   1335
         End
         Begin VB.TextBox txtCPro 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   8880
            TabIndex        =   31
            Top             =   1440
            Width           =   1335
         End
         Begin VB.TextBox txtSrT 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00;(#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3073
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   8880
            TabIndex        =   30
            Top             =   1080
            Width           =   1335
         End
         Begin VB.TextBox txtST 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   8880
            TabIndex        =   29
            Top             =   720
            Width           =   1335
         End
         Begin VB.TextBox txtRCusD 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   8040
            TabIndex        =   28
            Top             =   2160
            Width           =   615
         End
         Begin VB.TextBox txtRSurT 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   8040
            TabIndex        =   27
            Top             =   1800
            Width           =   615
         End
         Begin VB.TextBox txtRComPro 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   8040
            TabIndex        =   26
            Top             =   1440
            Width           =   615
         End
         Begin VB.TextBox txtRSrT 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   8040
            TabIndex        =   25
            Top             =   1080
            Width           =   615
         End
         Begin VB.TextBox txtRST 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   8040
            TabIndex        =   24
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox txtRTD 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1680
            TabIndex        =   9
            Top             =   1320
            Width           =   615
         End
         Begin VB.TextBox txtRCD 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1680
            TabIndex        =   8
            Top             =   1680
            Width           =   615
         End
         Begin VB.TextBox txttrdis 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00;(#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3073
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   2520
            TabIndex        =   7
            Top             =   1320
            Width           =   1335
         End
         Begin VB.TextBox txtCashDis 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2520
            TabIndex        =   6
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox txtNet 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00;(#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3073
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   2520
            TabIndex        =   5
            Top             =   2280
            Width           =   1335
         End
         Begin VB.TextBox txtGross 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00;(#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3073
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   2520
            TabIndex        =   4
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label27 
            Caption         =   "Rate %"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7920
            TabIndex        =   49
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label24 
            Caption         =   "Amount"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8880
            TabIndex        =   48
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label21 
            Caption         =   "Description"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6480
            TabIndex        =   47
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label14 
            Caption         =   "Sales Tax"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6480
            TabIndex        =   43
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label13 
            Caption         =   "Tax Credit"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6480
            TabIndex        =   42
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Label Label12 
            Caption         =   "Commercial Profit Tax"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6480
            TabIndex        =   41
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "Service Tax"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6480
            TabIndex        =   40
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label Label10 
            Caption         =   "Custom Duties"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6480
            TabIndex        =   39
            Top             =   2160
            Width           =   1455
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            Caption         =   "ÖÑÇÆÈ ÇáãÈíÚÇÊ "
            Height          =   255
            Left            =   9960
            TabIndex        =   38
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            Caption         =   "ÖÑÇÆÈ ÇáÎÏãÉ "
            Height          =   255
            Left            =   9960
            TabIndex        =   37
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            Caption         =   "ÇáÑÈÍ ÇáÊÌÇÑí"
            Height          =   255
            Left            =   9960
            TabIndex        =   36
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            Caption         =   "ÖÑÇÆÈ ÇÖÇÝíÉ "
            Height          =   255
            Left            =   9960
            TabIndex        =   35
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            Caption         =   "ÚãáÇÁ ãÓÊÍÞÉ "
            Height          =   255
            Left            =   9960
            TabIndex        =   34
            Top             =   2160
            Width           =   1335
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            Caption         =   "ÕÇÝí ÊßáÝÉ ÇáÈÖÇÚÉ "
            Height          =   255
            Left            =   3840
            TabIndex        =   21
            Top             =   2280
            Width           =   1575
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            Caption         =   "ÇáÎÕã ÇáäÞÏí"
            Height          =   255
            Left            =   4200
            TabIndex        =   20
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            Caption         =   "ÇáÎÕã ÇáÊÌÇÑí "
            Height          =   255
            Left            =   4080
            TabIndex        =   19
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "ãÌãæÚ  ÊßáíÝ ÇáÈÖÇÚÉ "
            Height          =   375
            Left            =   3960
            TabIndex        =   18
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Description"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "Amount"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2640
            TabIndex        =   15
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "Rate %"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1560
            TabIndex        =   14
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label6 
            Caption         =   "Trade Discount"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Label Label9 
            Caption         =   "Net Cost of Goods"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   2280
            Width           =   1935
         End
         Begin VB.Label Label15 
            Caption         =   "Cash Discount"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   1680
            Width           =   1455
         End
         Begin VB.Label Label5 
            Caption         =   "Gross Cost of Goods"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   720
            Width           =   1935
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1335
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   11535
         Begin VB.TextBox txtSerialNo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1440
            TabIndex        =   76
            Top             =   120
            Width           =   1095
         End
         Begin VB.TextBox txtSENo 
            Height          =   285
            Left            =   5040
            TabIndex        =   69
            Top             =   120
            Width           =   1095
         End
         Begin VB.TextBox txtPOno 
            Height          =   285
            Left            =   8880
            TabIndex        =   68
            Top             =   240
            Width           =   1095
         End
         Begin MSMask.MaskEdBox mskPOdate 
            Height          =   315
            Left            =   8880
            TabIndex        =   67
            Top             =   600
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtinvNo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1440
            TabIndex        =   17
            Top             =   480
            Width           =   1095
         End
         Begin MSMask.MaskEdBox txtINVdate 
            Height          =   315
            Left            =   1440
            TabIndex        =   57
            Top             =   840
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MSKdueDate 
            Height          =   315
            Left            =   5040
            TabIndex        =   58
            Top             =   840
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskSEDate 
            Height          =   315
            Left            =   5040
            TabIndex        =   70
            Top             =   480
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label42 
            Caption         =   "Serial Number"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   77
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label Label41 
            Caption         =   "Store Date"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4200
            TabIndex        =   74
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label40 
            Alignment       =   1  'Right Justify
            Caption         =   "ÊÇÑíÎ "
            Height          =   255
            Left            =   6480
            TabIndex        =   73
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label39 
            Caption         =   "Store No"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4200
            TabIndex        =   72
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label Label38 
            Alignment       =   1  'Right Justify
            Caption         =   " ÑÞã ÇáÅÖÇÝÉ"
            Height          =   255
            Left            =   6360
            TabIndex        =   71
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label37 
            Alignment       =   1  'Right Justify
            Caption         =   " ÑÞã Ã. ÇáÊæÑíÏ"
            Height          =   255
            Left            =   10440
            TabIndex        =   66
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label36 
            Caption         =   "P.O  Number"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7680
            TabIndex        =   65
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label35 
            Alignment       =   1  'Right Justify
            Caption         =   " ÊÇÑíÎ"
            Height          =   375
            Left            =   10320
            TabIndex        =   64
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label34 
            Caption         =   "P.O Date"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7680
            TabIndex        =   63
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            Caption         =   "ÊÇÑíÎ ÇáÝÇÊæÑÉ "
            Height          =   255
            Left            =   2640
            TabIndex        =   62
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label29 
            Caption         =   "Invoice Date"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   61
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            Caption         =   "ÊÇÑíÎ ÇáÇÓÊÍÞÇÞ"
            Height          =   255
            Left            =   6120
            TabIndex        =   60
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label31 
            Caption         =   "Due Date"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4200
            TabIndex        =   59
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            Caption         =   "ÝÇÊæÑÉ ÑÞã "
            Height          =   255
            Left            =   2760
            TabIndex        =   22
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   "Invoice Number"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   2
            Top             =   480
            Width           =   1455
         End
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2175
         Left            =   120
         TabIndex        =   23
         Top             =   4320
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   3836
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label43 
         Caption         =   "All Prior Payments"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   79
         Top             =   6720
         Width           =   2055
      End
      Begin VB.Label Label33 
         Caption         =   "Total Invoice"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   56
         Top             =   7440
         Width           =   2055
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         Caption         =   "ÇÌãÇáí ÇáÝÇÊæÑÉ "
         Height          =   255
         Left            =   6360
         TabIndex        =   55
         Top             =   7440
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "Net Invoice Amout"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   53
         Top             =   7080
         Width           =   2055
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         Caption         =   "ÕÇÝí ÇÌãÇáí ÇáÝÇÊæÑÉ"
         Height          =   255
         Left            =   6480
         TabIndex        =   52
         Top             =   7080
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmNewPurchse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TradeDis
Dim CashDis
Dim netCostofGood
Dim SalesTax
Dim saviceTax
Dim SurTax
Dim ComProTax
Dim ServTax
Dim TaxCre
Dim CustDue
Dim TotInvAmt
Dim GrossCost
Dim Pilappalam As Currency


Private Sub Command1_Click()

End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
Me.txtRTD.Visible = False
Me.txtRCD.Visible = False
Me.txtRST.Visible = False
Me.txtRSrT.Visible = False
Me.txtRComPro.Visible = False
Me.txtRSurT.Visible = False
Me.txtRCusD.Visible = False
Me.Label3.Visible = False
Me.Label27.Visible = False
ElseIf Check1.Value = 0 Then
Me.txtRTD.Visible = True
Me.txtRCD.Visible = True
Me.txtRST.Visible = True
Me.txtRSrT.Visible = True
Me.txtRComPro.Visible = True
Me.txtRSurT.Visible = True
Me.txtRCusD.Visible = True
Me.Label3.Visible = True
Me.Label27.Visible = True

End If
End Sub

Private Sub cmdGoPay_Click()
On Error Resume Next
FrmPayableSetup.Show 1
Unload Me
End Sub

Private Sub cmdsave_Click()

If txtGross.Text = "" Then
MsgBox "Please Select the Amount", vbInformation, "Amount Can not be Empty"
Exit Sub
End If



Pilappalam = 0
Dim CON1 As ADODB.Connection
Set CON1 = New ADODB.Connection
     Dim Lks As Currency

conStr = "Provider=MSDASQL;DSN=Finance;UID=sa; PWD=;"

CON1.Open conStr

Dim xPayInvoiceDetails As New ADODB.Recordset

xPayInvoiceDetails.Open "select * from PayTempInvoiceDetails", CON1, adOpenDynamic, adLockOptimistic

If cmdsave.caption = "&Culculate Çáå ÍÇÓÈÉ " Then
  If Check1.Value = 0 Then
'cmdSave.Caption = "&Save to Payable Setup ÍÝÙ ÊÍãíá ÇáãÏÝæÚÇÊ "

    GrossCost = Val(txtGross.Text)
    txtGross.Text = Format(GrossCost, "###,###,###.#0")
    
    TradeDis = Val(GrossCost) * txtRTD / 100
    txttrdis.Text = Format(TradeDis, "###,###,###.#0")
    
    
    CashDis = Val(GrossCost) * txtRCD / 100
    txtCashDis.Text = Format(CashDis, "###,###,###.#0")
    
    netCostofGood = Val(GrossCost) - Val(CashDis) - Val(TradeDis)
    
    txtNet.Text = Format(netCostofGood, "###,###,###.#0") 'Net cost of goods
    
    SalesTax = Val(netCostofGood) * txtRST / 100
    txtST.Text = Format(SalesTax, "###,###,###.#0")
    
    saviceTax = Val(netCostofGood) * txtRSrT / 100
    txtSrT.Text = Format(saviceTax, "###,###,###.#0")
    
    ComProTax = Val(netCostofGood) * txtRComPro / 100
    txtCPro.Text = Format(ComProTax, "###,###,###.#0")
    
    SurTax = Val(netCostofGood) * txtRSurT / 100
    txtSuT.Text = Format(SurTax, "###,###,###.#0")
    
'    TaxCre = Val(netCostofGood) * txtRTc / 100
'    txtTC.Text = Format(TaxCre, "###,###,###.#0")
    
    CustDue = Val(netCostofGood) * txtRCusD / 100
    txtCuDut.Text = Format(CustDue, "###,###,###.#0")
    
  '  TotInvAmt = netCostofGood + Val(SalesTax) + Val(saviceTax) + Val(ComProTax) + Val(CustDue) '+ Val(SurTax)
      TotInvAmt = netCostofGood + Val(SalesTax) + Val(saviceTax) + Val(ComProTax) + Val(CustDue) '+ Val(SurTax)
  
    txtTotalInvoice = Val(txtTotalInvoice) + Val(Format(TotInvAmt, "###,###,###.#0"))
    
    Me.Command2.caption = "Cancel"

ElseIf Check1.Value = 1 Then
    
    TotInvAmt = Val(txtGross.Text) - Val(txttrdis.Text) - Val(txtCashDis.Text) + Val(txtST.Text) + Val(txtSrT.Text) + Val(txtCPro.Text) + Val(txtCuDut.Text) '- Val(txtSuT.Text)
     'TotInvAmt = Val(txtGross.Text) - Val(txttrdis.Text) - Val(txtCashDis.Text) + Val(txtST.Text) + Val(txtSrT.Text) + Val(txtCPro.Text) + Val(txtCuDut.Text)

    txtTotalInvoice = Val(txtTotalInvoice) + Val(Format(TotInvAmt, "###,###,###.#0"))
    
    Me.Command2.caption = "Cancel"

End If





rosa = MsgBox("Do you want to Add the Multi Invoice", vbYesNo + vbInformation, "Confirmation")
If rosa = vbYes Then
     Set MItem = Me.ListView1.ListItems.Add(, , Trim(txtInvNo.Text))
     MItem.SubItems(1) = Trim(txtGross.Text)
     MItem.SubItems(2) = Trim(txtINVdate.Text)
     MItem.SubItems(3) = Trim(mskDueDate.Text)
     MItem.SubItems(4) = Trim(txttrdis.Text)
     MItem.SubItems(5) = Trim(txtCashDis.Text)
     MItem.SubItems(6) = Trim(txtST.Text)
     MItem.SubItems(7) = Trim(txtSrT.Text)
     MItem.SubItems(8) = Trim(txtCPro.Text)
     MItem.SubItems(9) = Trim(txtSuT.Text) 'Tax Credit
     MItem.SubItems(10) = Trim(txtCuDut.Text)
     MItem.SubItems(11) = Trim(txtTotalInvoice.Text)
     MItem.SubItems(12) = Trim(txtPoNo.Text)
     MItem.SubItems(13) = Trim(mskPOdate.Text)
     MItem.SubItems(14) = Trim(txtSENo.Text)
     MItem.SubItems(15) = Trim(MskSEDate.Text)

     Pilappalam = Val(Pilappalam) + txtTotalInvoice.Text
     Lks = txtTotalInvoice.Text


     txtLastAmt.Text = Val(Pilappalam)
     txtINVdate.Text = "__/__/____"
     mskPOdate.Text = "__/__/____"
     txtInvNo.Text = ""
     mskDueDate.Text = "__/__/____"
     txtGross.Text = ""
     txttrdis.Text = ""
     txtCashDis.Text = ""
     txtST.Text = ""
     txtNet.Text = ""
     txtSrT.Text = ""
     txtCPro.Text = ""
     txtSuT.Text = ""
     txtCuDut.Text = ""
    ' txtTotalInvoice.Text = ""
     txtPoNo.Text = ""
     txtPoNo.Text = ""
txtGross.SetFocus
Else  'Msgbox NO
klop = MsgBox("Do you want to Save the Details", vbYesNo, "Save")
If klop = vbNo Then
MsgBox "Nothing is saved you Lost all the details you Entered", vbInformation
Exit Sub
ElseIf klop = vbYes Then

     Set MItem = Me.ListView1.ListItems.Add(, , Trim(txtInvNo.Text))
     MItem.SubItems(1) = Trim(txtGross.Text)
     MItem.SubItems(2) = Trim(txtINVdate.Text)
     MItem.SubItems(3) = Trim(mskDueDate.Text)
     MItem.SubItems(4) = Trim(txttrdis.Text)
     MItem.SubItems(5) = Trim(txtCashDis.Text)
     MItem.SubItems(6) = Trim(txtST.Text)
     MItem.SubItems(7) = Trim(txtSrT.Text)
     MItem.SubItems(8) = Trim(txtCPro.Text)
     MItem.SubItems(9) = Trim(txtSuT.Text) 'tax Cre
     MItem.SubItems(10) = Trim(txtCuDut.Text)
     MItem.SubItems(11) = Trim(txtTotalInvoice.Text)
     MItem.SubItems(12) = Trim(txtPoNo.Text)
     MItem.SubItems(13) = Trim(mskPOdate.Text)
     MItem.SubItems(14) = Trim(txtSENo.Text)
     MItem.SubItems(15) = Trim(MskSEDate.Text)

     Lks = txtTotalInvoice.Text
     Pilappalam = Val(Pilappalam) + Val(Lks)

     

     txtINVdate.Text = "__/__/____"
     txtInvNo.Text = ""
     mskDueDate.Text = "__/__/____"
     mskPOdate.Text = "__/__/____"
     txtGross.Text = ""
     txttrdis.Text = ""
     txtCashDis.Text = ""
     txtST.Text = ""
     txtSrT.Text = ""
     txtCPro.Text = ""
     txtSuT.Text = ""
     txtCuDut.Text = ""
'     txtTotalInvoice.Text = ""
     txtNet.Text = ""
     txtPoNo.Text = ""
 txtSENo.Text = ""
 Dim TotTaxCred
 
        
              n = 0
              For n = 1 To Me.ListView1.ListItems.Count
                  LINVNo = Me.ListView1.ListItems.Item(n)
                  Linvamt = Me.ListView1.ListItems.Item(n).SubItems(1)
                  LinvDate = Me.ListView1.ListItems.Item(n).SubItems(2)
                  LDueDate = Me.ListView1.ListItems.Item(n).SubItems(3)
                  LtradDis = Me.ListView1.ListItems.Item(n).SubItems(4)
                  LCAshDis = Me.ListView1.ListItems.Item(n).SubItems(5)
                  LSalsTax = Me.ListView1.ListItems.Item(n).SubItems(6)
                  LServTax = Me.ListView1.ListItems.Item(n).SubItems(7)
                  ComPro = Me.ListView1.ListItems.Item(n).SubItems(8)
                  Addded = Me.ListView1.ListItems.Item(n).SubItems(9)
                  CustDut = Me.ListView1.ListItems.Item(n).SubItems(10)
                  TotInv = Me.ListView1.ListItems.Item(n).SubItems(11)
                  LPONo = Me.ListView1.ListItems.Item(n).SubItems(12)
                  LpoDate = Me.ListView1.ListItems.Item(n).SubItems(13)
                  LSENo = Me.ListView1.ListItems.Item(n).SubItems(14)
                  LSEDate = Me.ListView1.ListItems.Item(n).SubItems(15)
                     
              TotTaxCred = Val(TotTaxCred) + Val(Addded)
                   On Error GoTo 0
                     With xPayInvoiceDetails
                           .AddNew
                            !SerialNo = FrmPayableSetup.txtSerialNo.Text
                            !InvNo = LINVNo
                            !InvDate = LinvDate
                            !duedate = LDueDate
                            !invAmt = Linvamt
                            !TradeDis = LtradDis
                            !CashDis = LCAshDis
                            !SalesDis = LSalsTax
                            !ServDis = LServTax
                            !ComProf = ComPro
                            !AddDedTax1 = Addded
                            !CustDut = CustDut
                            !TotInv = TotInv
                            !PoNumber = LPONo
                            !PODate = LpoDate
                            !SENumber = LSENo
                            !SEDate = LSEDate
                       'On Error Resume Next
                           .Update
                      End With
                      
                      
               Next
               
''This is to find the All Prior Payments for the Same Invoice and put it
'Dim Prir
'Prir = Trim(txtInvNo)
'Dim PriorInv As New ADODB.Recordset
'PriorInv.Open "select sum(InvAmt) as InvAmt from  PayInvoiceDetails  where invno = " & "'" & Prir & "'" & "", conString, adOpenDynamic, adLockOptimistic
''Mambalam = "SELECT CUSTNO, SUM(debitamount) AS cashamount, SUM(checkreceipt) AS checkamount, SUM(debitamount) + SUM(checkreceipt) AS totalamount From vouchers WHERE (deleted = '0') AND (receiptDate <= " & "'" & Format(mfrom, "mm/dd/yyyy") & "'" & ") AND (CUSTNO = " & "'" & inv & "'" & ") GROUP BY CUSTNO"
'
'txtPriorPayments.Text = IIf(IsNull(PriorInv!InvAmt), 0, (PriorInv!InvAmt))
'
              
              
              
      
                    'To find the Total Invoice Amt
                    Dim CounttoTotList As Long
                    Dim FindTota, VArSum
                    CounttoTotList = Me.ListView1.ListItems.Count
                    
                    For i = 1 To CounttoTotList
                    FindTota = Me.ListView1.ListItems(i).SubItems(1)
                    VArSum = Val(VArSum) + FindTota
                    
                    Next
        
              
              
               
               
FrmPayableSetup.txtInvAmt.Text = Val(Pilappalam) '- Val(txtPriorPayments.Text)
'FrmPayableSetup.txtInvAmt.Text = Val(VArSum)
'FrmPayableSetup.txtInvAmt.Text = Format(TotInvAmt, "###,###,###.#0")
FrmPayableSetup.txtTaxCredit.Text = Val(TotTaxCred)
             MsgBox "records Has Been Added Successfully ", vbInformation, "SAVE"
Unload Me
End If
End If
Else


'FrmPayableSetup.Show


FrmPayableSetup.txtInvNo = Me.txtInvNo
FrmPayableSetup.txtInvAmt = Format(TotInvAmt, "###,###,###.#0")
FrmPayableSetup.Command3.Enabled = True
FrmPayableSetup.txtInvAmt.Enabled = True
FrmPayableSetup.txtInvNo.Enabled = True
'FrmPayableSetup.mskInvDate.SetFocus
Unload Me
End If

End Sub

Private Sub cmdUpdate_Click()
'--------------CAlculate -------------------------------------
  If Check1.Value = 0 Then
'cmdSave.Caption = "&Save to Payable Setup ÍÝÙ ÊÍãíá ÇáãÏÝæÚÇÊ "

    GrossCost = Val(txtGross.Text)
    txtGross.Text = Format(GrossCost, "###,###,###.#0")
    
    TradeDis = Val(GrossCost) * txtRTD / 100
    txttrdis.Text = Format(TradeDis, "###,###,###.#0")
    
    
    CashDis = Val(GrossCost) * txtRCD / 100
    txtCashDis.Text = Format(CashDis, "###,###,###.#0")
    
    netCostofGood = Val(GrossCost) - Val(CashDis) - Val(TradeDis)
    
    txtNet.Text = Format(netCostofGood, "###,###,###.#0") 'Net cost of goods
    
    SalesTax = Val(netCostofGood) * txtRST / 100
    txtST.Text = Format(SalesTax, "###,###,###.#0")
    
    saviceTax = Val(netCostofGood) * txtRSrT / 100
    txtSrT.Text = Format(saviceTax, "###,###,###.#0")
    
    ComProTax = Val(netCostofGood) * txtRComPro / 100
    txtCPro.Text = Format(ComProTax, "###,###,###.#0")
    
  SurTax = Val(netCostofGood) * txtRSurT / 100
  txtSuT.Text = Format(SurTax, "###,###,###.#0")

CustDue = Val(netCostofGood) * txtRCusD / 100
txtCuDut.Text = Format(CustDue, "###,###,###.#0")

TotInvAmt = netCostofGood + Val(SalesTax) + Val(saviceTax) + Val(ComProTax) + Val(CustDue) '+ Val(SurTax)
txtTotalInvoice = Format(TotInvAmt, "###,###,###.#0")

Me.Command2.caption = "Cancel"
ElseIf Check1.Value = 1 Then
TotInvAmt = Val(txtGross.Text) - Val(txttrdis.Text) - Val(txtCashDis.Text) + Val(txtST.Text) + Val(txtSrT.Text) + Val(txtCPro.Text) + Val(txtCuDut.Text) '- Val(txtSuT.Text)
txtTotalInvoice = Format(TotInvAmt, "###,###,###.#0")
Me.Command2.caption = "Cancel"
End If

'----------------------------------

  
'THIS IS TO DELETE THE RECORD IN THE LISTVIEW
 Dim DeleteItemX
 DeleteItemX = Me.ListView1.SelectedItem.Index
 
 Me.ListView1.ListItems.Remove DeleteItemX
'--------------------------------------------
     
     
     Set MItem = Me.ListView1.ListItems.Add(, , Trim(txtInvNo.Text))
     MItem.SubItems(1) = Trim(txtGross.Text)
     MItem.SubItems(2) = Trim(txtINVdate.Text)
     MItem.SubItems(3) = Trim(mskDueDate.Text)
     MItem.SubItems(4) = Trim(txttrdis.Text)
     MItem.SubItems(5) = Trim(txtCashDis.Text)
     MItem.SubItems(6) = Trim(txtST.Text)
     MItem.SubItems(7) = Trim(txtSrT.Text)
     MItem.SubItems(8) = Trim(txtCPro.Text)
     MItem.SubItems(9) = Trim(txtSuT.Text)
     MItem.SubItems(10) = Trim(txtCuDut.Text)
     MItem.SubItems(11) = Trim(txtTotalInvoice.Text)
     MItem.SubItems(12) = Trim(txtPoNo.Text)
     MItem.SubItems(13) = Trim(mskPOdate.Text)
     MItem.SubItems(14) = Trim(txtSENo.Text)
     MItem.SubItems(15) = Trim(MskSEDate.Text)

     
     
     
     
     
     txtINVdate.Text = "__/__/____"
     txtInvNo.Text = ""
     mskDueDate.Text = "__/__/____"
     mskPOdate.Text = "__/__/____"
     txtGross.Text = ""
     txttrdis.Text = ""
     txtCashDis.Text = ""
     txtST.Text = ""
     txtSrT.Text = ""
    txtCPro.Text = ""
     txtSuT.Text = ""
     txtCuDut.Text = ""
     txtTotalInvoice.Text = ""
      txtNet.Text = ""
      txtPoNo.Text = ""
      txtSENo.Text = ""
      MskSEDate.Text = "__/__/____"


'THIS IS TO DELETE ALL THE RELATED RECORDS IN THE TABLE

   Dim xItemFOrDel
  xItemFOrDel = Me.txtSerialNo.Text
      
       ' FrmPayableSetup.ListView1.ListItems.Remove xItemFOrDel
Dim rstDelPayInvoiceDetails As New ADODB.Recordset
rstDelPayInvoiceDetails.Open "delete from PayInvoiceDetails where serialno = " & "'" & xItemFOrDel & "'" & "", constring, adOpenDynamic, adLockOptimistic

'rstDelPayInvoiceDetails.Close


'--------------------------- Second ----------------------------

rosa = MsgBox("Do you want to Add the Multi Invoice", vbYesNo + vbInformation, "Confirmation")
If rosa = vbYes Then
'txtGross.SetFocus
   

'Call InvUpdateAdditional

Me.cmdsave.Visible = True
Me.cmdUpdate.caption = False
Exit Sub


Else  'Msgbox NO
klop = MsgBox("Do you want to Save the Details", vbYesNo, "Save")
    If klop = vbNo Then
    MsgBox "Nothing is saved you Lost all the details you Entered", vbInformation
    Exit Sub
    ElseIf klop = vbYes Then

End If 'klop = vbNo Then
End If 'Else  'Msgbox NO
'Exit Sub '

              
              
              
'
'                    'To find the Total Invoice Amt
'                    Dim CounttoTotList As Long
'                    Dim FindTota, VArSum
'                    CounttoTotList = Me.ListView1.ListItems.Count
'
'                    For i = 1 To CounttoTotList
'                    FindTota = Me.ListView1.ListItems(i).SubItems(1)
'                    VArSum = Val(VArSum) + FindTota
'
'                    Next
'
'
              
               
               

  
'---------------------------
  





'This is to Update the ListView Again (It will delete all the Existing and Save it Back

              n = 0
              For n = 1 To Me.ListView1.ListItems.Count
                  LINVNo = Me.ListView1.ListItems.Item(n)
                  Linvamt = Me.ListView1.ListItems.Item(n).SubItems(1)
                  LinvDate = Me.ListView1.ListItems.Item(n).SubItems(2)
                  LDueDate = Me.ListView1.ListItems.Item(n).SubItems(3)
                  LtradDis = Me.ListView1.ListItems.Item(n).SubItems(4)
                  LCAshDis = Me.ListView1.ListItems.Item(n).SubItems(5)
                  LSalsTax = Me.ListView1.ListItems.Item(n).SubItems(6)
                  LServTax = Me.ListView1.ListItems.Item(n).SubItems(7)
                  ComPro = Me.ListView1.ListItems.Item(n).SubItems(8)
                  Addded = Me.ListView1.ListItems.Item(n).SubItems(9)
                  CustDut = Me.ListView1.ListItems.Item(n).SubItems(10)
                  TotInv = Me.ListView1.ListItems.Item(n).SubItems(11)
                  LPONo = Me.ListView1.ListItems.Item(n).SubItems(12)
                  LpoDate = Me.ListView1.ListItems.Item(n).SubItems(13)
                  LSENo = Me.ListView1.ListItems.Item(n).SubItems(14)
                  LSEDate = Me.ListView1.ListItems.Item(n).SubItems(15)
                  
              
              Dim xPayInvoiceDetails2 As New ADODB.Recordset
              xPayInvoiceDetails2.Open "select * from PayInvoiceDetails", constring, adOpenDynamic, adLockOptimistic
    
                  
                     With xPayInvoiceDetails2
                           .AddNew
                            !SerialNo = FrmPayableSetup.txtSerialNo.Text
                            !InvNo = LINVNo
                            !InvDate = LinvDate
                            !duedate = LDueDate
                            !invAmt = Linvamt
                            !TradeDis = LtradDis
                            !CashDis = LCAshDis
                            !SalesDis = LSalsTax
                            !ServDis = LServTax
                            !ComProf = ComPro
                            !AddDedTax1 = Addded
                            !CustDut = CustDut
                            !TotInv = TotInv
                            !PoNumber = LPONo
                            !PODate = LpoDate
                            !SENumber = LSENumber
                            !SEDate = LSEDate
                       On Error Resume Next
                           .Update
                      End With
               Next
             FrmPayableSetup.txtInvAmt.Text = Val(Pilappalam) '+ Val(txtSuT.Text)
             
             MsgBox "records Has Been Added Successfully ", vbInformation, "SAVE"
             
             
                  
                    'To find the Total Invoice Amt
                    Dim CounttoTotList3 As Long
                    Dim FindTota3, VArSum3
                    CounttoTotList3 = Me.ListView1.ListItems.Count
                    
                    For i = 1 To CounttoTotList3
                    FindTota3 = Me.ListView1.ListItems(i).SubItems(1)
                    VArSum3 = Val(VArSum3) + FindTota3
                    
                    Next
 
             
             
             
FrmPayableSetup.txtInvNo = Me.txtInvNo
FrmPayableSetup.txtInvAmt = Format(VArSum3, "###,###,###.#0")
FrmPayableSetup.Command3.Enabled = True
FrmPayableSetup.txtInvAmt.Enabled = True
FrmPayableSetup.txtInvNo.Enabled = True
Unload Me

End Sub

Private Sub Command2_Click()
If Command2.caption = "E&xit" Then
Unload Me
Else
cmdsave.caption = "&Culculate Çáå ÍÇÓÈÉ "
For Each Control In Me
    If TypeOf Control Is TextBox Then
        Control.Text = ""
    End If
Next

txtRTD.Text = 0
txtRCD.Text = 0
txtRST.Text = 0
txtRSrT.Text = 0
txtRComPro.Text = 0
txtRSurT.Text = 0
txtRCusD.Text = 0
Command2.caption = "E&xit"
End If
End Sub
Private Sub InvUpdateAdditional()
    If Check1.Value = 0 Then
            GrossCost = Val(txtGross.Text)
            txtGross.Text = Format(GrossCost, "###,###,###.#0")
            
            TradeDis = Val(GrossCost) * txtRTD / 100
            txttrdis.Text = Format(TradeDis, "###,###,###.#0")
            
            
            CashDis = Val(GrossCost) * txtRCD / 100
            txtCashDis.Text = Format(CashDis, "###,###,###.#0")
            
            netCostofGood = Val(GrossCost) - Val(CashDis) - Val(TradeDis)
            
            txtNet.Text = Format(netCostofGood, "###,###,###.#0") 'Net cost of goods
            
            SalesTax = Val(netCostofGood) * txtRST / 100
            txtST.Text = Format(SalesTax, "###,###,###.#0")
            
            saviceTax = Val(netCostofGood) * txtRSrT / 100
            txtSrT.Text = Format(saviceTax, "###,###,###.#0")
            
            ComProTax = Val(netCostofGood) * txtRComPro / 100
            txtCPro.Text = Format(ComProTax, "###,###,###.#0")
            
            SurTax = Val(netCostofGood) * txtRSurT / 100
            txtSuT.Text = Format(SurTax, "###,###,###.#0")
            
            CustDue = Val(netCostofGood) * txtRCusD / 100
            txtCuDut.Text = Format(CustDue, "###,###,###.#0")
            
            TotInvAmt = netCostofGood + Val(SalesTax) + Val(saviceTax) + Val(ComProTax) + Val(CustDue) '+ Val(SurTax)
            
            txtTotalInvoice = Format(TotInvAmt, "###,###,###.#0")
            
            Me.Command2.caption = "Cancel"

    ElseIf Check1.Value = 1 Then
    
            TotInvAmt = Val(txtGross.Text) - Val(txttrdis.Text) - Val(txtCashDis.Text) + Val(txtST.Text) + Val(txtSrT.Text) + Val(txtCPro.Text) + Val(txtCuDut.Text) '- Val(txtSuT.Text)
        
            txtTotalInvoice = Format(TotInvAmt, "###,###,###.#0")
            
            Me.Command2.caption = "Cancel"

    End If


'-----------------------------------------------------------------


rosa = MsgBox("Do you want to Add the Multi Invoice", vbYesNo + vbInformation, "Confirmation")
If rosa = vbYes Then
     Set MItem = Me.ListView1.ListItems.Add(, , Trim(txtInvNo.Text))
     MItem.SubItems(1) = Trim(txtGross.Text)
     MItem.SubItems(2) = Trim(txtINVdate.Text)
     MItem.SubItems(3) = Trim(mskDueDate.Text)
     MItem.SubItems(4) = Trim(txttrdis.Text)
     MItem.SubItems(5) = Trim(txtCashDis.Text)
     MItem.SubItems(6) = Trim(txtST.Text)
     MItem.SubItems(7) = Trim(txtSrT.Text)
     MItem.SubItems(8) = Trim(txtCPro.Text)
     MItem.SubItems(9) = Trim(txtSuT.Text) 'Tax Credit
     MItem.SubItems(10) = Trim(txtCuDut.Text)
     MItem.SubItems(11) = Trim(txtTotalInvoice.Text)
     MItem.SubItems(12) = Trim(txtPoNo.Text)
     MItem.SubItems(13) = Trim(mskPOdate.Text)
     MItem.SubItems(14) = Trim(txtSENo.Text)
     MItem.SubItems(15) = Trim(MskSEDate.Text)

     Pilappalam = Val(Pilappalam) + txtTotalInvoice.Text
     Lks = txtTotalInvoice.Text


     txtLastAmt.Text = Val(Pilappalam)
     txtINVdate.Text = "__/__/____"
     mskPOdate.Text = "__/__/____"
     txtInvNo.Text = ""
     mskDueDate.Text = "__/__/____"
     txtGross.Text = ""
     txttrdis.Text = ""
     txtCashDis.Text = ""
     txtST.Text = ""
     txtNet.Text = ""
     txtSrT.Text = ""
     txtCPro.Text = ""
     txtSuT.Text = ""
     txtCuDut.Text = ""
     txtTotalInvoice.Text = ""
     txtPoNo.Text = ""
     txtPoNo.Text = ""
txtGross.SetFocus
Else  'Msgbox NO
    klop = MsgBox("Do you want to Save the Details", vbYesNo, "Save")
         If klop = vbNo Then
            MsgBox "Nothing is saved you Lost all the details you Entered", vbInformation
            Exit Sub
            ElseIf klop = vbYes Then
        
             Set MItem = Me.ListView1.ListItems.Add(, , Trim(txtInvNo.Text))
             MItem.SubItems(1) = Trim(txtGross.Text)
             MItem.SubItems(2) = Trim(txtINVdate.Text)
             MItem.SubItems(3) = Trim(mskDueDate.Text)
             MItem.SubItems(4) = Trim(txttrdis.Text)
             MItem.SubItems(5) = Trim(txtCashDis.Text)
             MItem.SubItems(6) = Trim(txtST.Text)
             MItem.SubItems(7) = Trim(txtSrT.Text)
             MItem.SubItems(8) = Trim(txtCPro.Text)
             MItem.SubItems(9) = Trim(txtSuT.Text) 'tax Cre
             MItem.SubItems(10) = Trim(txtCuDut.Text)
             MItem.SubItems(11) = Trim(txtTotalInvoice.Text)
             MItem.SubItems(12) = Trim(txtPoNo.Text)
             MItem.SubItems(13) = Trim(mskPOdate.Text)
             MItem.SubItems(14) = Trim(txtSENo.Text)
             MItem.SubItems(15) = Trim(MskSEDate.Text)
        
             Lks = txtTotalInvoice.Text
             Pilappalam = Val(Pilappalam) + Val(Lks)
        
             
        
             txtINVdate.Text = "__/__/____"
             txtInvNo.Text = ""
             mskDueDate.Text = "__/__/____"
             mskPOdate.Text = "__/__/____"
             txtGross.Text = ""
             txttrdis.Text = ""
             txtCashDis.Text = ""
             txtST.Text = ""
             txtSrT.Text = ""
             txtCPro.Text = ""
             txtSuT.Text = ""
             txtCuDut.Text = ""
             txtTotalInvoice.Text = ""
             txtNet.Text = ""
             txtPoNo.Text = ""
             txtSENo.Text = ""
             Dim TotTaxCred
         
                
                      n = 0
                      For n = 1 To Me.ListView1.ListItems.Count
                          LINVNo = Me.ListView1.ListItems.Item(n)
                          Linvamt = Me.ListView1.ListItems.Item(n).SubItems(1)
                          LinvDate = Me.ListView1.ListItems.Item(n).SubItems(2)
                          LDueDate = Me.ListView1.ListItems.Item(n).SubItems(3)
                          LtradDis = Me.ListView1.ListItems.Item(n).SubItems(4)
                          LCAshDis = Me.ListView1.ListItems.Item(n).SubItems(5)
                          LSalsTax = Me.ListView1.ListItems.Item(n).SubItems(6)
                          LServTax = Me.ListView1.ListItems.Item(n).SubItems(7)
                          ComPro = Me.ListView1.ListItems.Item(n).SubItems(8)
                          Addded = Me.ListView1.ListItems.Item(n).SubItems(9)
                          CustDut = Me.ListView1.ListItems.Item(n).SubItems(10)
                          TotInv = Me.ListView1.ListItems.Item(n).SubItems(11)
                          LPONo = Me.ListView1.ListItems.Item(n).SubItems(12)
                          LpoDate = Me.ListView1.ListItems.Item(n).SubItems(13)
                          LSENo = Me.ListView1.ListItems.Item(n).SubItems(14)
                          LSEDate = Me.ListView1.ListItems.Item(n).SubItems(15)
                             
                      TotTaxCred = Val(TotTaxCred) + Val(Addded)
                           On Error GoTo 0
                           
                           Dim xPayInvoiceDetails2 As New ADODB.Recordset
                           xPayInvoiceDetails2.Open "select * from PayTempInvoiceDetails", CON1, adOpenDynamic, adLockOptimistic
                           With xPayInvoiceDetails2
                                   .AddNew
                                    !SerialNo = FrmPayableSetup.txtSerialNo.Text
                                    !InvNo = LINVNo
                                    !InvDate = LinvDate
                                    !duedate = LDueDate
                                    !invAmt = Linvamt
                                    !TradeDis = LtradDis
                                    !CashDis = LCAshDis
                                    !SalesDis = LSalsTax
                                    !ServDis = LServTax
                                    !ComProf = ComPro
                                    !AddDedTax1 = Addded
                                    !CustDut = CustDut
                                    !TotInv = TotInv
                                    !PoNumber = LPONo
                                    !PODate = LpoDate
                                    !SENumber = LSENo
                                    !SEDate = LSEDate
                               'On Error Resume Next
                                   .Update
                              End With
                              
                              
                       Next
                       
        ''This is to find the All Prior Payments for the Same Invoice and put it
        'Dim Prir
        'Prir = Trim(txtInvNo)
        'Dim PriorInv As New ADODB.Recordset
        'PriorInv.Open "select sum(InvAmt) as InvAmt from  PayInvoiceDetails  where invno = " & "'" & Prir & "'" & "", conString, adOpenDynamic, adLockOptimistic
        ''Mambalam = "SELECT CUSTNO, SUM(debitamount) AS cashamount, SUM(checkreceipt) AS checkamount, SUM(debitamount) + SUM(checkreceipt) AS totalamount From vouchers WHERE (deleted = '0') AND (receiptDate <= " & "'" & Format(mfrom, "mm/dd/yyyy") & "'" & ") AND (CUSTNO = " & "'" & inv & "'" & ") GROUP BY CUSTNO"
        '
        'txtPriorPayments.Text = IIf(IsNull(PriorInv!InvAmt), 0, (PriorInv!InvAmt))
        '
                      
                      
                      
              
                            'To find the Total Invoice Amt
                            Dim CounttoTotList As Long
                            Dim FindTota, VArSum
                            CounttoTotList = Me.ListView1.ListItems.Count
                            
                            For i = 1 To CounttoTotList
                            FindTota = Me.ListView1.ListItems(i).SubItems(1)
                            VArSum = Val(VArSum) + FindTota
                            
                            Next
                
                      
                      
                       
                       
        'FrmPayableSetup.txtInvAmt.Text = Val(Pilappalam) '- Val(txtPriorPayments.Text)
       ' FrmPayableSetup.txtInvAmt.Text = Val(VArSum)
        FrmPayableSetup.txtInvAmt.Text = Format(TotInvAmt, "###,###,###.#0")
        FrmPayableSetup.txtTaxCredit.Text = Val(TotTaxCred)
                     MsgBox "records Has Been Added Successfully ", vbInformation, "SAVE"
        
                FrmPayableSetup.txtInvNo = Me.txtInvNo
                FrmPayableSetup.txtInvAmt = Format(TotInvAmt, "###,###,###.#0")
                FrmPayableSetup.Command3.Enabled = True
                FrmPayableSetup.txtInvAmt.Enabled = True
                FrmPayableSetup.txtInvNo.Enabled = True
                'FrmPayableSetup.mskInvDate.SetFocus
        Unload Me
        End If
        End If


End Sub

Private Sub Form_Activate()
Me.txtSerialNo.Text = FrmPayableSetup.txtSerialNo.Text
End Sub

Private Sub Form_Load()

Dim CON1 As ADODB.Connection
Set CON1 = New ADODB.Connection
    

conStr = "Provider=MSDASQL;DSN=Finance;UID=sa; PWD=;"

CON1.Open conStr

Dim Nathe
Dim rstNAt As New ADODB.Recordset
Dim PLO
Nathe = FrmPayableSetup.txtSerialNo.Text

PLO = "Select * from Payinvoicedetails where SerialNo = " & "'" & Nathe & "'" & ""
rstNAt.Open PLO, CON1, adOpenDynamic, adLockOptimistic

txtRTD.Text = 0
txtRCD.Text = 0
txtRST.Text = 0
txtRSrT.Text = 0
txtRComPro.Text = 0
txtRSurT.Text = 0
txtRCusD.Text = 0


Set ARRahman = Me.ListView1.ColumnHeaders.Add(, , "Inv No")
Set ARRahman = Me.ListView1.ColumnHeaders.Add(, , "Inv Amount")
Set ARRahman = Me.ListView1.ColumnHeaders.Add(, , "Inv Date")
Set ARRahman = Me.ListView1.ColumnHeaders.Add(, , "Due Date")
'Set ARRahman = Me.ListView1.ColumnHeaders.Add(, , "Inv Amount")
Set ARRahman = Me.ListView1.ColumnHeaders.Add(, , "Trade Discount")
Set ARRahman = Me.ListView1.ColumnHeaders.Add(, , "Cash Discount")
Set ARRahman = Me.ListView1.ColumnHeaders.Add(, , "Sales Tax")
Set ARRahman = Me.ListView1.ColumnHeaders.Add(, , "Service Tax")
Set ARRahman = Me.ListView1.ColumnHeaders.Add(, , "Com Profit")
Set ARRahman = Me.ListView1.ColumnHeaders.Add(, , "Tax Credit") '"Add/Ded Tax")
Set ARRahman = Me.ListView1.ColumnHeaders.Add(, , "Cust Details")
Set ARRahman = Me.ListView1.ColumnHeaders.Add(, , "Total Invoice")
Set ARRahman = Me.ListView1.ColumnHeaders.Add(, , "PO Number")
Set ARRahman = Me.ListView1.ColumnHeaders.Add(, , "PO Date")
Set ARRahman = Me.ListView1.ColumnHeaders.Add(, , "SE Number")
Set ARRahman = Me.ListView1.ColumnHeaders.Add(, , "SE Date")

Me.ListView1.ListItems.Clear
  While rstNAt.EOF = False
  
     Set MItem = Me.ListView1.ListItems.Add(, , Format(rstNAt!InvNo))
     MItem.SubItems(1) = Format(rstNAt!invAmt)
     MItem.SubItems(2) = Format(rstNAt!InvDate, "dd/mm/yyyy")
     MItem.SubItems(3) = Format(rstNAt!duedate, "dd/mm/yyyy")
'     mitem.SubItems(4) = Format(rstNAt!invAmt)
     MItem.SubItems(4) = Format(rstNAt!TradeDis)
     MItem.SubItems(5) = Format(rstNAt!CashDis)
     MItem.SubItems(6) = Format(rstNAt!SalesDis)
     MItem.SubItems(7) = Format(rstNAt!ServDis)
     MItem.SubItems(8) = Format(rstNAt!ComProf)
     MItem.SubItems(9) = Format(rstNAt!AddDedTax1)
     MItem.SubItems(10) = Format(rstNAt!CustDut)
     MItem.SubItems(11) = Format(rstNAt!TotInv)
     MItem.SubItems(12) = Format(rstNAt!PoNumber)
     MItem.SubItems(13) = Format(rstNAt!PODate)
     MItem.SubItems(14) = Format(rstNAt!SENumber)
     MItem.SubItems(15) = Format(rstNAt!SEDate)
 
     rstNAt.MoveNext
     Wend




End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  ListView1.SortKey = ColumnHeader.Index - 1
 ' Set Sorted to True to sort the list.
   ListView1.Sorted = True

End Sub

Private Sub ListView1_DblClick()



        If Me.ListView1.ListItems.Count = 0 Then
        MsgBox "No items Selected", vbInformation
        Exit Sub
        End If
      
      frmMenu.InvEdit.caption = "Update"
      cmdUpdate.Visible = True

       txtInvNo.Text = Me.ListView1.SelectedItem
        txtGross.Text = Me.ListView1.SelectedItem.SubItems(1)
        txtINVdate.Text = Me.ListView1.SelectedItem.SubItems(2)
        mskDueDate.Text = Me.ListView1.SelectedItem.SubItems(3)
        txttrdis.Text = Me.ListView1.SelectedItem.SubItems(4)
        txtCashDis.Text = Me.ListView1.SelectedItem.SubItems(5)
        txtST.Text = Me.ListView1.SelectedItem.SubItems(6)
        txtSrT.Text = Me.ListView1.SelectedItem.SubItems(7)
        txtCPro.Text = Me.ListView1.SelectedItem.SubItems(8)
        txtSuT.Text = Me.ListView1.SelectedItem.SubItems(9)
        txtCuDut.Text = Me.ListView1.SelectedItem.SubItems(10)
        txtTotalInvoice.Text = Me.ListView1.SelectedItem.SubItems(11)
        txtPoNo.Text = Me.ListView1.SelectedItem.SubItems(12)
        mskPOdate.Text = Me.ListView1.SelectedItem.SubItems(13)
        txtSENo.Text = Me.ListView1.SelectedItem.SubItems(14)
        On Error Resume Next
        MskSEDate.Text = IIf(IsNull(Me.ListView1.SelectedItem.SubItems(15)), "__/__/____", (IsNull(Me.ListView1.SelectedItem.SubItems(15))))
        On Error GoTo 0

'cmdSave.Caption = "&Update"
cmdsave.Visible = False


End Sub

Private Sub MSKdueDate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtSENo.SetFocus
End If
End Sub

Private Sub MSKdueDate_LostFocus()
Me.mskDueDate.Text = Format(Me.mskDueDate.Text, "dd/mm/yyyy")

End Sub

Private Sub mskPOdate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtGross.SetFocus
End If
End Sub

Private Sub mskPODate_LostFocus()
Me.mskPOdate.Text = Format(Me.mskPOdate.Text, "dd/mm/yyyy")
End Sub

Private Sub MskSEDate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtPoNo.SetFocus
End If

End Sub

Private Sub MskSEDate_LostFocus()
Me.MskSEDate.Text = Format(Me.MskSEDate.Text, "dd/mm/yyyy")

End Sub

Private Sub txtCashDis_KeyUp(KeyCode As Integer, Shift As Integer)
 
If KeyCode = 13 Or KeyCode = 17 Or KeyCode = 67 Or KeyCode = 86 Or KeyCode = 8 Or KeyCode = 18 Or KeyCode = 46 Or KeyCode = 27 Then

Exit Sub
End If
 
 
 If KeyCode <> 46 Then
 
 If KeyCode <> 13 Then
 If KeyCode <> 110 Then
 If KeyCode < 96 Or KeyCode > 105 Then

 If KeyCode < 48 Or KeyCode > 57 Then
    MsgBox "The Key you entered is not the Digit", vbInformation, "Try Again"
    Me.txtCashDis.Text = 0
    Exit Sub
End If
End If
End If
End If
End If
End Sub

Private Sub txtCPro_KeyUp(KeyCode As Integer, Shift As Integer)


If KeyCode = 13 Or KeyCode = 17 Or KeyCode = 67 Or KeyCode = 86 Or KeyCode = 8 Or KeyCode = 18 Or KeyCode = 46 Or KeyCode = 27 Then

Exit Sub
End If


'If KeyCode <> 46 Then
If KeyCode <> 13 Then
 If KeyCode <> 110 Then
 If KeyCode < 96 Or KeyCode > 105 Then

 If KeyCode < 48 Or KeyCode > 57 Then
    MsgBox "The Key you entered is not the Digit", vbInformation, "Try Again"
    Me.txtCPro = 0
    Exit Sub
End If
End If
End If
End If
'End If

End Sub

Private Sub txtCuDut_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Or KeyCode = 17 Or KeyCode = 67 Or KeyCode = 86 Or KeyCode = 8 Or KeyCode = 18 Or KeyCode = 46 Or KeyCode = 27 Then

Exit Sub
End If


If KeyCode <> 13 Then
 If KeyCode <> 110 Then
 If KeyCode < 96 Or KeyCode > 105 Then

 If KeyCode < 48 Or KeyCode > 57 Then
    MsgBox "The Key you entered is not the Digit", vbInformation, "Try Again"
    Me.txtCuDut.Text = 0
    Exit Sub
End If
End If
End If
End If

End Sub

Private Sub txtGross_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
cmdsave.SetFocus
End If
On Error GoTo 0
End Sub

Private Sub txtGross_KeyUp(KeyCode As Integer, Shift As Integer)
 
If KeyCode = 13 Or KeyCode = 17 Or KeyCode = 67 Or KeyCode = 86 Or KeyCode = 8 Or KeyCode = 18 Or KeyCode = 46 Or KeyCode = 27 Then

Exit Sub
End If
 
 
 If KeyCode <> 13 Then
 If KeyCode <> 110 Then
 If KeyCode < 96 Or KeyCode > 105 Then

 If KeyCode < 48 Or KeyCode > 57 Then
    MsgBox "The Key you entered is not the Digit", vbInformation, "Try Again"
    Me.txtGross.Text = 0
    Exit Sub
End If
End If
End If
End If

End Sub

Private Sub txtINVdate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
mskDueDate.SetFocus
End If
End Sub

Private Sub txtINVdate_LostFocus()
Me.txtINVdate.Text = Format(Me.txtINVdate.Text, "dd/mm/yyyy")

End Sub

Private Sub txtinvNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtINVdate.SetFocus
End If
End Sub

Private Sub txtinvNo_LostFocus()
tp = Me.ListView1.ListItems.Count
For i = 1 To tp
mycrvar = Me.ListView1.ListItems.Item(i)

    If Me.txtInvNo = mycrvar Then
     xmsg = MsgBox("You have entered same Invoice number ", vbInformation + vbOKOnly, "Message")
     Me.txtInvNo.SetFocus
     Exit Sub
    End If
Next


'This is to get all the related same invoice for this invoive no, if it is already in the list

'Dim PriorInv As New ADODB.Recordset
'PriorInv.Open "select * from PayInvoiceDetails where invno = " & "'" & Trim(txtInvNo) & "'" & " and Payee"



'----------------------------------


'This is to get the All Prior Payments by checking the Previos datas
'Dim PriorPAym As New ADODB.Recordset




End Sub

Private Sub txtPoNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
mskPOdate.SetFocus

End If
End Sub

Private Sub txtRCD_KeyUp(KeyCode As Integer, Shift As Integer)
 
If KeyCode = 13 Or KeyCode = 17 Or KeyCode = 67 Or KeyCode = 86 Or KeyCode = 8 Or KeyCode = 18 Or KeyCode = 46 Or KeyCode = 27 Then

Exit Sub
End If
 
 
 
 If KeyCode <> 13 Then
 If KeyCode <> 110 Then
 If KeyCode < 96 Or KeyCode > 105 Then

 If KeyCode < 48 Or KeyCode > 57 Then
    MsgBox "The Key you entered is not the Digit", vbInformation, "Try Again"
    Me.txtRCD.Text = 0
    Exit Sub
End If
End If
End If
End If

End Sub

Private Sub txtRComPro_KeyUp(KeyCode As Integer, Shift As Integer)
 
If KeyCode = 13 Or KeyCode = 17 Or KeyCode = 67 Or KeyCode = 86 Or KeyCode = 8 Or KeyCode = 18 Or KeyCode = 46 Or KeyCode = 27 Then

Exit Sub
End If
 
 
 If KeyCode <> 13 Then
 If KeyCode <> 110 Then
 If KeyCode < 96 Or KeyCode > 105 Then

 If KeyCode < 48 Or KeyCode > 57 Then
    MsgBox "The Key you entered is not the Digit", vbInformation, "Try Again"
    Me.txtRComPro.Text = 0
    Exit Sub
End If
End If
End If
End If

End Sub

Private Sub txtRCusD_KeyUp(KeyCode As Integer, Shift As Integer)
 
If KeyCode = 13 Or KeyCode = 17 Or KeyCode = 67 Or KeyCode = 86 Or KeyCode = 8 Or KeyCode = 18 Or KeyCode = 46 Or KeyCode = 27 Then

Exit Sub
End If
 
 
 If KeyCode <> 13 Then
 If KeyCode <> 110 Then
 If KeyCode < 96 Or KeyCode > 105 Then

 If KeyCode < 48 Or KeyCode > 57 Then
    MsgBox "The Key you entered is not the Digit", vbInformation, "Try Again"
    Me.txtRCusD.Text = 0
    Exit Sub
End If
End If
End If
End If

End Sub

Private Sub txtRSrT_KeyUp(KeyCode As Integer, Shift As Integer)
 
If KeyCode = 13 Or KeyCode = 17 Or KeyCode = 67 Or KeyCode = 86 Or KeyCode = 8 Or KeyCode = 18 Or KeyCode = 46 Or KeyCode = 27 Then

Exit Sub
End If
 
 
 If KeyCode <> 13 Then
 If KeyCode <> 110 Then
 If KeyCode < 96 Or KeyCode > 105 Then

 If KeyCode < 48 Or KeyCode > 57 Then
    MsgBox "The Key you entered is not the Digit", vbInformation, "Try Again"
    Me.txtRSrT.Text = 0
    Exit Sub
End If
End If
End If
End If

End Sub

Private Sub txtRST_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 17 Or KeyCode = 67 Or KeyCode = 86 Or KeyCode = 8 Or KeyCode = 18 Or KeyCode = 46 Or KeyCode = 27 Then

Exit Sub
End If
 
 
 If KeyCode <> 13 Then
 If KeyCode <> 110 Then
 If KeyCode < 96 Or KeyCode > 105 Then

 If KeyCode < 48 Or KeyCode > 57 Then
    MsgBox "The Key you entered is not the Digit", vbInformation, "Try Again"
    Me.txtRST.Text = 0
    Exit Sub
End If
End If
End If
End If

End Sub

Private Sub txtRSurT_KeyUp(KeyCode As Integer, Shift As Integer)
 
If KeyCode = 13 Or KeyCode = 17 Or KeyCode = 67 Or KeyCode = 86 Or KeyCode = 8 Or KeyCode = 18 Or KeyCode = 46 Or KeyCode = 27 Then

Exit Sub
End If
 
 
 If KeyCode <> 13 Then
 If KeyCode <> 110 Then
 If KeyCode < 96 Or KeyCode > 105 Then

 If KeyCode < 48 Or KeyCode > 57 Then
    MsgBox "The Key you entered is not the Digit", vbInformation, "Try Again"
    Me.txtRSurT.Text = 0
    Exit Sub
End If
End If
End If
End If
End Sub

Private Sub txtRTD_KeyUp(KeyCode As Integer, Shift As Integer)
 
If KeyCode = 13 Or KeyCode = 17 Or KeyCode = 67 Or KeyCode = 86 Or KeyCode = 8 Or KeyCode = 18 Or KeyCode = 46 Or KeyCode = 27 Then

Exit Sub
End If
 
 
 If KeyCode <> 13 Then
 If KeyCode <> 110 Then
 If KeyCode < 96 Or KeyCode > 105 Then

 If KeyCode < 48 Or KeyCode > 57 Then
    MsgBox "The Key you entered is not the Digit", vbInformation, "Try Again"
    Me.txtRTD.Text = 0
    Exit Sub
End If
End If
End If
End If

End Sub

Private Sub txtSENo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
MskSEDate.SetFocus
End If

End Sub

Private Sub txtSrT_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Or KeyCode = 17 Or KeyCode = 67 Or KeyCode = 86 Or KeyCode = 8 Or KeyCode = 18 Or KeyCode = 46 Or KeyCode = 27 Then

Exit Sub
End If



If KeyCode <> 13 Then
 If KeyCode <> 110 Then
 If KeyCode < 96 Or KeyCode > 105 Then

 If KeyCode < 48 Or KeyCode > 57 Then
    MsgBox "The Key you entered is not the Digit", vbInformation, "Try Again"
    Me.txtSrT.Text = 0
    Exit Sub
End If
End If
End If
End If

End Sub

Private Sub txtSuT_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Or KeyCode = 17 Or KeyCode = 67 Or KeyCode = 86 Or KeyCode = 8 Or KeyCode = 18 Or KeyCode = 46 Or KeyCode = 27 Then

Exit Sub
End If


If KeyCode <> 13 Then
 If KeyCode <> 110 Then
 If KeyCode < 96 Or KeyCode > 105 Then

 If KeyCode < 48 Or KeyCode > 57 Then
    MsgBox "The Key you entered is not the Digit", vbInformation, "Try Again"
    Me.txtSuT.Text = 0
    Exit Sub
End If
End If
End If
End If

End Sub

Private Sub txtTotalInvoice_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
rosa = MsgBox("Do you want to Add the Multi Invoice", vbYesNo + vbInformation, "Confirmation")
End If

If rosa = vbYes Then
     Set MItem = Me.ListView1.ListItems.Add(, , Trim(txtInvNo.Text))
     MItem.SubItems(1) = Trim(txtGross.Text)
     MItem.SubItems(2) = Trim(txtINVdate.Text)
     MItem.SubItems(3) = Trim(mskDueDate.Text)
     
     MItem.SubItems(4) = Trim(txttrdis.Text)
     MItem.SubItems(5) = Trim(txtCashDis.Text)
     MItem.SubItems(6) = Trim(txtST.Text)
     MItem.SubItems(7) = Trim(txtSrT.Text)
     MItem.SubItems(8) = Trim(txtCPro.Text)
     MItem.SubItems(9) = Trim(txtSuT.Text)
     MItem.SubItems(10) = Trim(txtCuDut.Text)
     MItem.SubItems(11) = Trim(txtTotalInvoice.Text)
     MItem.SubItems(12) = Trim(txtPoNo.Text)
     MItem.SubItems(13) = Trim(mskPOdate.Text)
     MItem.SubItems(14) = Trim(txtSENo.Text)
     MItem.SubItems(15) = Trim(MskSEDate.Text)
 
     
     
     
     
     
     mskDueDate.Text = ""
     txtINVdate.Text = ""
     txtInvNo.Text = ""
     txtGross.Text = ""
     txttrdis.Text = ""
     txtCashDis.Text = ""
     txtST.Text = ""
     txtSrT.Text = ""
     txtCPro.Text = ""
     txtSuT.Text = ""
     txtCuDut.Text = ""
     txtTotalInvoice.Text = ""
     txtPoNo.Text = ""
     mskPOdate.Text = ""
End If

End Sub

Private Sub txttrdis_KeyUp(KeyCode As Integer, Shift As Integer)
 
 
 If KeyCode = 13 Or KeyCode = 17 Or KeyCode = 67 Or KeyCode = 86 Or KeyCode = 8 Or KeyCode = 18 Or KeyCode = 46 Or KeyCode = 27 Then

Exit Sub
End If

 
 If KeyCode <> 46 Then
 If KeyCode <> 13 Then
 If KeyCode <> 110 Then
 If KeyCode < 96 Or KeyCode > 105 Then

 If KeyCode < 48 Or KeyCode > 57 Then
    MsgBox "The Key you entered is not the Digit", vbInformation, "Try Again"
    Me.txttrdis.Text = 0
    Exit Sub
End If
End If
End If
End If
End If
End Sub

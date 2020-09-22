VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FrmPaymentAnalysis 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payment Analysis"
   ClientHeight    =   7710
   ClientLeft      =   135
   ClientTop       =   435
   ClientWidth     =   11400
   Icon            =   "frmPaymentAnalysis.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   11400
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPaymentLevel 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   68
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Set to Arabic ÇáÊÍæíá ááÛÉ ÇáÚÑÈíÉ"
      Height          =   255
      Left            =   7080
      TabIndex        =   63
      Top             =   1440
      Width           =   2775
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   9000
      TabIndex        =   62
      Text            =   "Text4"
      Top             =   3600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3073
         SubFormatType   =   1
      EndProperty
      Height          =   330
      Left            =   8160
      Locked          =   -1  'True
      TabIndex        =   43
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox Text2 
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
      Height          =   330
      Left            =   8160
      Locked          =   -1  'True
      TabIndex        =   42
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox Text1 
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
      Height          =   330
      Left            =   8160
      Locked          =   -1  'True
      TabIndex        =   41
      Top             =   240
      Width           =   1575
   End
   Begin VB.TextBox txtIdentifyNewData 
      Height          =   375
      Left            =   9000
      TabIndex        =   40
      Top             =   2640
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox txtAmtreq 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   38
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox txtSerialNo 
      Height          =   325
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   34
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Left            =   10080
      TabIndex        =   33
      ToolTipText     =   "Add  New Entry"
      Top             =   6600
      Width           =   1005
   End
   Begin VB.CommandButton cmdExit1 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Left            =   10080
      TabIndex        =   32
      ToolTipText     =   "Close Window"
      Top             =   7320
      Width           =   1005
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "&Edit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Left            =   10080
      TabIndex        =   31
      ToolTipText     =   "Print Payable Setup"
      Top             =   6960
      Width           =   1005
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Go Payable Setup.."
      Height          =   255
      Left            =   10440
      TabIndex        =   30
      Top             =   7320
      Width           =   135
   End
   Begin VB.ComboBox CmbApprovedBy 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9120
      TabIndex        =   28
      Top             =   5880
      Width           =   2055
   End
   Begin VB.ComboBox CmbNotedBy 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9120
      TabIndex        =   26
      Top             =   5280
      Width           =   2055
   End
   Begin VB.ComboBox CmbPrepBy 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9120
      TabIndex        =   24
      Top             =   4680
      Width           =   2055
   End
   Begin VB.Frame Frame3 
      Caption         =   "D E B I T "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   8655
      Begin VB.TextBox txtTotList3 
         Alignment       =   1  'Right Justify
         Height          =   320
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox txtTemp2List3 
         Height          =   285
         Left            =   2160
         TabIndex        =   22
         Text            =   "Text2"
         Top             =   1200
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtTemp1List3 
         Height          =   285
         Left            =   840
         TabIndex        =   21
         Text            =   "Text2"
         Top             =   1200
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox txtPartic 
         Height          =   315
         Left            =   2040
         TabIndex        =   5
         Top             =   480
         Width           =   4665
      End
      Begin VB.ComboBox txtAccNo 
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox txtAmo 
         Alignment       =   1  'Right Justify
         Height          =   320
         Left            =   7200
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H8000000A&
         Height          =   315
         Left            =   6720
         Picture         =   "frmPaymentAnalysis.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Find Accounts"
         Top             =   480
         Width           =   375
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   1215
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   2143
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         OLEDragMode     =   1
         OLEDropMode     =   1
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         OLEDragMode     =   1
         OLEDropMode     =   1
         NumItems        =   0
      End
      Begin VB.Label Label10 
         Caption         =   "Total DB Amount"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4200
         TabIndex        =   65
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÅÌãÇáí ÇáØÑÝ ÇáãÏíä"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5400
         TabIndex        =   64
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÇáãÈáÛ"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   7680
         TabIndex        =   56
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÅÓã ÇáÍÓÇÈ"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5160
         TabIndex        =   54
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÑÞã ÇáÍÓÇÈ"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1200
         TabIndex        =   52
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label12 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Account Code"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   " Amount"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   7200
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label32 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Account Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2040
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "C R E D I T "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   4800
      Width           =   8655
      Begin VB.TextBox txtTemp2List6 
         Height          =   285
         Left            =   2160
         TabIndex        =   19
         Text            =   "Text2"
         Top             =   1440
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtTemp1List6 
         Height          =   285
         Left            =   1080
         TabIndex        =   18
         Text            =   "Text2"
         Top             =   1440
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtTotList6 
         Alignment       =   1  'Right Justify
         Height          =   320
         Left            =   7080
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   2400
         Width           =   1455
      End
      Begin VB.ComboBox txtDBPartic 
         Height          =   315
         Left            =   2040
         TabIndex        =   12
         Top             =   480
         Width           =   4695
      End
      Begin VB.ComboBox txtDBAccNo 
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox txtDBAmo 
         Alignment       =   1  'Right Justify
         Height          =   320
         Left            =   7200
         TabIndex        =   10
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H8000000A&
         Height          =   315
         Left            =   6720
         Picture         =   "frmPaymentAnalysis.frx":09DC
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Find Accounts"
         Top             =   480
         Width           =   375
      End
      Begin MSComctlLib.ListView ListView6 
         Height          =   1215
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   2143
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         OLEDragMode     =   1
         OLEDropMode     =   1
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         OLEDragMode     =   1
         OLEDropMode     =   1
         NumItems        =   0
      End
      Begin VB.Label Label15 
         Caption         =   "Total CR Amount"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   67
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " ÅÌãÇáí ÇáØÑÝ ÇáÏÇÆä"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5280
         TabIndex        =   66
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÇáãÈáÛ"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   7680
         TabIndex        =   55
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÅÓã ÇáÍÓÇÈ"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5040
         TabIndex        =   53
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÑÞã ÇáÍÓÇÈ"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1080
         TabIndex        =   51
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Account Code"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label3 
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   7200
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Account Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2040
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   330
      Left            =   1800
      TabIndex        =   35
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label17 
      Caption         =   "Payment Level"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   69
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label54 
      Alignment       =   1  'Right Justify
      Caption         =   "ÇáÊÇÑíÎ"
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
      Left            =   3480
      TabIndex        =   61
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label57 
      Alignment       =   1  'Right Justify
      Caption         =   "ÑÞã ÇáÊÓáÓá "
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
      Left            =   3600
      TabIndex        =   60
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label23 
      Caption         =   "ÇáÊæÞíÚ ÈæÇÓØÉ"
      Height          =   255
      Left            =   10080
      TabIndex        =   59
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Label Label48 
      AutoSize        =   -1  'True
      Caption         =   "íÊÍÞÞ ÈæÇÓØÉ "
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
      TabIndex        =   58
      Top             =   5040
      Width           =   960
   End
   Begin VB.Label Label20 
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
      Left            =   10080
      RightToLeft     =   -1  'True
      TabIndex        =   57
      Top             =   4440
      Width           =   960
   End
   Begin VB.Label Label33 
      Caption         =   "Invoice Amount"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6960
      TabIndex        =   50
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label22 
      Caption         =   "Outs-ing Bal"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6960
      TabIndex        =   49
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label43 
      Alignment       =   1  'Right Justify
      Caption         =   "ãÌãæÚ ÇáÝÇÊæÑÉ "
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
      Left            =   9840
      TabIndex        =   48
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label46 
      Alignment       =   1  'Right Justify
      Caption         =   "ÕÇÝí ÇáÑÕíÏ"
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
      Left            =   9840
      TabIndex        =   47
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label47 
      Alignment       =   1  'Right Justify
      Caption         =   "ãÌãæÚ ÇáÏÝÚ "
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
      Left            =   9840
      TabIndex        =   46
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label21 
      Caption         =   "Paid  Before"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6960
      TabIndex        =   45
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label40 
      Alignment       =   1  'Right Justify
      Caption         =   "ãÌãæÚ ÇáØáÈ"
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
      Left            =   3480
      TabIndex        =   44
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Amount Requested"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   39
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   37
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Serial Number"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   36
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label29 
      Caption         =   "Approved By"
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
      Left            =   9120
      TabIndex        =   29
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Label Label28 
      Caption         =   "Noted By"
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
      Left            =   9120
      TabIndex        =   27
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label Label27 
      Caption         =   "Prepared By"
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
      Left            =   9120
      TabIndex        =   25
      Top             =   4440
      Width           =   1815
   End
End
Attribute VB_Name = "FrmPaymentAnalysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn1 As New ADODB.Connection

Dim xItem As ListItem
'Dim SQLtable As Boolean
Dim CON1 As ADODB.Connection
'Dim rstSource As ADODB.Recordset
Dim rstEmp As ADODB.Recordset
'Dim rstVen As ADODB.Recordset
'Dim rstrate As ADODB.Recordset
'Dim rstPAyFor As ADODB.Recordset
'Dim rstCosPro As ADODB.Recordset
'Dim strVen As String
'Dim serial As String
Dim rstPaySetup As ADODB.Recordset
Dim rstPayment As ADODB.Recordset
Dim rstReceipt As ADODB.Recordset
Dim rstpayment2 As New ADODB.Recordset

'Dim rstterm As ADODB.Recordset
Dim rstChart As ADODB.Recordset
'Dim rstJournal As ADODB.Recordset
'dim RstPaid As ADODB.Recordset
Dim TotList3


Dim catName As String
Dim Prevcap As String
Dim acctNo As String



Public Sub UpdateList6XRecei()
 If rstReceipt.EOF = False Then
 
   rstReceipt.MoveFirst
      While rstReceipt.EOF = False
       Dim xRvar As String
      ' xRvar = Trim(rstReceipt!accno)
         If Trim(rstReceipt!SerialNo) = Trim(FrmPaymentAnalysis.txtSerialNo.Text) And FrmPaymentAnalysis.txtTemp1List6.Text = Trim(rstReceipt!AccNo) And FrmPaymentAnalysis.txtTemp2List6.Text = Trim(rstReceipt!AccName) Then
        rstReceipt!SerialNo = FrmPaymentAnalysis.txtSerialNo.Text
        rstReceipt!AccNo = FrmPaymentAnalysis.txtDBAccNo.Text
        rstReceipt!AccName = FrmPaymentAnalysis.txtDBPartic.Text
        rstReceipt!amount = FrmPaymentAnalysis.txtDBAmo.Text
        rstReceipt!total = FrmPaymentAnalysis.txtTotList6.Text
        frmMenu.Edit.caption = "&Edit"
        MyRval = "Gotit"
      End If
   rstReceipt.MoveNext
   Wend
If MyRval = "Gotit" Then
MsgBox "Records for Receipt Details Updated Seccussfully"
End If
rstReceipt.Close
End If
On Error Resume Next
rstReceipt.Open
'Refresh ListView 3
''..This is to REFRESH the ListView6, after updating the Details
  Dim uj
  uj = 0
   FrmPaymentAnalysis.ListView6.ListItems.Clear
    If rstReceipt.EOF = False Then
   
   rstReceipt.MoveFirst
   While rstReceipt.EOF = False
       
            If Trim(rstReceipt!SerialNo) = Trim(FrmPaymentAnalysis.txtSerialNo.Text) Then
        
     Set MItem = FrmPaymentAnalysis.ListView6.ListItems.Add(, , Trim(rstReceipt!AccNo))
     MItem.SubItems(1) = Trim(rstReceipt!AccName)
     MItem.SubItems(2) = Format(rstReceipt!amount, "###########.#0")
     uj = Val(uj) + Val(rstReceipt!amount)
        End If
     rstReceipt.MoveNext
     Wend
Me.txtTotList6 = uj

'........................................
End If

         txtSerialNo.Text = ""
         txtDBAccNo.Text = ""
         txtDBPartic.Text = ""
         txtDBAmo.Text = ""
         frmMenu.REdit.caption = "Edit"



If Me.txtTotList3.Text <> Me.txtTotList6.Text Then
MsgBox "Total Debit and Credit Balances not Equal", vbInformation, "Check the Balance"
End If


End Sub


Private Sub Check1_Click()
i = 0
If Me.Check1.Value = 1 Then
  For Each Control In Me
    'If TypeOf Control Is ComboBox Then
      On Error Resume Next
        Control.RightToLeft = True
    'End If
  Next
 Else
   For Each Control In Me
    'If TypeOf Control Is ComboBox Then
    On Error Resume Next
        Control.RightToLeft = False
    'End If
  Next
End If
'If Me.Check1.Value = 1 Then
'       txtDBPartic.RightToLeft = True
'       txtPartic.RightToLeft = True
'Else
'        txtDBPartic.RightToLeft = False
'       txtPartic.RightToLeft = False
'
'End If

End Sub

Private Sub CmbApprovedBy_GotFocus()
Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(CmbApprovedBy.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)

End Sub

Private Sub CmbNotedBy_GotFocus()
Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(CmbNotedBy.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)

End Sub

Private Sub CmbNotedBy_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
CmbApprovedBy.SetFocus
End If

End Sub

Private Sub CmbPrepBy_GotFocus()
Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(CmbPrepBy.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)

End Sub

Private Sub CmbPrepBy_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
CmbNotedBy.SetFocus
End If
End Sub

Private Sub CMDEDIT_Click()
If txtTotList3.Text <> txtTotList6.Text Then
MsgBox "Total Credit and Total Debit Balances not equal", vbInformation, "Balance Conflicted"
Exit Sub
End If


If txtPaymentLevel.Text = "First" Or txtPaymentLevel.Text = "Full" Then
    
    If Format(txtTotList3.Text, "#############.##") <> Format(Text1.Text, "#############.##") Then
    MsgBox "Entered Amount should be Equal to The Invoice Amount", vbInformation, "Records not Saved"
    Exit Sub
    End If
    
    
    If Format(txtTotList6.Text, "#############.##") <> Format(Text1.Text, "#############.##") Then
    MsgBox "It is the - Accrued Liabilities ,Entered Amount should be Equal to The Invoice Amount", vbInformation, "Records not Saved"
    Exit Sub
    End If

Else ' this will be second , third,forth,fifth,final
    If Format(txtTotList3.Text, "#############.##") <> Format(txtAmtReq.Text, "#############.##") Then
    MsgBox "Entered Amount should be Equal to The Amount Requested", vbInformation, "Records not Saved"
    Exit Sub
    End If
    
    
    If Format(txtTotList6.Text, "#############.##") <> Format(txtAmtReq.Text, "#############.##") Then
    MsgBox "Entered Amount should be Equal to The Amount Requested", vbInformation, "Records not Saved"
    Exit Sub
    End If
End If


        X = MsgBox("Are You sure Adding this Records ?", vbYesNo, "SAVE")
        If X = vbNo Then
        Exit Sub
        End If



Call saveme2
End Sub

Private Sub cmdExit1_Click()

frmMenu.sDel.Enabled = True
frmMenu.Rdel.Enabled = True


If cmdExit1.caption = "E&xit" Then  'This is to Exit From form
Unload Me

Else    'This is to Cancel the Job
FrmPaymentAnalysis.CmdNew.Visible = True

frmMenu.shedit.Enabled = True
FrmPayableSetup.cmdedit.caption = Trim("E&dit")
'FrmPayableSetup.ListView3.ListItems.clear
frmMenu.sEdit.caption = "Edit" 'Once Press Cancel this Caption Agian will be "Edit"
frmMenu.Edit.caption = "Edit" 'Once Press Cancel this Caption Agian will be "Edit"

CmdNew.caption = "&New"
cmdExit1.caption = "E&xit"

txtDBAccNo.Enabled = False
txtDBPartic.Enabled = False
txtDBAmo.Enabled = False
txtAccNo.Enabled = False
txtPartic.Enabled = False
txtAmo.Enabled = False

 CmbApprovedBy.Enabled = False
 CmbNotedBy.Enabled = False
 CmbPrepBy.Enabled = False

 CmbApprovedBy.Text = ""
 CmbNotedBy.Text = ""
 CmbPrepBy.Text = ""
txtDBAccNo.Text = ""
txtDBAmo.Text = ""
txtDBPartic = ""

txtAccNo.Text = ""
txtAmo.Text = ""
txtPartic = ""


ListView6.ListItems.Clear
ListView3.ListItems.Clear

For Each Control In Me
    If TypeOf Control Is TextBox Then
        Control.Text = ""
    End If
Next

End If

End Sub

Public Sub saveme2()

Dim CON1 As New ADODB.Connection
Dim rstlistDel As New ADODB.Recordset
Dim rstRecDel As New ADODB.Recordset
conStr = "Provider=MSDASQL;DSN=Finance;UID=sa; PWD=;"
CON1.Open conStr
Dim hhhh1, hhhh2, Ser
Ser = txtSerialNo.Text

hhhh1 = "delete from  xPayment where serialno =" & "'" & Ser & "'" & ""
hhhh2 = "delete  from xreceipt where serialno =" & "'" & Ser & "'" & ""

rstlistDel.Open hhhh1, CON1, adOpenDynamic, adLockOptimistic
rstRecDel.Open hhhh2, CON1, adOpenDynamic, adLockOptimistic

Dim xx, MyCat
xx = 0
              'this is to save the listview(Payments) details to a table
              i = 0
              For i = 1 To Me.ListView3.ListItems.Count
                  Sn = Me.txtSerialNo.Text
                  TDt = FrmPaymentAnalysis.MaskEdBox1.Text
                  Ac = Me.ListView3.ListItems.Item(i)
                  
                        'Here i have to call the proceduer to got the Father Category
                            acctNo = Trim(Ac)
                            Prevcap = Trim(Me.caption)
                            Call DisplayCats(Prevcap, acctNo, catName)
                            'MyCat = "Payable Setup" & "//" & catName
                            MyCat = catName

                  
                  
                  an = Me.ListView3.ListItems.Item(i).SubItems(1)
                  am = Me.ListView3.ListItems.Item(i).SubItems(2)
                  tot = Me.txtTotList3
                     With rstPayment
                           .AddNew
                           !SerialNo = Sn
                           !TRansDate = TDt
                           !AccNo = Ac
                           !AccName = an
                           !amount = am
                           !ConfirmedMark = "uc"
                           !Postmark = "No"
                           !category = MyCat
                           '!Dbamount = Db
                           !total = tot
                           'On Error Resume Next
                           xx = 1
                           .Update
                      End With
               Next






''''''''''''''''rstPaySetup.MoveFirst
''''''''''''''''While rstPaySetup.EOF = False
''''''''''''''''If Sn = rstPaySetup!serialno Then
''''''''''''''''rstPaySetup!totcramt = tot
''''''''''''''''End If
''''''''''''''''rstPaySetup.MoveNext
''''''''''''''''Wend

Dim MyCat2

              'this is to save the listview(Receipt) details to a table
              i = 0
              For i = 1 To FrmPaymentAnalysis.ListView6.ListItems.Count
                  sn2 = FrmPaymentAnalysis.txtSerialNo.Text
                  TDt2 = FrmPaymentAnalysis.MaskEdBox1.Text
                  AC2 = FrmPaymentAnalysis.ListView6.ListItems.Item(i)
                  an2 = FrmPaymentAnalysis.ListView6.ListItems.Item(i).SubItems(1)
                  am2 = FrmPaymentAnalysis.ListView6.ListItems.Item(i).SubItems(2)
                  tot2 = txtTotList6.Text
                  
                        'Here i have to call the proceduer to got the Father Category
                            acctNo = Trim(AC2)
                            Prevcap = Trim(Me.caption)
                            Call DisplayCats(Prevcap, acctNo, catName)
                            'MyCat = "Payable Setup" & "//" & catName
                            MyCat2 = catName
                  
                  
                  
                  
                  
                  
                     With rstReceipt
                           .AddNew
                           !SerialNo = sn2
                           !TRansDate = TDt2
                           !AccNo = AC2
                           !AccName = an2
                           !amount = am2
                           !total = tot2
                           !ConfirmedMark = "uc"
                           !Prepby = cLogUser
                           !Postmark = "No"
                           !category = MyCat2
                           !recordedDate = Date
                           'On Error Resume Next
                           .Update
                           xx = 1
                      End With
               Next
Me.ListView3.ListItems.Clear
Me.ListView6.ListItems.Clear

'This is to Update the Payable setup Total Credit Amount when i came from the  Payablesetup -->Edit
If Me.cmdedit.caption = "&Update" Then
Dim PaySetLast As New ADODB.Recordset

vvvvv = "Select * from Payablesetup setup where serialno = " & "'" & sn2 & "'" & ""
PaySetLast.Open vvvvv, CON1, adOpenDynamic, adLockOptimistic

On Error Resume Next
PaySetLast.MoveFirst
On Error GoTo 0
While PaySetLast.EOF = False
  PaySetLast!TotCrAmt = Trim(tot)
PaySetLast.MoveNext
Wend
PaySetLast.Close


 'Here Refresh  Payable setup ----> ListView 1
 
  Dim PaySetLast2 As New ADODB.Recordset
  
 vvv = "Select * from Payablesetup setup where CancelledMark = '0' And ConfirmedMark = '0' And Paidmark = '0' And Deletemark = '0' and Post = 'No'"
PaySetLast2.Open vvv, CON1, adOpenDynamic, adLockOptimistic

 FrmPayableSetup.ListView1.ListItems.Clear
  
If PaySetLast2.EOF = False Then
PaySetLast2.MoveFirst
End If



While PaySetLast2.EOF = False
'If Trim(PaySetLast2!cancelledmark) = 0 And Trim(PaySetLast2!ConfirmedMark) = 0 And Trim(PaySetLast2!Paidmark) = 0 Then   'This is not Cancelled

Set MItem = FrmPayableSetup.ListView1.ListItems.Add(, , Format(PaySetLast2!SerialNo))
MItem.SubItems(1) = Format(PaySetLast2!Xdate, "dd/mm/yyyy")
MItem.SubItems(2) = Format(PaySetLast2!Requester)
MItem.SubItems(3) = Format(PaySetLast2!DateDue, "dd/mm/yyyy")
MItem.SubItems(4) = Format(PaySetLast2!RefNo)
MItem.SubItems(5) = Format(PaySetLast2!printmark)
MItem.SubItems(6) = Format(PaySetLast2!journaledmark)
MItem.SubItems(7) = Format(PaySetLast2!amtreqested, "#############.#0")




'On Error Resume Next
Totlist1 = Val(Totlist1) + Val(Trim(PaySetLast2!amtreqested)) 'This is for the Total of the List
'On Error GoTo 0
'End If
PaySetLast2.MoveNext
Wend
FrmPayableSetup.txtTotList1.Text = Totlist1




End If

Me.CmdNew.caption = "&New"
If xx = 1 Then
MsgBox "datas Saved Successfully"


End If
'On Error Resume Next
Unload FrmPaymentAnalysis
'On Error GoTo 0
End Sub

Private Sub cmdNew_Click()

If CmdNew.caption = "&New" Then

'This is to Add New Records
CmdNew.caption = "&Save"
cmdExit1.caption = "&Cancel"
txtAccNo.Enabled = True
txtPartic.Enabled = True
txtAmo.Enabled = True
txtDBAccNo.Enabled = True
txtDBPartic.Enabled = True
txtDBAmo.Enabled = True
CmbPrepBy.Enabled = True
CmbNotedBy.Enabled = True
CmbApprovedBy.Enabled = True

If rstPaySetup.EOF Then
On Error Resume Next
End If


Dim strCount
strCount = 0
If rstPaySetup.EOF = False Then
rstPaySetup.MoveFirst
End If

While rstPaySetup.EOF = False
'strCount = rstPaySetup!serialno
    If strCount < Trim(rstPaySetup!SerialNo) Then
  strCount = Trim(rstPaySetup!SerialNo)
    End If
 rstPaySetup.MoveNext
 Wend
 
 
''This is to increase the serial number one by one
'txtSerialNo.Text = Trim(strCount) + 1
'MaskEdBox1.Text = Date
'
' This is to Save the Records
 ElseIf CmdNew.caption = "&Save" Then
 
'----------------
If txtTotList3.Text <> txtTotList6.Text Then
MsgBox "Total Credit and Total Debit Balances not equal", vbInformation, "Balance Conflicted"
Exit Sub
End If
    
If txtPaymentLevel.Text = "First" Or txtPaymentLevel.Text = "Full" Then
    
    If Format(txtTotList3.Text, "#############.##") <> Format(Text1.Text, "#############.##") Then
    MsgBox "Entered Amount should be Equal to The Invoice Amount", vbInformation, "Records not Saved"
    Exit Sub
    End If
    
    
    If Format(txtTotList6.Text, "#############.##") <> Format(Text1.Text, "#############.##") Then
    MsgBox "It is the - Accrued Liabilities ,Entered Amount should be Equal to The Invoice Amount", vbInformation, "Records not Saved"
    Exit Sub
    End If

Else ' this will be second , third,forth,fifth,final
    If Format(txtTotList3.Text, "#############.##") <> Format(txtAmtReq.Text, "#############.##") Then
    MsgBox "Entered Amount should be Equal to The Amount Requested", vbInformation, "Records not Saved"
    Exit Sub
    End If
    
    
    If Format(txtTotList6.Text, "#############.##") <> Format(txtAmtReq.Text, "#############.##") Then
    MsgBox "Entered Amount should be Equal to The Amount Requested", vbInformation, "Records not Saved"
    Exit Sub
    End If
End If

If CmbPrepBy.Text = "" Then
MsgBox "Fill the Combo Prepared by", vbInformation, "Try Again"
Exit Sub
End If

        X = MsgBox("Are You sure Adding this Records ?", vbYesNo, "SAVE")
        If X = vbNo Then
        Exit Sub
        End If


'-------------



'This will Immediately call the Password Form once we select "YES" to save from MSG Box
frmPassword.txtBuffer = "SaveCr"
frmPassword.txtPrepBy.Text = Me.CmbPrepBy.Text 'Useful to varify the password for him
frmPassword.Show 1
On Error Resume Next
frmPassword.txtUserId.SetFocus 'From the Password it will call SaveMe

ElseIf CmdNew.caption = "&Update" Then
frmPassword.txtBuffer = "Update"
frmPassword.txtPrepBy.Text = Me.CmbPrepBy.Text 'Useful to varify the password for him
frmPassword.Show 1
On Error Resume Next

frmPassword.txtUserId.SetFocus 'From the Password it will call SaveMe


End If

End Sub



Public Sub UpdateList3Xpay()

 If rstPayment.EOF = False Then
   rstPayment.MoveFirst
      While rstPayment.EOF = False
       Dim xvar As String
      ' xvar = Trim(rstPayment!accno)
         If Trim(rstPayment!SerialNo) = Trim(FrmPayableSetup.txtSerialNo.Text) And FrmPaymentAnalysis.txtTemp1List3.Text = Trim(rstPayment!AccNo) And FrmPaymentAnalysis.txtTemp2List3.Text = Trim(rstPayment!AccName) Then
        rstPayment!SerialNo = FrmPaymentAnalysis.txtSerialNo.Text
        rstPayment!AccNo = FrmPaymentAnalysis.txtAccNo.Text
        rstPayment!AccName = FrmPaymentAnalysis.txtPartic.Text
        rstPayment!amount = FrmPaymentAnalysis.txtAmo.Text
       ' rstPayment!amount = 0
        rstPayment!total = FrmPaymentAnalysis.txtTotList3.Text
        frmMenu.Edit.caption = "&Edit"
        Myval = "Gotit"
      End If
   rstPayment.MoveNext
   Wend
If Myval = "Gotit" Then
MsgBox "Records for Payment Details Updated Seccussfully"
End If
rstPayment.Close
End If

'Refresh ListView 3
''..This is to REFRESH the ListView3, after updating the Details

'rstPayment.Open "Select * from xPayment", con1, adOpenDynamic, adLockOptimistic
Dim CON1 As ADODB.Connection
Set CON1 = New ADODB.Connection

conStr = "Provider=MSDASQL;DSN=Finance;UID=sa; PWD=;"

CON1.Open conStr


rstpayment2.Open "Select * from xPayment", CON1, adOpenDynamic, adLockOptimistic

'rstPayment.Open
Dim Vj
Vj = 0
   FrmPaymentAnalysis.ListView3.ListItems.Clear
   
If rstpayment2.EOF = False Then
rstpayment2.MoveFirst
End If

   
   While rstpayment2.EOF = False
       
            If Trim(rstpayment2!SerialNo) = Trim(FrmPayableSetup.txtSerialNo.Text) Then
        
     Set MItem = FrmPaymentAnalysis.ListView3.ListItems.Add(, , Trim(rstpayment2!AccNo))
     MItem.SubItems(1) = Trim(rstpayment2!AccName)
     MItem.SubItems(2) = Format(rstpayment2!amount, "###########.#0")
    Vj = Val(Vj) + Val(rstpayment2!amount)
     End If
     rstpayment2.MoveNext
     Wend
 Me.txtTotList3 = Vj
    
If Me.txtTotList3.Text <> Me.txtTotList6.Text Then
MsgBox "Total Debit and Credit Balances not Equal", vbInformation, "Check the Balance"
End If

'........................................
End Sub


Private Sub Command1_Click()
FindAcctNAme = True
FindAcctNames.Show 1

End Sub

Private Sub Command2_Click()
FindAcctNAme = True
FindAcctNames.Show 1
End Sub

Private Sub Command3_Click()
'On Error Resume Next
'FrmPayableSetup.Show 1
End Sub

Private Sub Form_Load()
Dim CON1 As ADODB.Connection
Set CON1 = New ADODB.Connection
Set rstEmp = New ADODB.Recordset
Set rstPaySetup = New ADODB.Recordset
Set rstPayment = New ADODB.Recordset
Set rstReceipt = New ADODB.Recordset
Set rstChart = New ADODB.Recordset
txtDBAccNo.Enabled = False
txtDBPartic.Enabled = False
txtDBAmo.Enabled = False
txtAccNo.Enabled = False
txtPartic.Enabled = False
txtAmo.Enabled = False

conStr = "Provider=MSDASQL;DSN=Finance;UID=sa; PWD=;"


Set xcol = Me.ListView3.ColumnHeaders.Add(, , "Ch of A/c No", 2200)
Set xcol = Me.ListView3.ColumnHeaders.Add(, , "A/c Name", 4500)
Set xcol = Me.ListView3.ColumnHeaders.Add(, , "Amount", 1500)

Set xcol = Me.ListView6.ColumnHeaders.Add(, , "Ch of A/c No", 2200)
Set xcol = Me.ListView6.ColumnHeaders.Add(, , "A/c Name", 4500)
Set xcol = Me.ListView6.ColumnHeaders.Add(, , "Amount", 1500)

Me.ListView6.ColumnHeaders(3).Alignment = lvwColumnRight
Me.ListView3.ColumnHeaders(3).Alignment = lvwColumnRight

CON1.Open conStr

rstEmp.Open "Select * from Newemployee", CON1, adOpenDynamic, adLockOptimistic
rstPaySetup.Open "Select * from PayableSetup", CON1, adOpenDynamic, adLockOptimistic
rstPayment.Open "Select * from xPayment", CON1, adOpenDynamic, adLockOptimistic
rstReceipt.Open "Select * from xReceipt", CON1, adOpenDynamic, adLockOptimistic
'rstChart.Open "Select * from financemaster order by Accountcode", con1, adOpenDynamic, adLockOptimistic
 Dim xClass As New HabitatClass
 Dim xtable As String
 Dim sqltable As Boolean
 'Set conn1 = New ADODB.Connection
 Set rstChart = New ADODB.Recordset
 xtable = "Select * from FinanceMaster order by AccountCode"
 sqltable = True
 xClass.GetTables rstChart, conn1, xtable, constring, sqltable


'If rstChart.EOF = False Then
'rstChart.MoveFirst
'End If

While rstChart.EOF = False
 If Mid(Trim(rstChart!AccountCode), 1, 5) <> "11101" Then
  If Mid(Trim(rstChart!AccountCode), 1, 5) <> "11102" Then
   If Mid(Trim(rstChart!AccountCode), 1, 5) <> "11105" Then
    If Mid(Trim(rstChart!AccountCode), 1, 5) <> "11106" Then
     If Mid(Trim(rstChart!AccountCode), 1, 5) <> "11107" Then
      If rstChart!Active = 1 Then
       txtDBAccNo.AddItem rstChart!AccountCode
       txtDBPartic.AddItem rstChart!accountnameeng & "\" & RTrim(rstChart!accountnamearab)
       
       txtAccNo.AddItem rstChart!AccountCode
       txtPartic.AddItem rstChart!accountnameeng & "\" & RTrim(rstChart!accountnamearab)
      End If
      Else
      txtAccNo.AddItem rstChart!AccountCode
      txtPartic.AddItem rstChart!accountnameeng & "\" & RTrim(rstChart!accountnamearab)
     End If
     Else
     txtAccNo.AddItem rstChart!AccountCode
     txtPartic.AddItem rstChart!accountnameeng & "\" & RTrim(rstChart!accountnamearab)
    End If
    Else
    txtAccNo.AddItem rstChart!AccountCode
    txtPartic.AddItem rstChart!accountnameeng & "\" & RTrim(rstChart!accountnamearab)
   End If
   Else
   txtAccNo.AddItem rstChart!AccountCode
   txtPartic.AddItem rstChart!accountnameeng & "\" & RTrim(rstChart!accountnamearab)
  End If
 Else
  txtAccNo.AddItem rstChart!AccountCode
  txtPartic.AddItem rstChart!accountnameeng & "\" & RTrim(rstChart!accountnamearab)
 End If
 rstChart.MoveNext
 Wend

If rstEmp.EOF = False Then
rstEmp.MoveFirst
End If

While rstEmp.EOF = False
       CmbPrepBy.AddItem rstEmp!Name
       CmbNotedBy.AddItem rstEmp!Name
       CmbApprovedBy.AddItem rstEmp!Name
   rstEmp.MoveNext
Wend
  

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
rstChart.Close
conn1.Close
End Sub

Private Sub ListView3_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  ListView3.SortKey = ColumnHeader.Index - 1
   ListView3.Sorted = True

End Sub

Private Sub ListView3_DblClick()
If Me.ListView3.ListItems.Count <> 0 Then
frmMenu.sEdit.caption = "Update"
FrmPaymentAnalysis.txtAccNo.Text = FrmPaymentAnalysis.ListView3.SelectedItem.Text
FrmPaymentAnalysis.txtPartic.Text = FrmPaymentAnalysis.ListView3.SelectedItem.SubItems(1)
FrmPaymentAnalysis.txtAmo.Text = FrmPaymentAnalysis.ListView3.SelectedItem.SubItems(2)
FrmPaymentAnalysis.Text4.Text = "Sthoothy"
End If

frmMenu.sDel.Enabled = False

End Sub

Private Sub ListView3_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If CmdNew.caption = "&Save" Or cmdedit.caption = "&Update" Then

If Button = 2 Then
  PopupMenu frmMenu.Rate
End If
End If


End Sub


Private Sub ListView6_DblClick()
If Me.ListView6.ListItems.Count <> 0 Then
frmMenu.REdit.caption = "Update"
FrmPaymentAnalysis.txtDBAccNo.Text = FrmPaymentAnalysis.ListView6.SelectedItem.Text
FrmPaymentAnalysis.txtDBPartic.Text = FrmPaymentAnalysis.ListView6.SelectedItem.SubItems(1)
FrmPaymentAnalysis.txtDBAmo.Text = FrmPaymentAnalysis.ListView6.SelectedItem.SubItems(2)
FrmPaymentAnalysis.Text4.Text = "Sukran"
End If
frmMenu.Rdel.Enabled = False


End Sub

Private Sub ListView6_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If CmdNew.caption = "&Save" Or cmdedit.caption = "&Update" Then

If Button = 2 Then
  PopupMenu frmMenu.Rec
End If
End If

End Sub



Private Sub txtAccNo_Click()
'rstChart.MoveFirst

'While rstChart.EOF = False
'If Trim(txtAccNo.Text) = Trim(rstChart!AccountCode) Then
'  txtPartic.Text = rstChart!AccountNameEng '& "\" & RTrim(rstChart!AccountNameArab)
'End If
'rstChart.MoveNext
'Wend
Dim rstAN As New ADODB.Recordset
rstAN.Open "select * from Financemaster where accountCode=" & "'" & Trim(Me.txtAccNo.Text) & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
If rstAN.EOF = False Then
    txtPartic.Text = rstAN!accountnameeng & "\" & RTrim(rstAN!accountnamearab)
  Else
  txtPartic.Text = ""
End If

Dim catName As String
Dim Prevcap As String
Dim acctNo As String
acctNo = Trim(txtAccNo.Text)
Prevcap = Trim(Me.caption)
Call DisplayCats(Prevcap, acctNo, catName)
Me.caption = "Payable Setup" & "//" & catName
DrCat = catName
End Sub

Private Sub txtAccNo_GotFocus()
Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(txtAccNo.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)

End Sub


Private Sub txtAccNo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyControl Then
txtDBAccNo.SetFocus
End If
End Sub

Private Sub txtAccNo_KeyPress(KeyAscii As Integer)
txtAccNo_Click


If KeyAscii = 13 Then
txtPartic.SetFocus
End If
'txtAccNo_Click

End Sub

Private Sub txtAccNo_LostFocus()


'rstChart.MoveFirst
'While rstChart.EOF = False
'If txtAccNo.Text = rstChart!AccountCode Then
'txtPartic.Text = rstChart!AccountNameEng & "\" & RTrim(rstChart!AccountNamearab)
'End If
'rstChart.MoveNext
'Wend



X = ListView6.ListItems.Count
For i = 1 To X
myDBvar = Me.ListView6.ListItems.Item(i)

If Me.txtAccNo <> "" And myDBvar <> "" Then
    If Me.txtAccNo = myDBvar Then
      If Me.txtAccNo <> "242010000000" Then
     xmsg = MsgBox("You have entered same Account number in Debit Side", vbInformation + vbOKOnly, "Message")
     Me.txtAccNo.SetFocus
     Exit Sub
    End If
    End If
End If
Next



u = ListView3.ListItems.Count
For f = 1 To u
mycrvar4 = Me.ListView3.ListItems.Item(f)

If Me.txtAccNo <> "" And mycrvar4 <> "" Then
     If frmMenu.sEdit.caption <> "Update" Then
         If Me.txtAccNo = mycrvar4 Then
     xmsg = MsgBox("You have entered same Account number in the Same Side", vbInformation + vbOKOnly, "Message")
     Me.txtAccNo.SetFocus
     Exit Sub
     End If
    End If
End If
Next

End Sub

Private Sub txtAmo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

If frmMenu.sEdit.caption = "Edit" And FrmPaymentAnalysis.CmdNew.caption = "&Update" And FrmPaymentAnalysis.CmdNew.Visible = True Then
'This will save the additional field to the file insted of saving all the available datas in the listbox
xok = MsgBox("Are you sure you want to save this Debit details", vbOKCancel, "SAVE")
    If xok = vbOK Then
    'Save

     With rstPayment
        .AddNew
            rstPayment!SerialNo = FrmPaymentAnalysis.txtSerialNo.Text
            rstPayment!AccNo = FrmPaymentAnalysis.txtAccNo.Text
            rstPayment!AccName = FrmPaymentAnalysis.txtPartic.Text
            rstPayment!amount = FrmPaymentAnalysis.txtAmo.Text
            rstPayment!total = FrmPaymentAnalysis.txtTotList3.Text
       '   frmMenu.edit.Caption = "&Edit"

     .Update

    MsgBox "Records for Payments Details Saved Seccussfully"
    End With
    End If
'
FrmPaymentAnalysis.txtAccNo = ""
FrmPaymentAnalysis.txtPartic = ""
FrmPaymentAnalysis.txtAmo = ""

'REFRESH ListView 3
   FrmPaymentAnalysis.ListView3.ListItems.Clear
   
If rstPayment.EOF = False Then
rstPayment.MoveFirst
End If

   
   While rstPayment.EOF = False
       
      If Trim(rstPayment!SerialNo) = Trim(FrmPayableSetup.txtSerialNo.Text) Then
        
     Set MItem = FrmPaymentAnalysis.ListView3.ListItems.Add(, , Trim(rstPayment!AccNo))
     MItem.SubItems(1) = Trim(rstPayment!AccName)
     MItem.SubItems(2) = Format(rstPayment!amount, "###########.#0")
        End If
     rstPayment.MoveNext
     Wend


'-0---0-0-0-0-0-0-0
ElseIf frmMenu.sEdit.caption = "Edit" And FrmPaymentAnalysis.CmdNew.caption = "&Save" Then


'This is to send the combo,textbox details to Listview
     Set MItem = Me.ListView3.ListItems.Add(, , Trim(txtAccNo.Text))
     MItem.SubItems(1) = Trim(txtPartic.Text)
     MItem.SubItems(2) = Trim(txtAmo.Text)
TotList3 = Val(txtTotList3) + Val(Trim(txtAmo.Text)) 'This is for the Total of the List
txtAccNo.Text = ""
txtPartic.Text = ""
txtAmo.Text = ""
txtAccNo.SetFocus
'CmbPrepBy.SetFocus
txtTotList3.Text = TotList3

'This is New  when i edit from Payable setup and to add the details to the Listview
ElseIf frmMenu.sEdit.caption = "Edit" And FrmPaymentAnalysis.cmdedit.caption = "&Update" Then

'This is to send the combo,textbox details to Listview
     Set MItem = Me.ListView3.ListItems.Add(, , Trim(txtAccNo.Text))
     MItem.SubItems(1) = Trim(txtPartic.Text)
     MItem.SubItems(2) = Trim(txtAmo.Text)
TotList3 = Val(txtTotList3) + Val(Trim(txtAmo.Text)) 'This is for the Total of the List
txtAccNo.Text = ""
txtPartic.Text = ""
txtAmo.Text = ""
txtAccNo.SetFocus
'CmbPrepBy.SetFocus
txtTotList3.Text = TotList3




ElseIf FrmPaymentAnalysis.Text4.Text = "Sthoothy" And frmMenu.sEdit.caption = "Update" Then
'This is when i did the Item click to Edit and to Update the details
DeleteTermj1 = FrmPaymentAnalysis.ListView3.SelectedItem.Index
Dim amous  As Currency

 amous = FrmPaymentAnalysis.ListView3.ListItems(DeleteTermj1).SubItems(2)
 
FrmPaymentAnalysis.txtTotList3.Text = Val(FrmPaymentAnalysis.txtTotList3.Text) - Val(amous)
 
 FrmPaymentAnalysis.ListView3.ListItems.Remove DeleteTermj1
     Set MItem = Me.ListView3.ListItems.Add(, , Trim(txtAccNo.Text))
     MItem.SubItems(1) = Trim(txtPartic.Text)
     MItem.SubItems(2) = Trim(txtAmo.Text)
     
FrmPaymentAnalysis.txtTotList3.Text = Val(FrmPaymentAnalysis.txtTotList3.Text) + Val(txtAmo.Text)
FrmPaymentAnalysis.txtAccNo.Text = ""
FrmPaymentAnalysis.txtPartic.Text = ""
FrmPaymentAnalysis.txtAmo.Text = ""
FrmPaymentAnalysis.Text4 = ""
frmMenu.sEdit.caption = "Edit"


'ElseIf frmMenu.sEdit.Caption = "Update" And FrmPaymentAnalysis.cmdNew.Caption = "&Update" Then

'frmPassword.txtBuffer.Text = "UpdateList3"
'frmPassword.txtUserId.Text = FrmPaymentAnalysis.CmbPrepBy.Text
'frmPassword.Caption = "Enter Password to Update"
'frmPassword.txtPrepBy = FrmPaymentAnalysis.CmbPrepBy.Text
'------------------

'frmPassword.Show 1
End If
End If
End Sub

Private Sub txtAmo_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Or KeyCode = 17 Or KeyCode = 67 Or KeyCode = 86 Or KeyCode = 8 Or KeyCode = 18 Or KeyCode = 46 Or KeyCode = 27 Then
Exit Sub
End If



If KeyCode <> 86 Then
 If KeyCode <> 17 Then
 If KeyCode <> 13 Then
 If KeyCode <> 110 Then
 If KeyCode < 96 Or KeyCode > 105 Then

 If KeyCode < 48 Or KeyCode > 57 Then
    MsgBox "The Key you entered is not the Digit", vbExclamation, "Try Again"
    Me.txtAmo = ""
    Exit Sub
End If
End If
End If
End If
End If
End If

End Sub

Private Sub txtAmo_LostFocus()
txtAmo.Text = Format(txtAmo.Text, "############.#0")
End Sub

Private Sub txtAmtreq_Change()
txtAmtReq.Text = Format(txtAmtReq.Text, "############.#0")

End Sub

Private Sub txtDBAccNo_Click()
'rstChart.MoveFirst
'While rstChart.EOF = False
'If txtDBAccNo.Text = rstChart!AccountCode Then
'txtDBPartic.Text = rstChart!AccountNameEng & "\" & RTrim(rstChart!AccountNameArab)
'Exit Sub
'End If
'rstChart.MoveNext
'Wend
Dim rstAN As New ADODB.Recordset
rstAN.Open "select * from Financemaster where accountCode=" & "'" & Trim(Me.txtDBAccNo.Text) & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
If rstAN.EOF = False Then
    txtDBPartic.Text = rstAN!accountnameeng & "\" & RTrim(rstAN!accountnamearab)
  Else
  txtDBPartic.Text = ""
End If

acctNo = Trim(txtDBAccNo)
Prevcap = Trim(Me.caption)
Call DisplayCats(Prevcap, acctNo, catName)
Me.caption = "Payable Setup" & "//" & catName
DrCat = catName
End Sub

Private Sub txtDBAccNo_GotFocus()
Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(txtDBAccNo.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)

End Sub

Private Sub txtDBAccNo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyControl Then
CmbPrepBy.SetFocus
End If
End Sub

Private Sub txtDBAccNo_KeyPress(KeyAscii As Integer)

txtDBAccNo_Click

If KeyAscii = 13 Then
txtDBPartic.SetFocus
End If

'txtDBAccNo_Click

End Sub

Private Sub txtDBAccNo_LostFocus()
t = ListView3.ListItems.Count
For i = 1 To t
mycrvar = Me.ListView3.ListItems.Item(i)

If Me.txtDBAccNo <> "" And mycrvar <> "" And frmMenu.REdit.caption <> "Update" Then
    If Me.txtDBAccNo = mycrvar Then
        If Me.txtDBAccNo <> "242010000000" Then
   
     xmsg = MsgBox("You have entered same Account number in Credit Side", vbInformation + vbOKOnly, "Message")
     Me.txtDBAccNo.SetFocus
     Exit Sub
    End If
    End If
End If
Next


y = ListView6.ListItems.Count
For r = 1 To y
mykikkili = Me.ListView6.ListItems.Item(r)

If Me.txtDBAccNo <> "" And mykikkili <> "" And frmMenu.REdit.caption <> "Update" Then
If Me.txtDBAccNo = mykikkili Then
  xmsg = MsgBox("You have entered same Account number in The Same Side", vbInformation + vbOKOnly, "Message")
     Me.txtDBAccNo.SetFocus
     Exit Sub
    End If
End If
Next

End Sub

Private Sub txtDBAmo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

'If frmMenu.REdit.Caption = "Edit" And FrmPaymentAnalysis.cmdNew.Caption = "&Update" And FrmPaymentAnalysis.cmdNew.Visible = True Then
''This will save the additional field to the file insted of saving all the available datas in the listbox
'xok = MsgBox("Are you sure you want to save this Debit details", vbOKCancel, "SAVE")
'    If xok = vbOK Then
'    'Save
'
'     With rstReceipt
'        .AddNew
'            rstReceipt!serialno = FrmPaymentAnalysis.txtSerialNo.Text
'            rstReceipt!AccNo = FrmPaymentAnalysis.txtDBAccNo.Text
'            rstReceipt!AccName = FrmPaymentAnalysis.txtDBPartic.Text
'            rstReceipt!amount = FrmPaymentAnalysis.txtDBAmo.Text
'            rstReceipt!total = FrmPaymentAnalysis.txtTotList6.Text
'       '   frmMenu.edit.Caption = "&Edit"
'
'     .Update
'
'    MsgBox "Records for Receipt Details Saved Seccussfully"
'    End With
'    End If
''
''REFRESH ListView 6
'   FrmPaymentAnalysis.ListView6.ListItems.clear
'   If rstReceipt.EOF = False Then
'   rstReceipt.MoveFirst
'   End If
'
'   While rstReceipt.EOF = False
'
'            If Trim(rstReceipt!serialno) = Trim(FrmPaymentAnalysis.txtSerialNo.Text) Then
'
'     Set mitem = FrmPaymentAnalysis.ListView6.ListItems.Add(, , Trim(rstReceipt!AccNo))
'     mitem.SubItems(1) = Trim(rstReceipt!AccName)
'     mitem.SubItems(2) = Format(rstReceipt!amount, "###########.#0")
'        End If
'     rstReceipt.MoveNext
'     Wend
'
'
'
'
'
If frmMenu.REdit.caption = "Edit" And FrmPaymentAnalysis.CmdNew.caption = "&Save" Then


'This is to send the combo,textbox details to Listview
     Set MItem = Me.ListView6.ListItems.Add(, , Trim(txtDBAccNo.Text))
     MItem.SubItems(1) = Trim(txtDBPartic.Text)
     MItem.SubItems(2) = Trim(txtDBAmo.Text)
TotList6 = Val(txtTotList6) + Val(Trim(txtDBAmo.Text)) 'This is for the Total of the List
txtDBAccNo.Text = ""
txtDBPartic.Text = ""
txtDBAmo.Text = ""
txtDBAccNo.SetFocus
'CmbPrepBy.SetFocus
txtTotList6.Text = TotList6




ElseIf FrmPaymentAnalysis.Text4.Text <> "Sukran" And frmMenu.REdit.caption = "Update" And FrmPaymentAnalysis.CmdNew.caption = "&Update" Then

Call UpdateList6XRecei
'frmPassword.txtBuffer.Text = "UpdateList6"
'frmPassword.txtUserId.Text = FrmPaymentAnalysis.CmbPrepBy.Text
'frmPassword.Caption = "Enter Password to Update"
'frmPassword.txtPrepBy = FrmPaymentAnalysis.CmbPrepBy.Text
'frmPassword.Show 1


'This is New to get the Details to the ListView when it is to be Edited
ElseIf frmMenu.REdit.caption = "Edit" And FrmPaymentAnalysis.cmdedit.caption = "&Update" Then


'This is to send the combo,textbox details to Listview
     Set MItem = Me.ListView6.ListItems.Add(, , Trim(txtDBAccNo.Text))
     MItem.SubItems(1) = Trim(txtDBPartic.Text)
     MItem.SubItems(2) = Trim(txtDBAmo.Text)
TotList6 = Val(txtTotList6) + Val(Trim(txtDBAmo.Text)) 'This is for the Total of the List
txtDBAccNo.Text = ""
txtDBPartic.Text = ""
txtDBAmo.Text = ""
txtDBAccNo.SetFocus
'CmbPrepBy.SetFocus
txtTotList6.Text = TotList6

ElseIf FrmPaymentAnalysis.Text4.Text = "Sukran" And frmMenu.REdit.caption = "Update" Then
'This is when i did the Item click to Edit and to Update the details

DeleteTermj1 = FrmPaymentAnalysis.ListView6.SelectedItem.Index
Dim amou  As Currency

 amou = FrmPaymentAnalysis.ListView6.ListItems(DeleteTermj1).SubItems(2)
 
FrmPaymentAnalysis.txtTotList6.Text = Val(FrmPaymentAnalysis.txtTotList6.Text) - Val(amou)
 
 FrmPaymentAnalysis.ListView6.ListItems.Remove DeleteTermj1
     Set MItem = Me.ListView6.ListItems.Add(, , Trim(txtDBAccNo.Text))
     MItem.SubItems(1) = Trim(txtDBPartic.Text)
     MItem.SubItems(2) = Trim(txtDBAmo.Text)
     
FrmPaymentAnalysis.txtTotList6.Text = Val(FrmPaymentAnalysis.txtTotList6.Text) + Val(txtDBAmo.Text)
FrmPaymentAnalysis.txtDBAccNo.Text = ""
FrmPaymentAnalysis.txtDBPartic.Text = ""
FrmPaymentAnalysis.txtDBAmo.Text = ""
FrmPaymentAnalysis.Text4 = ""
frmMenu.REdit.caption = "Edit"
End If
End If

End Sub

Private Sub txtDBAmo_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Or KeyCode = 17 Or KeyCode = 67 Or KeyCode = 86 Or KeyCode = 8 Or KeyCode = 18 Or KeyCode = 46 Or KeyCode = 27 Then

Exit Sub
End If

 
If KeyCode <> 86 Then
 If KeyCode <> 17 Then
 
 If KeyCode <> 13 Then
 If KeyCode <> 110 Then
 If KeyCode < 96 Or KeyCode > 105 Then

 If KeyCode < 48 Or KeyCode > 57 Then
    MsgBox "The Key you entered is not the Digit", vbExclamation, "Try Again"
    Me.txtDBAmo = ""
    Exit Sub
End If
End If
End If
End If
End If
End If

End Sub

Private Sub txtDBAmo_LostFocus()
txtDBAmo.Text = Format(txtDBAmo.Text, "############.#0")

End Sub

Private Sub txtDBPartic_Click()
X = 0
For X = 1 To Len(txtDBPartic)
    If Mid(Trim(txtDBPartic), X, 1) = "\" Then Exit For
     xname = xname & Mid(txtDBPartic, X, 1)
Next
rstChart.MoveFirst

'While rstChart.EOF = False
'If Trim(xname) = Trim(rstChart!AccountNameEng) Then ' & "\" & RTrim(rstChart!AccountNameArab) Then
'    txtDBAccNo.Text = rstChart!Accountcode
'End If
'rstChart.MoveNext
'Wend

Dim rstAN As New ADODB.Recordset
rstAN.Open "select * from Financemaster where ltrim(accountnameeng)=" & "'" & Trim(xname) & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
If rstAN.EOF = False Then
    txtDBAccNo.Text = rstAN!AccountCode
  Else
  txtDBAccNo.Text = ""
End If

Dim catName As String
Dim Prevcap As String
Dim acctNo As String
acctNo = Trim(xname)
Prevcap = Trim(Me.caption)
Call DisplayCatsName(Prevcap, acctNo, catName)
Me.caption = "Payable Setup " & "//" & catName
DrCat = catName
End Sub

Private Sub txtDBPartic_GotFocus()
Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(txtDBPartic.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)

End Sub

Private Sub txtDBPartic_KeyPress(KeyAscii As Integer)


txtDBPartic_Click

If KeyAscii = 13 Then
txtDBAmo.SetFocus
End If

End Sub

Private Sub txtDBPartic_LostFocus()
'  rstChart.MoveFirst
'
'While rstChart.EOF = False
'If txtDBPartic.Text = rstChart!AccountNameEng Then
'txtDBAccNo.Text = rstChart!AccountCode
'Exit Sub
'End If
'rstChart.MoveNext
'Wend

End Sub

Private Sub txtPartic_Click()

X = 0
For X = 1 To Len(txtPartic)
    If Mid(Trim(txtPartic), X, 1) = "\" Then Exit For
     xname = xname & Mid(txtPartic, X, 1)
Next

'rstChart.MoveFirst
'
'While rstChart.EOF = False
'If Trim(xname) = Trim(rstChart!AccountNameEng) Then '& "\" & RTrim(rstChart!AccountNameArab) Then
' txtAccNo.Text = rstChart!AccountCode
'End If
'rstChart.MoveNext
'Wend

Dim rstAN As New ADODB.Recordset
rstAN.Open "select * from Financemaster where ltrim(accountnameeng)=" & "'" & Trim(xname) & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
If rstAN.EOF = False Then
    txtAccNo.Text = rstAN!AccountCode
  Else
  txtAccNo.Text = ""
End If

Dim catName As String
Dim Prevcap As String
Dim acctNo As String
acctNo = Trim(xname)
Prevcap = Trim(Me.caption)
Call DisplayCatsName(Prevcap, acctNo, catName)
Me.caption = "Payable Setup " & "//" & catName
DrCat = catName
End Sub

Private Sub txtPartic_GotFocus()
Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(txtPartic.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)

End Sub

Private Sub txtPartic_KeyPress(KeyAscii As Integer)

txtPartic_Click

If KeyAscii = 13 Then
txtAmo.SetFocus
End If

End Sub

Private Sub txtPartic_LostFocus()
'txtAccNo_LostFocus
X = ListView6.ListItems.Count
For i = 1 To X
myDBvar = Me.ListView6.ListItems.Item(i)

If Me.txtAccNo <> "" And myDBvar <> "" Then
    If Me.txtAccNo = myDBvar Then
     xmsg = MsgBox("You have entered same Account number in Debit Side", vbInformation + vbOKOnly, "Message")
     Me.txtAccNo.SetFocus
     Exit Sub
    End If
End If
Next



u = ListView3.ListItems.Count
For f = 1 To u
mycrvar4 = Me.ListView3.ListItems.Item(f)

If Me.txtAccNo <> "" And mycrvar4 <> "" Then
    If Me.txtAccNo = mycrvar4 Then
     xmsg = MsgBox("You have entered same Account number in the Same Side", vbInformation + vbOKOnly, "Message")
     Me.txtAccNo.SetFocus
     Exit Sub
    End If
End If
Next

End Sub

Private Sub txtTotList3_Change()
txtTotList3.Text = Format(txtTotList3.Text, "############.#0")

End Sub

Private Sub txtTotList6_Change()
txtTotList6.Text = Format(txtTotList6.Text, "############.#0")
End Sub


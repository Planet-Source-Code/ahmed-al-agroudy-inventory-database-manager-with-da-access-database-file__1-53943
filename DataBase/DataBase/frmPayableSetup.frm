VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmPayableSetup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payable Setup"
   ClientHeight    =   7665
   ClientLeft      =   165
   ClientTop       =   3960
   ClientWidth     =   11970
   Icon            =   "frmPayableSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1.42873e5
   ScaleMode       =   0  'User
   ScaleWidth      =   2.78545e6
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   4680
      TabIndex        =   144
      Text            =   "Text1"
      Top             =   8040
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   0
      TabIndex        =   134
      Top             =   6840
      Width           =   11895
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
         Height          =   375
         Left            =   10800
         TabIndex        =   136
         ToolTipText     =   "Add  New Entry"
         Top             =   240
         Width           =   1000
      End
      Begin VB.CommandButton Command6 
         Caption         =   "N&ext >>"
         Height          =   350
         Left            =   1200
         TabIndex        =   141
         ToolTipText     =   "Go to next Tap"
         Top             =   240
         Width           =   1000
      End
      Begin VB.CommandButton Command5 
         Caption         =   "<<&Back"
         Height          =   350
         Left            =   120
         TabIndex        =   140
         ToolTipText     =   "Go Back to Previos Tap "
         Top             =   240
         Width           =   1000
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
         Left            =   2280
         TabIndex        =   139
         ToolTipText     =   "Close Window"
         Top             =   240
         Width           =   1000
      End
      Begin VB.CommandButton cmdPrintPayReq 
         Caption         =   "Print Payment &Request  ØÈÇÚÉ ØáÈ ÇáãÏÝæÚÇÊ "
         Height          =   375
         Left            =   8400
         TabIndex        =   138
         Top             =   240
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10680
         TabIndex        =   137
         ToolTipText     =   "Print Payable Setup"
         Top             =   240
         Visible         =   0   'False
         Width           =   1000
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
         Left            =   5400
         TabIndex        =   135
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label68 
         Alignment       =   1  'Right Justify
         Caption         =   "ÇÚÏ ÈæÇÓØÉ "
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
         Left            =   7440
         TabIndex        =   143
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Prepared By"
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
         Left            =   4320
         TabIndex        =   142
         Top             =   285
         Width           =   900
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   840
      Top             =   8160
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   12091
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Payable Entry  ÇÏÎÇá ÇáãÏÝæÚÇÊ  "
      TabPicture(0)   =   "frmPayableSetup.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label33"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label25"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label22"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label19"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label8"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label4"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label11"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label13"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label6"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label28"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label29"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label14"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label1"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label43"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label46"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label51"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label54"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label55"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label57"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label58"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label59"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label60"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label62"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label63"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Label65"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Label69"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Label70"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Label56"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Label47"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Label17"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Label48"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Label36"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Label49"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Label37"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Label71"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Label73"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Label74"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Label75"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Label38"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "Label72"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "Label52"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "Label53"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "Label18"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "Label40"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "Label64"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "Label39"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "Label7"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "Label12"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "Label61"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "Label3"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "Label9"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "Label16"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "mskDateDue"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "MskDate"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "txtInvAmt"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "txtBranch"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "Frame1"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "txtserialNo"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "txtOutBal"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "txtRefNo"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "List1"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "cmbProfCenter"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).Control(64)=   "cmbCostCenter"
      Tab(0).Control(64).Enabled=   0   'False
      Tab(0).Control(65)=   "txtDocuNo"
      Tab(0).Control(65).Enabled=   0   'False
      Tab(0).Control(66)=   "Command2"
      Tab(0).Control(66).Enabled=   0   'False
      Tab(0).Control(67)=   "CmbNotedBy"
      Tab(0).Control(67).Enabled=   0   'False
      Tab(0).Control(68)=   "CmbApprovedBy"
      Tab(0).Control(68).Enabled=   0   'False
      Tab(0).Control(69)=   "Command3"
      Tab(0).Control(69).Enabled=   0   'False
      Tab(0).Control(70)=   "Frame2"
      Tab(0).Control(70).Enabled=   0   'False
      Tab(0).Control(71)=   "cmbPaymentFor"
      Tab(0).Control(71).Enabled=   0   'False
      Tab(0).Control(72)=   "CmbSource"
      Tab(0).Control(72).Enabled=   0   'False
      Tab(0).Control(73)=   "txtAmtPaid"
      Tab(0).Control(73).Enabled=   0   'False
      Tab(0).Control(74)=   "txtCreditNote"
      Tab(0).Control(74).Enabled=   0   'False
      Tab(0).Control(75)=   "txtDebitNote"
      Tab(0).Control(75).Enabled=   0   'False
      Tab(0).Control(76)=   "cmbPaymode"
      Tab(0).Control(76).Enabled=   0   'False
      Tab(0).Control(77)=   "cmbCurrency"
      Tab(0).Control(77).Enabled=   0   'False
      Tab(0).Control(78)=   "txtEnter1"
      Tab(0).Control(78).Enabled=   0   'False
      Tab(0).Control(79)=   "TxtEnter2"
      Tab(0).Control(79).Enabled=   0   'False
      Tab(0).Control(80)=   "txtNoOfAtt"
      Tab(0).Control(80).Enabled=   0   'False
      Tab(0).Control(81)=   "cmbPayee2"
      Tab(0).Control(81).Enabled=   0   'False
      Tab(0).Control(82)=   "Check1"
      Tab(0).Control(82).Enabled=   0   'False
      Tab(0).Control(83)=   "txtTaxCredit"
      Tab(0).Control(83).Enabled=   0   'False
      Tab(0).Control(84)=   "txtAmtReq"
      Tab(0).Control(84).Enabled=   0   'False
      Tab(0).Control(85)=   "txtAmountDue"
      Tab(0).Control(85).Enabled=   0   'False
      Tab(0).Control(86)=   "Check2"
      Tab(0).Control(86).Enabled=   0   'False
      Tab(0).Control(87)=   "cmbPayment"
      Tab(0).Control(87).Enabled=   0   'False
      Tab(0).Control(88)=   "txtFCamount"
      Tab(0).Control(88).Enabled=   0   'False
      Tab(0).Control(89)=   "cmbPaymentLevel"
      Tab(0).Control(89).Enabled=   0   'False
      Tab(0).ControlCount=   90
      TabCaption(1)   =   "Voucher  List"
      TabPicture(1)   =   "frmPayableSetup.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label30"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label24"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label26"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label23"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "ListView7"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "ListView1"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "txtTotList1"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "txtTotList7"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Command1"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "CmdEdit"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "CmdPost"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "txtToday"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "txtSearch1"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "cmdSearch"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).ControlCount=   14
      TabCaption(2)   =   "Cancelled ÇáÛÇÁ "
      TabPicture(2)   =   "frmPayableSetup.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtTotCancelled"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "ListView4"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cmdSearch2"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "txtSearch2"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "txtTotCanceleldBal"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Paid Voucher ÝÇÊæÑÉ ÇáãÏÝæÚÇÊ "
      TabPicture(3)   =   "frmPayableSetup.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label31"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "ListView5"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "cmdSearch3"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "txtSearch3"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "txtTotVoch"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).ControlCount=   5
      TabCaption(4)   =   "Journal  íæãíÉ "
      TabPicture(4)   =   "frmPayableSetup.frx":04B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label35"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "ListView8"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "txtJdb"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "txtJCr"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "Command4"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "txtSearch8"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "txtForFrmRef"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).ControlCount=   7
      Begin VB.TextBox txtForFrmRef 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -69720
         TabIndex        =   152
         Top             =   4200
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.ComboBox cmbPaymentLevel 
         Height          =   315
         Left            =   5520
         TabIndex        =   150
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox txtFCamount 
         Height          =   315
         Left            =   1080
         TabIndex        =   148
         Top             =   2280
         Width           =   1935
      End
      Begin VB.ComboBox cmbPayment 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1080
         TabIndex        =   145
         Top             =   1560
         Width           =   1935
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Check2"
         Height          =   255
         Left            =   3480
         TabIndex        =   133
         Top             =   5880
         Width           =   255
      End
      Begin VB.TextBox txtAmountDue 
         Alignment       =   1  'Right Justify
         Height          =   320
         Left            =   9480
         TabIndex        =   118
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox txtAmtReq 
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
         Height          =   320
         Left            =   9480
         TabIndex        =   117
         Top             =   3120
         Width           =   1095
      End
      Begin VB.TextBox txtTaxCredit 
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
         Height          =   320
         Left            =   9480
         TabIndex        =   114
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Height          =   375
         Left            =   7440
         TabIndex        =   113
         Top             =   3360
         Width           =   255
      End
      Begin VB.ComboBox cmbPayee2 
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmPayableSetup.frx":04CE
         Left            =   1080
         List            =   "frmPayableSetup.frx":04D0
         OLEDropMode     =   1  'Manual
         RightToLeft     =   -1  'True
         Sorted          =   -1  'True
         TabIndex        =   109
         Top             =   3360
         Width           =   6375
      End
      Begin VB.ListBox txtNoOfAtt 
         Height          =   645
         Left            =   5520
         TabIndex        =   107
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox TxtEnter2 
         Height          =   320
         Left            =   5520
         TabIndex        =   106
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtEnter1 
         Height          =   320
         Left            =   1080
         TabIndex        =   105
         Top             =   4440
         Width           =   2655
      End
      Begin VB.ComboBox cmbCurrency 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1080
         TabIndex        =   102
         Top             =   1920
         Width           =   2655
      End
      Begin VB.ComboBox cmbPaymode 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1080
         TabIndex        =   98
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox txtDebitNote 
         Alignment       =   1  'Right Justify
         Height          =   320
         Left            =   9480
         TabIndex        =   95
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtCreditNote 
         Alignment       =   1  'Right Justify
         Height          =   320
         Left            =   9480
         TabIndex        =   92
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txtAmtPaid 
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
         Height          =   320
         Left            =   9480
         TabIndex        =   89
         Top             =   960
         Width           =   1095
      End
      Begin VB.ComboBox CmbSource 
         Height          =   315
         Left            =   1080
         TabIndex        =   69
         Top             =   2640
         Width           =   2655
      End
      Begin VB.ComboBox cmbPaymentFor 
         Height          =   315
         Left            =   1080
         TabIndex        =   68
         Top             =   3000
         Width           =   2655
      End
      Begin VB.Frame Frame2 
         Caption         =   "Choice"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   5520
         TabIndex        =   65
         Top             =   2400
         Width           =   1935
         Begin VB.OptionButton Option1 
            Caption         =   "Purchase"
            Height          =   255
            Left            =   120
            TabIndex        =   67
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            Caption         =   "&Others"
            Height          =   255
            Left            =   120
            TabIndex        =   66
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label67 
            Alignment       =   1  'Right Justify
            Caption         =   "ÇÎÑ"
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
            Left            =   1200
            TabIndex        =   85
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label66 
            Alignment       =   1  'Right Justify
            Caption         =   "ãÔÊÑíÇÊ "
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
            Left            =   1200
            TabIndex        =   84
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Go &Purchase Setup ÊÍãíá ÇáãÔÊÑíÇÊ "
         Enabled         =   0   'False
         Height          =   375
         Left            =   8400
         TabIndex        =   64
         Top             =   6360
         Width           =   3375
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
         Left            =   5400
         TabIndex        =   61
         Top             =   6360
         Width           =   1935
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
         Left            =   1080
         TabIndex        =   60
         Top             =   6360
         Width           =   2055
      End
      Begin VB.TextBox txtSearch8 
         Height          =   325
         Left            =   -73800
         TabIndex        =   59
         Top             =   6400
         Width           =   1455
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Search ÈÍË"
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
         Left            =   -74880
         TabIndex        =   58
         ToolTipText     =   "Enter Journal No to Search"
         Top             =   6400
         Width           =   1005
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
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
         Left            =   -74880
         TabIndex        =   57
         Top             =   6400
         Width           =   1000
      End
      Begin VB.TextBox txtSearch1 
         Height          =   325
         Left            =   -73800
         TabIndex        =   56
         Top             =   6400
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Go Payment Analysis  ÊÍáíá ÇáãÏÝæÚÇÊ "
         Height          =   375
         Left            =   8400
         TabIndex        =   55
         Top             =   5880
         Width           =   3375
      End
      Begin VB.TextBox txtToday 
         Height          =   350
         Left            =   -64680
         TabIndex        =   54
         Top             =   360
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtDocuNo 
         Height          =   320
         Left            =   1080
         TabIndex        =   32
         Top             =   3720
         Width           =   1935
      End
      Begin VB.ComboBox cmbCostCenter 
         Height          =   315
         Left            =   1080
         TabIndex        =   31
         Top             =   5820
         Width           =   2415
      End
      Begin VB.ComboBox cmbProfCenter 
         Height          =   315
         Left            =   1080
         TabIndex        =   30
         Top             =   5490
         Width           =   2655
      End
      Begin VB.ListBox List1 
         DataField       =   "Branch"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         Height          =   645
         Left            =   1080
         TabIndex        =   29
         Top             =   4800
         Width           =   2655
      End
      Begin VB.TextBox txtRefNo 
         Height          =   320
         Left            =   5520
         TabIndex        =   28
         Top             =   1680
         Width           =   1935
      End
      Begin VB.CommandButton CmdPost 
         Caption         =   "&Confirm"
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
         Left            =   -69960
         TabIndex        =   27
         ToolTipText     =   "Confirm transaction"
         Top             =   6400
         Width           =   1000
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "E&dit"
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
         Left            =   -71040
         TabIndex        =   26
         ToolTipText     =   "Select List item to Edit"
         Top             =   6400
         Width           =   1000
      End
      Begin VB.TextBox txtOutBal 
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
         Height          =   320
         Left            =   9480
         TabIndex        =   25
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox txtserialNo 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   320
         Left            =   1080
         TabIndex        =   24
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox txtTotCanceleldBal 
         Alignment       =   1  'Right Justify
         Height          =   325
         Left            =   -64680
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   6400
         Width           =   1455
      End
      Begin VB.TextBox txtTotVoch 
         Alignment       =   1  'Right Justify
         Height          =   325
         Left            =   -64680
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   6400
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
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
         Left            =   -72120
         TabIndex        =   19
         ToolTipText     =   "Add New Entry"
         Top             =   6400
         Width           =   1000
      End
      Begin VB.TextBox txtSearch2 
         Height          =   325
         Left            =   -73800
         TabIndex        =   18
         Top             =   6400
         Width           =   1455
      End
      Begin VB.CommandButton cmdSearch2 
         Caption         =   "&Search"
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
         Left            =   -74880
         TabIndex        =   17
         ToolTipText     =   "Search required Journal No"
         Top             =   6400
         Width           =   1000
      End
      Begin VB.TextBox txtSearch3 
         Height          =   325
         Left            =   -73800
         TabIndex        =   16
         Top             =   6400
         Width           =   1455
      End
      Begin VB.CommandButton cmdSearch3 
         Caption         =   "&Search"
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
         Left            =   -74880
         TabIndex        =   15
         ToolTipText     =   "Enter Journal No to Search"
         Top             =   6400
         Width           =   1000
      End
      Begin VB.TextBox txtTotList7 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   -64680
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   6360
         Width           =   1455
      End
      Begin VB.TextBox txtTotList1 
         Alignment       =   1  'Right Justify
         Height          =   325
         Left            =   -64680
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Frame Frame1 
         Caption         =   "Terms"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   5520
         TabIndex        =   6
         Top             =   4080
         Width           =   6255
         Begin VB.TextBox TxtKallapundai 
            Height          =   285
            Left            =   3720
            TabIndex        =   112
            Text            =   "Text1"
            Top             =   1080
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.ComboBox txtTMDes 
            Height          =   320
            Left            =   960
            TabIndex        =   108
            Top             =   240
            Width           =   1890
         End
         Begin VB.ComboBox txtTMmode 
            Height          =   315
            ItemData        =   "frmPayableSetup.frx":04D2
            Left            =   3720
            List            =   "frmPayableSetup.frx":04D4
            TabIndex        =   100
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtTmp2 
            Height          =   285
            Left            =   2040
            TabIndex        =   11
            Text            =   "Text2"
            Top             =   720
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox txtTmp1 
            Height          =   285
            Left            =   1080
            TabIndex        =   10
            Text            =   "Text2"
            Top             =   720
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox txtTmRate 
            Height          =   320
            Left            =   135
            TabIndex        =   9
            Top             =   240
            Width           =   880
         End
         Begin VB.TextBox txtTMDays 
            Height          =   320
            Left            =   2880
            TabIndex        =   8
            Top             =   240
            Width           =   870
         End
         Begin VB.TextBox txtTMlevel 
            Height          =   320
            Left            =   4680
            TabIndex        =   7
            Top             =   240
            Width           =   855
         End
         Begin MSComctlLib.ListView ListView2 
            Height          =   975
            Left            =   120
            TabIndex        =   12
            Top             =   600
            Width           =   5415
            _ExtentX        =   9551
            _ExtentY        =   1720
            View            =   3
            LabelWrap       =   0   'False
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
      End
      Begin VB.ComboBox txtBranch 
         Height          =   315
         Left            =   1080
         TabIndex        =   5
         Top             =   4080
         Width           =   2655
      End
      Begin VB.TextBox txtInvAmt 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   320
         Left            =   9480
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtJCr 
         Alignment       =   1  'Right Justify
         Height          =   325
         Left            =   -64680
         Locked          =   -1  'True
         TabIndex        =   2
         ToolTipText     =   "Total Credit Balance"
         Top             =   6400
         Width           =   1455
      End
      Begin VB.TextBox txtJdb 
         Alignment       =   1  'Right Justify
         Height          =   325
         Left            =   -66240
         Locked          =   -1  'True
         TabIndex        =   1
         ToolTipText     =   "total Debit Balance"
         Top             =   6400
         Width           =   1455
      End
      Begin MSComctlLib.ListView ListView8 
         Height          =   6015
         Left            =   -74880
         TabIndex        =   3
         Top             =   360
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   10610
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
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
      Begin MSComctlLib.ListView ListView5 
         Height          =   6015
         Left            =   -74880
         TabIndex        =   22
         Top             =   360
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   10610
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
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
      Begin MSComctlLib.ListView ListView1 
         Height          =   2535
         Left            =   -74880
         TabIndex        =   23
         Top             =   720
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   4471
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
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
      Begin MSComctlLib.ListView ListView4 
         Height          =   6015
         Left            =   -74880
         TabIndex        =   33
         Top             =   360
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   10610
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
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
      Begin MSMask.MaskEdBox MskDate 
         Height          =   315
         Left            =   1080
         TabIndex        =   34
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
      Begin MSComctlLib.ListView ListView7 
         Height          =   2655
         Left            =   -74880
         TabIndex        =   35
         Top             =   3720
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   4683
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
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
      Begin MSMask.MaskEdBox mskDateDue 
         Height          =   315
         Left            =   9480
         TabIndex        =   119
         Top             =   3480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label16 
         Caption         =   "Pmt Level"
         Height          =   255
         Left            =   4680
         TabIndex        =   149
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "F.C Amount"
         Height          =   255
         Left            =   120
         TabIndex        =   147
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Payment"
         Height          =   255
         Left            =   120
         TabIndex        =   146
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label61 
         AutoSize        =   -1  'True
         Caption         =   "Invoice Details     ÈíÇäÇÊ ÇáÝæÇÊíÑ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   7800
         MousePointer    =   99  'Custom
         TabIndex        =   126
         Top             =   3840
         Width           =   2760
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Amount Due"
         Height          =   195
         Left            =   8280
         TabIndex        =   125
         Top             =   2760
         Width           =   885
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Date Due"
         Height          =   195
         Left            =   8280
         TabIndex        =   124
         Top             =   3480
         Width           =   690
      End
      Begin VB.Label Label39 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ãÌãæÚ ÇáÇÓÊÍÞÇÞ"
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
         Left            =   10710
         TabIndex        =   123
         Top             =   2760
         Width           =   945
      End
      Begin VB.Label Label64 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ÊÇÑíÎ ÇÓÊÍÞÞ ÇáÏÝÚ "
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
         Left            =   10590
         TabIndex        =   122
         Top             =   3360
         Width           =   1065
      End
      Begin VB.Label Label40 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   10935
         TabIndex        =   121
         Top             =   3120
         Width           =   720
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Amt Requested"
         Height          =   195
         Left            =   8280
         TabIndex        =   120
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label Label53 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ÖÑÇÆÈ Ã. Ê. Õ."
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
         Left            =   10605
         TabIndex        =   116
         Top             =   2400
         Width           =   1050
      End
      Begin VB.Label Label52 
         AutoSize        =   -1  'True
         Caption         =   "Tax Credit"
         Height          =   195
         Left            =   8280
         TabIndex        =   115
         Top             =   2400
         Width           =   720
      End
      Begin VB.Label Label72 
         Alignment       =   1  'Right Justify
         Caption         =   "ãÏÝæÚ"
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
         Left            =   7560
         TabIndex        =   111
         Top             =   3480
         Width           =   495
      End
      Begin VB.Label Label38 
         Caption         =   "Payee"
         Height          =   255
         Left            =   120
         TabIndex        =   110
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label Label75 
         Alignment       =   1  'Right Justify
         Caption         =   "äæÚ ÇáÚãáÉ"
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
         Left            =   3600
         TabIndex        =   104
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label74 
         Caption         =   "Currency"
         Height          =   255
         Left            =   120
         TabIndex        =   103
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label73 
         Alignment       =   1  'Right Justify
         Caption         =   "ØÑíÞÉ ÇáÏÝÚ"
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
         Left            =   3600
         TabIndex        =   101
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label71 
         Caption         =   "Pay Mode"
         Height          =   255
         Left            =   120
         TabIndex        =   99
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         Caption         =   "Debit Note"
         Height          =   195
         Left            =   8280
         TabIndex        =   97
         Top             =   1680
         Width           =   765
      End
      Begin VB.Label Label49 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ãáÇÍÙÇÊ ÇáãÏíäæä "
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
         Left            =   10620
         TabIndex        =   96
         Top             =   1680
         Width           =   1035
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         Caption         =   "Credit Note"
         Height          =   195
         Left            =   8280
         TabIndex        =   94
         Top             =   1320
         Width           =   795
      End
      Begin VB.Label Label48 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ãáÇÍÙÇÊ ÇáÏáÆäæä "
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
         Left            =   10650
         TabIndex        =   93
         Top             =   1320
         Width           =   1005
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Paid  Before"
         Height          =   195
         Left            =   8280
         TabIndex        =   91
         Top             =   960
         Width           =   870
      End
      Begin VB.Label Label47 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ÇáãÏÝæÚ ÓÇÈÞÇð"
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
         Left            =   10920
         TabIndex        =   90
         Top             =   960
         Width           =   705
      End
      Begin VB.Label Label56 
         Caption         =   "ãÏÝæÚ Çáí "
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
         Left            =   3720
         TabIndex        =   88
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label70 
         Alignment       =   1  'Right Justify
         Caption         =   "ÊæÞíÚ ÈæÇÓØÉ "
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
         Left            =   7320
         TabIndex        =   87
         Top             =   6360
         Width           =   855
      End
      Begin VB.Label Label69 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ãÑÌÚ ÈæÇÓØÉ "
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
         Left            =   3240
         TabIndex        =   86
         Top             =   6360
         Width           =   780
      End
      Begin VB.Label Label65 
         Alignment       =   1  'Right Justify
         Caption         =   "ãÑÝÞÇÊ "
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
         Left            =   7560
         TabIndex        =   83
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label63 
         Caption         =   "ãÑßÒ ÇáÊßáÝÉ "
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
         Left            =   3960
         TabIndex        =   82
         Top             =   5880
         Width           =   615
      End
      Begin VB.Label Label62 
         Alignment       =   1  'Right Justify
         Caption         =   "ãÑßÒ ÇáÑÈÍ "
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
         Left            =   3960
         TabIndex        =   81
         Top             =   5520
         Width           =   615
      End
      Begin VB.Label Label60 
         Alignment       =   1  'Right Justify
         Caption         =   "ÊÝÓíÑ"
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
         TabIndex        =   80
         Top             =   4440
         Width           =   975
      End
      Begin VB.Label Label59 
         Alignment       =   1  'Right Justify
         Caption         =   "ÝÑÚ"
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
         TabIndex        =   79
         Top             =   4080
         Width           =   975
      End
      Begin VB.Label Label58 
         Alignment       =   1  'Right Justify
         Caption         =   "ÑÞã ÇáæËíÞÉ "
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
         Left            =   3600
         TabIndex        =   78
         Top             =   3720
         Width           =   855
      End
      Begin VB.Label Label57 
         Alignment       =   1  'Right Justify
         Caption         =   "ÑÞã ÇáãÓáÓá "
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
         Left            =   3480
         TabIndex        =   77
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label55 
         Alignment       =   1  'Right Justify
         Caption         =   "ãÏÝæÚ Úáí "
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
         TabIndex        =   76
         Top             =   3000
         Width           =   975
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
         TabIndex        =   75
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label51 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ãÑÌÚ Èå "
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
         Left            =   7560
         TabIndex        =   74
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label46 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   10920
         TabIndex        =   73
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label43 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ãÌãæÚ "
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
         Left            =   11235
         TabIndex        =   72
         Top             =   600
         Width           =   420
      End
      Begin VB.Label Label1 
         Caption         =   "Payment To"
         Height          =   255
         Left            =   120
         TabIndex        =   71
         Top             =   2685
         Width           =   975
      End
      Begin VB.Label Label14 
         Caption         =   "Payment For"
         Height          =   255
         Left            =   120
         TabIndex        =   70
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "Approved By"
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
         Left            =   4320
         TabIndex        =   63
         Top             =   6360
         Width           =   960
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "Noted By"
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
         TabIndex        =   62
         Top             =   6360
         Width           =   660
      End
      Begin VB.Label Label2 
         Caption         =   "Doc. No"
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   3795
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Cost Center"
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   5880
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "Profit Center"
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   5490
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "Explanation"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   4440
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Branch"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   4110
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Date"
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Requester"
         Height          =   255
         Left            =   4680
         TabIndex        =   47
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Outs-ing Bal"
         Height          =   195
         Left            =   8280
         TabIndex        =   46
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label22 
         Caption         =   " Attachmet"
         Height          =   255
         Left            =   4680
         TabIndex        =   45
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label25 
         Caption         =   "Serial No"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   480
         Width           =   855
      End
      Begin VB.Label txtTotCancelled 
         AutoSize        =   -1  'True
         Caption         =   "Total Amount (Cancelled)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -68760
         TabIndex        =   43
         Top             =   6480
         Width           =   2085
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "Total Paid Vouchers  ÅÌãÇáì ÇáãÏÝæÚÇÊ "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -68760
         TabIndex        =   42
         Top             =   6480
         Width           =   2775
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Total Confirmed Balance"
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
         Left            =   -68760
         TabIndex        =   41
         Top             =   6480
         Width           =   2100
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Total Unconfirmed Balance  "
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
         Left            =   -69120
         TabIndex        =   40
         Top             =   3360
         Width           =   2445
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Confirmed List    ãÄßÏ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -74760
         TabIndex        =   39
         Top             =   3360
         Width           =   1695
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "Un Confirmed List     ÇáÛíÑ ãÄßÏ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -74760
         TabIndex        =   38
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   " Amount"
         Height          =   195
         Left            =   8280
         TabIndex        =   37
         Top             =   600
         Width           =   585
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   " Balance ÑÕíÏ "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -68760
         TabIndex        =   36
         Top             =   6480
         Width           =   1065
      End
   End
   Begin MSMask.MaskEdBox mskInvDate 
      Height          =   315
      Left            =   960
      TabIndex        =   127
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtInvNo 
      Enabled         =   0   'False
      Height          =   320
      Left            =   960
      TabIndex        =   132
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtIdentificationForPaymentAnalisis 
      Height          =   285
      Left            =   2640
      TabIndex        =   151
      Text            =   "Text2"
      Top             =   4320
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Label Label50 
      Alignment       =   1  'Right Justify
      Caption         =   "ÝÇÊæÑÉ ÑÞã "
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
      Left            =   2280
      TabIndex        =   131
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label34 
      Caption         =   "Invoice No"
      Height          =   255
      Left            =   0
      TabIndex        =   130
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "ÊÇÑíÎ ÇáÝÇÊæÑÉ "
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
      Left            =   2160
      TabIndex        =   129
      Top             =   360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lablee 
      Caption         =   " Inv Date"
      Height          =   255
      Left            =   0
      TabIndex        =   128
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "FrmPayableSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim xItem As ListItem
'Dim MyClass As New HabitatClass
Dim sqltable As Boolean
Dim CON1 As ADODB.Connection
Dim rstSource As ADODB.Recordset
Dim rstEmp As ADODB.Recordset
Dim rstVen As ADODB.Recordset
Dim rstrate As ADODB.Recordset
Dim rstPAyFor As ADODB.Recordset
Dim rstCosPro As ADODB.Recordset
Dim strVen As String
Dim serial As String
Dim rstPaySetup As ADODB.Recordset
Dim rstPayment As ADODB.Recordset
Dim rstReceipt As ADODB.Recordset
Dim rstterm As ADODB.Recordset
Dim rstChart As ADODB.Recordset
Dim rstJournal As ADODB.Recordset
Dim TotList3
Dim RstPaid As ADODB.Recordset
Dim TotList6
Dim Totlist1
Dim Totlist4
Public MskDa
Public SNom
Public DeleteTerm As Integer
Public ItsForRefrenceForm




Private Sub Check1_Click()
If Me.Check1.Value = 1 Then
        cmbPayee2.RightToLeft = True
Else
        cmbPayee2.RightToLeft = False
End If

End Sub

Private Sub Check2_Click()
If Me.Check2.Value = 1 Then
        cmbCostCenter.RightToLeft = True
Else
        cmbCostCenter.RightToLeft = False
End If

End Sub

Private Sub CmbApprovedBy_click()
'txtApprBy.Text = Trim(CmbApprovedBy.Text)
End Sub

Private Sub CmbApprovedBy_GotFocus()
Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(CmbApprovedBy.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)

End Sub


Private Sub cmbCostCenter_GotFocus()
Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(cmbCostCenter.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)

End Sub

Private Sub cmbCostCenter_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
mskDateDue.SetFocus
End If
End Sub

Private Sub cmbCurrency_GotFocus()
Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(cmbCurrency.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)

End Sub

Private Sub cmbCurrency_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtFCamount.SetFocus
End If

End Sub

Private Sub CmbNotedBy_click()
'txtNotedBy.Text = Trim(CmbNotedBy.Text)

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

Private Sub cmbPayee2_GotFocus()
Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(cmbPayee2.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)

End Sub

Private Sub cmbPayee2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtDocuNo.SetFocus
End If
End Sub


Private Sub cmbPayment_Click()
If cmbPayment.Text = "Forign Currency" Then
cmbCurrency.Enabled = True
txtFCamount.Enabled = True
End If

End Sub

Private Sub cmbPayment_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If cmbPayment.Text = "Forign Currency" Then
cmbCurrency.SetFocus
End If
End If

End Sub

Private Sub cmbPaymentFor_Click()
If Me.CmbSource = "" Then
MsgBox "You should select the Payment To", vbInformation, "Alert"
End If

End Sub

Private Sub cmbPaymentFor_GotFocus()
Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(cmbPaymentFor.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)

End Sub




Private Sub cmbPaymentFor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbPayee2.SetFocus
End If

End Sub

Private Sub cmbPaymode_GotFocus()
Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(cmbPaymode.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)

End Sub

Private Sub cmbPaymode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbPayment.SetFocus
End If
End Sub

Private Sub CmbPrepBy_GotFocus()


Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(CmbPrepBy.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)

End Sub

Private Sub CmbPrepBy_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 If cmdNew.caption <> "&New" Then
'frmPassword.Show (1)
CmbNotedBy.SetFocus
End If
End If
End Sub

Private Sub cmbProfCenter_Click()
'Dim serialpro As String
'Dim serialdep As String
'Dim serialxsec As String
'cmbCostCenter.clear
'
''        If rstCosPro.EOF = False Then
'        rstCosPro.MoveFirst
' '       End If
'
'
'Do Until rstCosPro.EOF
'
'        If cmbProfCenter = rstCosPro!profitcenter Then
'            If serialpro <> rstCosPro!productionUnit Then
'            cmbCostCenter.AddItem rstCosPro!productionUnit
'            End If
'            serialpro = rstCosPro!productionUnit
'
'             If serialdep <> rstCosPro!department Then
'            cmbCostCenter.AddItem rstCosPro!department
'            End If
'            serialdep = rstCosPro!department
'
'            If serialxsec <> rstCosPro!xsection Then
'            cmbCostCenter.AddItem rstCosPro!xsection
'            End If
'            serialxsec = rstCosPro!xsection
'
'
'        End If
'
'         rstCosPro.MoveNext
'  Loop

End Sub

Private Sub cmbProfCenter_GotFocus()
Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(cmbProfCenter.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)

End Sub

Private Sub cmbProfCenter_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbCostCenter.SetFocus
End If

End Sub

Private Sub CmbSource_GotFocus()
Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(CmbSource.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)

End Sub

Private Sub CmbSource_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbPaymentFor.SetFocus
End If

End Sub




Private Sub cmdExit1_Click()
If cmdExit1.caption = "E&xit" Then  'This is to Exit From form
Unload Me

Else    'This is to Cancel the Job

If Me.cmdNew.caption = "&Save" Then
    Dim serv As New ADODB.Recordset
    serv.Open "select * from SerialNom", constring, adOpenDynamic, adLockOptimistic
    If serv.EOF = False Then
    serv.MoveFirst
    End If
    While serv.EOF = False
     txtserialNo.Text = serv!SENo
     serv!SENo = serv!SENo - 1
     serv.Update
    
    serv.MoveNext
    Wend
End If


txtNoOfAtt.clear
List1.clear
cmdPrintPayReq.Visible = False
frmMenu.shedit.Enabled = True
FrmPayableSetup.CmdEdit.caption = Trim("E&dit")
FrmPayableSetup.ListView2.ListItems.clear
'FrmPayableSetup.ListView3.ListItems.clear
frmMenu.sEdit.caption = "Edit" 'Once Press Cancel this Caption Agian will be "Edit"
frmMenu.edit.caption = "Edit" 'Once Press Cancel this Caption Agian will be "Edit"
cmdNew.Visible = True
cmdNew.caption = "&New"
Command1.caption = "&New"
cmdExit1.caption = "E&xit"

 txtAmountDue.Enabled = False
 'txtApprBy.Enabled = False
 txtBranch.Enabled = False
 txtDocuNo.Enabled = False
 mskDateDue.Enabled = False
' txtNotedBy.Enabled = False
 'txtPrepBy.Enabled = False
 txtRefNo.Enabled = False
 CmbApprovedBy.Enabled = False
 cmbCostCenter.Enabled = False
 CmbNotedBy.Enabled = False

 CmbSource.Enabled = False
' Combo5.Enabled = False
'txtStorEntryNo.Enabled = False
'mskStoreEnDate.Enabled = False
'txtvenCode.Enabled = False
cmbPaymentFor.Enabled = False
cmbProfCenter.Enabled = False
txtInvAmt.Enabled = False
'txtInvNo.Enabled = False

List1.Text = ""

For Each Control In Me
    If TypeOf Control Is TextBox Then
        Control.Text = ""
    End If
Next

For Each Control In Me
    If TypeOf Control Is ComboBox Then
        Control.Text = ""
    End If
    
Next

For Each Control In Me
    If TypeOf Control Is MaskEdBox Then
        Control.Text = "__/__/____"
    End If
Next



FrmPayableSetup.txtTMDays.Visible = False
FrmPayableSetup.txtTMDes.Visible = False
FrmPayableSetup.txtTMlevel.Visible = False
FrmPayableSetup.txtTMmode.Visible = False
FrmPayableSetup.txtTmRate.Visible = False

End If
End Sub

Private Sub CmdExit2_Click()
If CmdExit2.caption = "&Exit" Then
Unload Me
Else
cmdExit1_Click
End If
End Sub

Private Sub cmdNew_Click()
Me.Timer1.Enabled = True
Dim JOurnalNo As ADODB.Recordset
Set JOurnalNo = New ADODB.Recordset

If cmdNew.caption = "&New" Then

    Option1.Enabled = True
    Option2.Enabled = True
    cmbPayment.Enabled = True
    txtEnter1.Enabled = True
    TxtEnter2.Enabled = True
   'cmbCurrency.Enabled = True
    txtCreditNote.Enabled = True
    txtDebitNote.Enabled = True
    txtOutBal.Enabled = True
    txtAmtPaid.Enabled = True
    cmbPayee2.Enabled = True
    txtAmtReq.Enabled = True
    cmbPaymode.Enabled = True
    cmdNew.caption = "&Save"
    Command1.caption = "&save"
    cmdExit1.caption = "&Cancel"
    txtAmountDue.Enabled = True
    txtBranch.Enabled = True
    txtDocuNo.Enabled = True
    mskDateDue.Enabled = True
    txtRefNo.Enabled = True
    CmbApprovedBy.Enabled = True
    cmbCostCenter.Enabled = True
    CmbNotedBy.Enabled = True
    CmbPrepBy.Enabled = True
    CmbSource.Enabled = True
    cmbPaymentFor.Enabled = True
    cmbProfCenter.Enabled = True
    
    MskDate.Enabled = True
    List1.Enabled = True
    cmbPaymentFor.Enabled = True


'This is Temp for the Test
Command3.Enabled = True
txtInvAmt.Enabled = True
'-----------------------



FrmPayableSetup.txtTMDays.Visible = True
FrmPayableSetup.txtTMDes.Visible = True
FrmPayableSetup.txtTMlevel.Visible = True
FrmPayableSetup.txtTMmode.Visible = True
FrmPayableSetup.txtTmRate.Visible = True

MskDate.Text = Format(Date, "dd/mm/yyyy")
cmbPaymode.SetFocus
On Error Resume Next

If rstPaySetup.EOF Then
End If
On Error GoTo 0
 
 

Dim Ser As New ADODB.Recordset
Dim serv As ADODB.Recordset
Set serv = New ADODB.Recordset
serv.Open "select * from SerialNom", constring, adOpenDynamic, adLockOptimistic
'If serv.EOF = False Then
'serv.MoveFirst
'End If
'While serv.EOF = False
'
' txtSerialNo.Text = serv!SENo
' serv!SENo = serv!SENo + 1
' serv.Update
'serv.MoveNext
'Wend


'This is the New One
    Dim rsPaySetupz As New ADODB.Recordset
    rsPaySetupz.Open "PayableSetup order by serialno", constring, adOpenKeyset, adLockPessimistic, adCmdTable
    If rsPaySetupz.EOF = False Then
     rsPaySetupz.MoveFirst
     rsPaySetupz.MoveLast
     Me.txtserialNo.Text = Val(rsPaySetupz!SerialNo) + 1
    End If








' This is to Save the Records
 ElseIf cmdNew.caption = "&Save" Then
 
 
 
 
     
        If txtAmtReq.Text = "" Then
        MsgBox "You have to Fill the Textbox 'AmountRequested'", vbInformation, "Try Again"
        Exit Sub
        End If
 
        If cmbPaymentLevel.Text = "" Then
        MsgBox "You have to Fill the Combo 'Payment Level'", vbInformation, "Try Again"
        Exit Sub
        End If

        If Me.cmbPaymentFor = "" Then
        MsgBox "You have to Fill the Combo 'PaymentFor' ", vbInformation, "Try Again"
        Exit Sub
        End If
         
        If Me.cmbCostCenter = "" Then
        MsgBox "You have to Fill the combo 'CostCenter'", vbInformation, "Try Again"
        Exit Sub
        End If
 
        If Me.CmbPrepBy = "" Then
        MsgBox "You have to Fill the Combo 'PrepBy'", vbInformation, "Try Again"
        Exit Sub
        End If
    
        If mskDateDue = "__/__/____" Then
        MsgBox "You have to Fill the Box DateDue", vbInformation, "Try Again"
        Exit Sub
        End If
 
         If Me.txtInvAmt = "" Then
        MsgBox "Invoice Amount is Empty", vbExclamation, "Operation Cancelled"
        
        'Here i can Validate the Invoice Amount in the Future
        
        Exit Sub
        End If
 
 
         X = MsgBox("Are You sure Adding this Records ?", vbYesNo, "SAVE")
        
        If X = vbNo Then
          Exit Sub
          
         End If

 
 
 Me.Text1.Text = Me.txtserialNo.Text 'this is to put the Previos SerialNo
 
 
 
 
    Dim rsPaySetup As New ADODB.Recordset
    rsPaySetup.Open "PayableSetup order by serialno", constring, adOpenKeyset, adLockPessimistic, adCmdTable
    If rsPaySetup.EOF = False Then
     rsPaySetup.MoveFirst
     rsPaySetup.MoveLast
     Me.txtserialNo.Text = Val(rsPaySetup!SerialNo) + 1
    End If
    
     
     

     
     
'    If Me.cmbPaymentFor = "" Or Me.cmbCostCenter.Text = "" Or Me.CmbPrepBy = "" Or mskDateDue = "__/__/____" Or txtAmtReq = "" Then
'    MsgBox "There is no Sufficient Records to Save ", vbCritical, "Not Saved"
'    Me.cmbPaymentFor.SetFocus
'    Exit Sub
'    End If





'This will Immediately call the Password Form once we select "YES" to save from MSG Box
frmPassword.txtBuffer = "Save"
frmPassword.txtPrepBy.Text = Me.CmbPrepBy.Text 'Useful to varify the password for him
frmPassword.Show 1
On Error Resume Next
frmPassword.txtUserId.SetFocus 'From the Password it will call SaveMe

ElseIf cmdNew.caption = "&Update" Then




    If Me.cmbPaymentFor = "" Or Me.cmbCostCenter.Text = "" Or Me.CmbPrepBy = "" Or mskDateDue = "__/__/____" Or txtAmtReq = "" Or cmbPaymentLevel = "" Then
    MsgBox "There is no Sufficient Records to Update ", vbCritical, "Not Saved"
    'Me.cmbPaymentFor.SetFocus
    Exit Sub
    End If



        X = MsgBox("Are You sure Updating this Records ?", vbYesNo, "UPDATE")
        
        If X = vbNo Then
          Exit Sub
          
         End If


    frmPassword.txtBuffer = "Update"
    frmPassword.txtPrepBy.Text = Me.CmbPrepBy.Text 'Useful to varify the password for him
    Dim rstrepo1 As New ADODB.Recordset
    rstrepo1.Open "Delete  ReportPayReq", constring, adOpenDynamic, adLockOptimistic, adCmdText
    frmPassword.Show 1
On Error Resume Next

frmPassword.txtUserId.SetFocus 'From the Password it will call SaveMe


End If

End Sub
Public Sub UpdateMe()



'This is to Save the Invoice Deteils
Dim rsTempPayInv As New ADODB.Recordset
Dim seria
seria = Me.txtserialNo.Text
rsTempPayInv.Open "Select * from PayTEMPInvoiceDetails where serialno = " & "'" & seria & "'" & "", constring, adOpenDynamic, adLockOptimistic

     Dim vInvNo, vInvDate, vDueDate, vInvAmt, vTradeDis, vCashDis, vSalesDis, vServDis, vComProf, vAddDedTax1, vCustDut, vTotInv, vPONumber, vPODate, vSENumber, vSEDate
     
 Dim rsPayInv22 As New ADODB.Recordset
 rsPayInv22.Open "Select * from PayInvoiceDetails", constring, adOpenDynamic, adLockOptimistic
     
     
     'SerialNo
  While rsTempPayInv.EOF = False
  
      vInvNo = rsTempPayInv!InvNo
       vInvDate = rsTempPayInv!InvDate
       vDueDate = rsTempPayInv!duedate
        vInvAmt = rsTempPayInv!invAmt
        vTradeDis = rsTempPayInv!TradeDis
       vCashDis = rsTempPayInv!CashDis
       vSalesDis = rsTempPayInv!SalesDis
       vServDis = rsTempPayInv!ServDis
      vComProf = rsTempPayInv!ComProf
       vAddDedTax1 = rsTempPayInv!AddDedTax1
       vCustDut = rsTempPayInv!CustDut
        vTotInv = rsTempPayInv!TotInv
        vPONumber = rsTempPayInv!PoNumber
       vPODate = rsTempPayInv!PODate
       vSENumber = rsTempPayInv!SENumber
       vSEDate = rsTempPayInv!SEDate
       
       
                                  With rsPayInv22
                            .addnew
                            !SerialNo = txtserialNo.Text
                            !InvNo = vInvNo
                            !InvDate = vInvDate
                            !duedate = vDueDate
                            !invAmt = vInvAmt
                            !TradeDis = vTradeDis
                            !CashDis = vCashDis
                            !SalesDis = vSalesDis
                            !ServDis = vServDis
                            !ComProf = vComProf
                            !AddDedTax1 = vAddDedTax1
                            !CustDut = vCustDut
                            !TotInv = vTotInv
                            !PoNumber = vPONumber
                            !PODate = vPODate
                            !SENumber = vSENumber
                            !SEDate = vSEDate
                            
                            .Update
                            End With
 
       
       
   rsTempPayInv.MoveNext
   Wend
'---------------------------------------------
rsPayInv22.close

'Here I have to Delete the Temporary PayInvoice  Tablce

Dim rsTempPayInvDelete2 As New ADODB.Recordset
rsTempPayInvDelete2.Open "Delete  from PayTEMPInvoiceDetails where SerialNo= " & "'" & seria & "'" & "", constring, adOpenDynamic, adLockOptimistic

'Here Delete the TempPayInv for the Seralno


Dim rstPaysetupNew As New ADODB.Recordset
rstPaysetupNew.Open "Select * from Payablesetup", constring, adOpenDynamic, adLockOptimistic


'this will Update the Existing Datas
'If cmdNew.Caption = "&Update" Then
'If cmdNew.Caption = "&Update" Then
        rstPaysetupNew.MoveFirst

  On Error GoTo 0
   Dim TextSerNo As Long
   TextSerNo = Me.txtserialNo.Text
      While rstPaysetupNew.EOF = False
        If Trim(rstPaysetupNew!SerialNo) = Trim(TextSerNo) Then

            rstPaysetupNew!SerialNo = txtserialNo.Text
            rstPaysetupNew!payto = CmbSource.Text
            rstPaysetupNew!taxCrdit = txtTaxCredit.Text
            rstPaysetupNew!Requester = txtRefNo.Text
            rstPaysetupNew!payfor = cmbPaymentFor.Text
            rstPaysetupNew!DocNo = txtDocuNo.Text
            rstPaysetupNew!branch = txtBranch.Text
            rstPaysetupNew!Explanation = List1.Text
            rstPaysetupNew!NoOfAttech = TxtEnter2.Text
            'rstPaysetupNew!PoNumber = txtPoNo.Text
            rstPaysetupNew!paymode = cmbPaymode.Text
            rstPaysetupNew!costcenter = cmbCostCenter.Text
            rstPaysetupNew!ProfCenter = cmbProfCenter.Text
           ' rstPaysetupNew!PODate = mskPOdate.Text

            rstPaysetupNew!debitnote = txtCreditNote.Text
            rstPaysetupNew!creditnote = txtDebitNote.Text
'            rstPaysetupNew!taxcredit = txtTaxCredit.Text

            rstPaysetupNew!PaymentLevels = cmbPaymentLevel.Text

            rstPaysetupNew!FcAmount = IIf((txtFCamount.Text) = "", 0, (txtFCamount.Text))
            
            If mskDateDue.Text = "__/__/____" Then
            Else
            rstPaysetupNew!DateDue = Format(mskDateDue.Text, "mm/dd/yyyy")
            End If
            
            rstPaysetupNew!ExCurrency = cmbCurrency.Text
            rstPaysetupNew!FcAmount = IIf((txtFCamount.Text) = "", 0, (txtFCamount.Text))
            rstPaysetupNew!Payee = cmbPayee2.Text

            If Me.MskDate.Text = "__/__/____" Then
            Else
            rstPaysetupNew!Xdate = Format(MskDate.Text, "dd/mm/yyyy")
            End If
            
            rstPaysetupNew!AmtPaidBefore = txtAmtPaid.Text
            rstPaysetupNew!amtreqested = txtAmtReq.Text
            rstPaysetupNew!outbal = txtOutBal.Text
            On Error Resume Next
            rstPaysetupNew!invAmt = IIf(IsNull(txtInvAmt.Text), "0.00", (txtInvAmt.Text))
            On Error GoTo 0


            y = 0
            For y = 0 To txtNoOfAtt.ListCount
             rstPaysetupNew!NoOfAttech = rstPaysetupNew!NoOfAttech & " " & txtNoOfAtt.List(y)
            Next
            
            i = 0
            For i = 0 To Me.List1.ListCount
            rstPaysetupNew!Explanation = rstPaysetupNew!Explanation & " " & Me.List1.List(i)
            Next

           '
        'rstPaysetupNew!NoOfAttech = txtNoOfAtt.Text
            rstPaysetupNew!AmtDue = txtAmountDue.Text
            rstPaysetupNew!RefNo = txtRefNo.Text
            rstPaysetupNew!cancelledmark = 0

rstPaysetupNew.Update

MsgBox "Details Updated Successfully", vbInformation, "Updated"
cmdPrintPayReq.Visible = True
cmdNew.Visible = False
CmdEdit.Value = False
Exit Sub

    End If

   rstPaysetupNew.MoveNext
   Wend
cmdPrintPayReq.Visible = True
End Sub
Public Sub saveme()


    Dim rsPaySetupk As New ADODB.Recordset
    rsPaySetupk.Open "PayableSetup order by serialno", constring, adOpenKeyset, adLockPessimistic, adCmdTable
    If rsPaySetupk.EOF = False Then
     rsPaySetupk.MoveFirst
     rsPaySetupk.MoveLast
     Me.txtserialNo.Text = Val(rsPaySetupk!SerialNo) + 1
    End If
  rsPaySetupk.close


'This is to Save the Invoice Deteils
Dim rsTempPayInv As New ADODB.Recordset
Dim seria
seria = Me.Text1.Text
rsTempPayInv.Open "Select * from PayTEMPInvoiceDetails where serialno = " & "'" & seria & "'" & "", constring, adOpenDynamic, adLockOptimistic

     Dim vInvNo, vInvDate, vDueDate, vInvAmt, vTradeDis, vCashDis, vSalesDis, vServDis, vComProf, vAddDedTax1, vCustDut, vTotInv, vPONumber, vPODate, vSENumber, vSEDate
     
 Dim rsPayInv2 As New ADODB.Recordset
 rsPayInv2.Open "Select * from PayInvoiceDetails", constring, adOpenDynamic, adLockOptimistic
     
     
     'SerialNo
  While rsTempPayInv.EOF = False
  
      vInvNo = rsTempPayInv!InvNo
       vInvDate = rsTempPayInv!InvDate
       vDueDate = rsTempPayInv!duedate
        vInvAmt = rsTempPayInv!invAmt
        vTradeDis = rsTempPayInv!TradeDis
       vCashDis = rsTempPayInv!CashDis
       vSalesDis = rsTempPayInv!SalesDis
       vServDis = rsTempPayInv!ServDis
      vComProf = rsTempPayInv!ComProf
       vAddDedTax1 = rsTempPayInv!AddDedTax1
       vCustDut = rsTempPayInv!CustDut
        vTotInv = rsTempPayInv!TotInv
        vPONumber = rsTempPayInv!PoNumber
       vPODate = rsTempPayInv!PODate
       vSENumber = rsTempPayInv!SENumber
       vSEDate = rsTempPayInv!SEDate

     
                            With rsPayInv2
                            .addnew
                            !SerialNo = txtserialNo.Text
                            !InvNo = vInvNo
                            !InvDate = vInvDate
                            !duedate = vDueDate
                            !invAmt = vInvAmt
                            !TradeDis = vTradeDis
                            !CashDis = vCashDis
                            !SalesDis = vSalesDis
                            !ServDis = vServDis
                            !ComProf = vComProf
                            !AddDedTax1 = vAddDedTax1
                            !CustDut = vCustDut
                            !TotInv = vTotInv
                            !PoNumber = vPONumber
                            !PODate = vPODate
                            !SENumber = vSENumber
                            !SEDate = vSEDate
                            
                            .Update
                            End With


rsTempPayInv.MoveNext
Wend




'Here I have to Delete the Temporary PayInvoice  Tablce

Dim rsTempPayInvDelete As New ADODB.Recordset
rsTempPayInvDelete.Open "Delete  from PayTEMPInvoiceDetails where SerialNo= " & "'" & seria & "'" & "", constring, adOpenDynamic, adLockOptimistic

Dim rstPaySetupAddnew As New ADODB.Recordset
rstPaySetupAddnew.Open "Select * from Payablesetup", constring, adOpenDynamic, adLockOptimistic


Dim Var
Dim Var22122002
rstPaySetupAddnew.addnew
            
            rstPaySetupAddnew!SerialNo = txtserialNo.Text
            rstPaySetupAddnew!payto = CmbSource.Text
            rstPaySetupAddnew!payfor = cmbPaymentFor.Text
            rstPaySetupAddnew!DocNo = txtDocuNo.Text
            rstPaySetupAddnew!branch = txtBranch.Text
            rstPaySetupAddnew!taxCrdit = txtTaxCredit.Text

            i = 0
            For i = 0 To Me.List1.ListCount
            rstPaySetupAddnew!Explanation = rstPaySetupAddnew!Explanation & " " & Me.List1.List(i)
            Next
            
            rstPaySetupAddnew!debitnote = txtCreditNote.Text
            rstPaySetupAddnew!creditnote = txtDebitNote.Text
            rstPaySetupAddnew!paymode = cmbPaymode.Text
            rstPaySetupAddnew!Payee = cmbPayee2.Text
            rstPaySetupAddnew!Requester = txtRefNo.Text
            rstPaySetupAddnew!ExCurrency = cmbCurrency.Text
            rstPaySetupAddnew!FcAmount = IIf((txtFCamount.Text) = "", 0, (txtFCamount.Text))
            rstPaySetupAddnew!PaymentLevels = cmbPaymentLevel.Text

            On Error Resume Next

            rstPaySetupAddnew!costcenter = cmbCostCenter.Text
            rstPaySetupAddnew!ProfCenter = cmbProfCenter.Text
            
            If mskDateDue.Text = "__/__/____" Then
            On Error Resume Next
            rstPaySetupAddnew!DateDue = mskDateDue.ClipText
            On Error GoTo 0
            Else
            rstPaySetupAddnew!DateDue = Format(mskDateDue.Text, "dd/mm/yyyy")
            End If
            
            
            
            If MskDate.Text = "__/__/____" Then
            On Error Resume Next
            rstPaySetupAddnew!Xdate = MskDate.ClipText
            On Error GoTo 0
            Else
            rstPaySetupAddnew!Xdate = Format(MskDate.Text, "dd/mm/yyyy")
            End If

            
            If txtAmtPaid = "" Then
            rstPaySetupAddnew!AmtPaidBefore = 0
            Else
            rstPaySetupAddnew!AmtPaidBefore = txtAmtPaid.Text
            End If
            
            
            If txtAmtReq = "" Then
            rstPaySetupAddnew!amtreqested = 0
            Else
            rstPaySetupAddnew!amtreqested = txtAmtReq.Text
            End If
            
            
            If txtOutBal = "" Then
            rstPaySetupAddnew!outbal = 0
            Else
            rstPaySetupAddnew!outbal = txtOutBal.Text
            End If
            
            
            rstPaySetupAddnew!Prepby = cLogUser
            rstPaySetupAddnew!NotedBy = IIf(IsNull(CmbNotedBy.Text), "", (CmbNotedBy.Text))
            rstPaySetupAddnew!AppBy = IIf(IsNull(CmbApprovedBy.Text), "", (CmbApprovedBy.Text))
            
          On Error Resume Next
          
          
        If rstReceipt.EOF = False Then
        rstReceipt.MoveFirst
        End If
  
          
            While rstReceipt.EOF = False
            If Me.txtserialNo.Text = rstReceipt!SerialNo Then
            Var = Trim(rstReceipt!total)
            End If
            rstReceipt.MoveNext
            Wend
            
 '          rstPaySetupAddnew!totcramt = IIf(IsNull(txtTotList3.Text), "", (txtTotList3.Text))
            rstPaySetupAddnew!TotCrAmt = Trim(Var)
            On Error GoTo 0
'            rstPaySetupAddnew!invoiceno = IIf(IsNull(txtinvNo.Text), "", (txtinvNo.Text))
            
            
            If txtInvAmt.Text = "" Then
            rstPaySetupAddnew!invAmt = 0
            Else
            rstPaySetupAddnew!invAmt = txtInvAmt.Text
            End If
            
            
            y = 0
            For y = 0 To txtNoOfAtt.ListCount
             rstPaySetupAddnew!NoOfAttech = rstPaySetupAddnew!NoOfAttech & " " & txtNoOfAtt.List(y)
            Next
            
            rstPaySetupAddnew!AmtDue = txtAmountDue.Text
            rstPaySetupAddnew!RefNo = txtRefNo.Text
            rstPaySetupAddnew!cancelledmark = 0
            rstPaySetupAddnew!ConfirmedMark = 0
            rstPaySetupAddnew!deletemark = 0
            rstPaySetupAddnew!Paidmark = 0
            rstPaySetupAddnew!Post = "No"
rstPaySetupAddnew.Update
rstPaySetupAddnew.close
              
              


         'this is to save the listview(Terms) details to a table  (RATE)
              n = 0
              For n = 1 To Me.ListView2.ListItems.Count
                  Sn = Me.txtserialNo.Text
                  'ra = Me.ListView2.SelectedItem
                  ra = Me.ListView2.ListItems.Item(n)
                  des = Me.ListView2.ListItems.Item(n).SubItems(1)
                  xdays = Me.ListView2.ListItems.Item(n).SubItems(2)
                  xmode = Me.ListView2.ListItems.Item(n).SubItems(3)
                  xlevel = Me.ListView2.ListItems.Item(n).SubItems(4)
                     With rstterm
                           .addnew
                           !SerialNo = Sn
                           !rate = ra
                           !descr = des
                           !days = xdays
                           !Mode = xmode
                           !xlevel = xlevel
                           'On Error Resume Next
                           .Update
                      End With
               Next

             MsgBox "records Has Been Added Successfully ", vbInformation, "SAVE"


'this is to add Listview 1
Me.ListView1.ListItems.clear

Dim rstpaysetNew As New ADODB.Recordset
rstpaysetNew.Open "Select * from Payablesetup where deletemark = '0'  and Post = 'No' and cancelledmark = 0 and ConfirmedMark = 0 and Paidmark = 0 order by SerialNo", constring, adOpenDynamic, adLockOptimistic

        If rstpaysetNew.EOF = False Then
        rstpaysetNew.MoveFirst
        End If

  While rstpaysetNew.EOF = False
 ' If Trim(rstPaySetup!cancelledmark) = 0 And Trim(rstPaySetup!ConfirmedMark) = 0 And Trim(rstPaySetup!Paidmark) = 0 Then   'This is not Cancelled
  
     Set MItem = Me.ListView1.ListItems.Add(, , Format(rstpaysetNew!SerialNo))
     MItem.SubItems(1) = Format(rstpaysetNew!Xdate, "dd/mm/yyyy")
     MItem.SubItems(2) = Format(rstpaysetNew!Requester)
     MItem.SubItems(3) = Format(rstpaysetNew!DateDue, "dd/mm/yyyy")
     MItem.SubItems(4) = Format(rstpaysetNew!RefNo)
       MItem.SubItems(5) = Format(rstpaysetNew!printmark)
       MItem.SubItems(6) = Format(rstpaysetNew!journaledmark)
     MItem.SubItems(7) = Format(rstpaysetNew!amtreqested, "##########.#0")
        
        
     On Error Resume Next
       Totlist1 = Val(Totlist1) + Val(Trim(rstpaysetNew!amtreqested)) 'This is for the Total of the List
     On Error GoTo 0
 'End If
     rstpaysetNew.MoveNext
     Wend


txtTotList1.Text = Totlist1

rstpaysetNew.close
On Error GoTo 0


cmdNew.Visible = False
cmdPrintPayReq.Visible = True
cmdprint.Visible = False
CmdEdit.Visible = False

Dim serv As ADODB.Recordset
Set serv = New ADODB.Recordset

serv.Open "select * from SerialNom", constring, adOpenDynamic, adLockOptimistic
If serv.EOF = False Then
serv.MoveFirst
End If
While serv.EOF = False
'txtSerialNo.Text = Serv!seNo
'Pee = "P"
'Zero = "00000"
''Zeros & Trim(Val(nextjn) + 1)
'Serv!seNo = Zero & Trim(Val(Serv!seNo) + 1)
serv!SENo = Trim(Val(serv!SENo) + 1)
serv.Update

serv.MoveNext
Wend

'cmdExit1.Visible = False
'cmdPrintPayReq_Click
End Sub

Private Sub cmdPrintPayReq_Click()
Dim conx As ADODB.Connection
Set conx = New ADODB.Connection
'conString = "Provider=MSDASQL;DSN=payrollcairo;UID=; PWD=;"
conStr1 = "Provider=MSDASQL;DSN=Finance;UID=sa; PWD=;"

conx.Open conStr1

'is for the New Table
Dim PSSno 'this is for the Payable setup Serial Numbetr
Dim nSerialNo, ncode, npayto, npayfor, nPaymentFC, nPaymentLevel, npaymode, nvencode, npayee2, nDebitnote, nCreditnote, nCurrency, nPayee, nDocNo, nBranch, nExplanation, nProfCenter, nCostCenter, taxCr, nDateDue, nStoreEntNo, nStoreEntDate, nPoNumber, nInvoiceno, nInvAmt, nXDate, nAmtPaidBefore, nAmtRequested, nOutBal, nProfitTax, nRemarks, nPercentage, nNoOfAttech, nDebitAC, nAmtDue, nRefNo, TotCrAmt, nRemainingBal, nTaxCredit
Dim nInvoiceDate As Date
Dim npoDate As String
Dim Desc1, Rate1, Days1, Mode1, xlevel1, Desc2, Rate2, Days2, Mode2, xlevel2, Desc3, Rate3, Days3, Mode3, xlevel3, Desc4, Rate4, Days4, Mode4, xlevel4, Desc5, Rate5, Days5, Mode5, xlevel5
Dim Accnox1, AccNamex1, DebAmt1, Accnox2, AccNamex2, DebAmt2, Accnox3, AccNamex3, DebAmt3, Accnox4, AccNamex4, DebAmt4, Accnox5, AccNamex5, DebAmt5, Accnox6, AccNamex6, DebAmt6, Accnox7, AccNamex7, DebAmt7, Accnox8, AccNamex8, DebAmt8, Accnox9, AccNamex9, DebAmt9, Accnox10, AccNamex10, DebAmt10 'is for the xPayment
Dim Accnoy1, AccNamey1, CRAmt1, Accnoy2, AccNamey2, CrAmt2, Accnoy3, AccNamey3, CrAmt3, Accnoy4, AccNamey4, CrAmt4, Accnoy5, AccNamey5, CrAmt5, Accnoy6, AccNamey6, CrAmt6, Accnoy7, AccNamey7, CrAmt7, Accnoy8, AccNamey8, CrAmt8, Accnoy9, AccNamey9, CrAmt9, Accnoy10, AccNamey10, CrAmt10 'is for the xReceipt
Dim invno1, invno2, InvNo3, InvNo4, InvNo5, InvNo6, InvDate1, InvDate2, InvDate3, InvDate4, InvDate5, InvDate6, DueDate1, DueDate2, DueDate3, DueDate4, DueDate5, DueDate6

'is for the particular tables
Dim Meena, Ratha, Kajol, Maina



Dim rstpay As ADODB.Recordset  'this is for the Payable setup
Set rstpay = New ADODB.Recordset
Dim rstTerms As ADODB.Recordset 'is for the Term
Set rstTerms = New ADODB.Recordset
Dim rstxpay As ADODB.Recordset 'is for the XPayment
Set rstxpay = New ADODB.Recordset
Dim rstxRec As ADODB.Recordset 'is for the XReceipt
Set rstxRec = New ADODB.Recordset
Dim RstRepodel As ADODB.Recordset
Set RstRepodel = New ADODB.Recordset 'is for the new table for Addnew
Dim xRstInv As New ADODB.Recordset
'rstNewPay.Open "select * from PayAnalForPrint", conx, adOpenDynamic, adLockOptimistic

RstRepodel.Open "Delete  ReportPayReq", conx, adOpenDynamic, adLockOptimistic, adCmdText

If cmdPrintPayReq.caption <> "Print Payment &Request  ØÈÇÚÉ ØáÈ ÇáãÏÝæÚÇÊ " Then

'Delete Table
'RstRepo.Close

'PaymentRequestLast.refresh

'RstRepo.Close
cmdPrintPayReq.caption = "Print Payment &Request  ØÈÇÚÉ ØáÈ ÇáãÏÝæÚÇÊ "
'cmdPrintPayReq.Visible = False
cmdExit1.Visible = True
Exit Sub
Else
    
    
    
    PSSno = FrmPayableSetup.txtserialNo.Text
    
    

    
'    RstRepo.Open "select * from ReportPayReq", conx, adOpenDynamic, adLockOptimistic

   ppp = "Select * from PayableSetup where serialno = " & "'" & PSSno & "'" & ""
   rstpay.Open ppp, conx, adOpenDynamic, adLockOptimistic
If rstpay.EOF = False Then
rstpay.MoveFirst
While rstpay.EOF = False     'OPEN PAYABLE SETUP
' PSSno = rstpay!serialno

    xPundamora = "select * from PayInvoiceDetails where SerialNo = " & "'" & PSSno & "'" & ""
    xRstInv.Open xPundamora, conx, adOpenDynamic, adLockOptimistic

     
    AAAk = "select * from  Term where serialno= " & "'" & PSSno & "'" & ""
    rstTerms.Open AAAk, conx, adOpenDynamic, adLockOptimistic
    
    xxx = "select * from  xpayment where serialno= " & "'" & PSSno & "'" & ""
    rstxpay.Open xxx, conx, adOpenDynamic, adLockOptimistic
    
    yyy = "select * from  xreceipt where serialno= " & "'" & PSSno & "'" & ""
    rstxRec.Open yyy, conx, adOpenDynamic, adLockOptimistic
      
      
      
      PSSno = rstpay!SerialNo
      ncode = rstpay!Code
      npayto = rstpay!payto
      npayfor = rstpay!payfor
      npaymode = rstpay!paymode
      nvencode = rstpay!vencode
      nPayee = rstpay!Payee
      nCurrency = rstpay!ExCurrency
     ' npoDate = rstpay!PODate ', "mm/dd/yyyy")
     ' nTaxCredit = rstpay!taxcredit
     ' nPayee = rstpay!payee
      nDocNo = rstpay!DocNo
      nBranch = rstpay!branch
      nExplanation = rstpay!Explanation
      nProfitCenter = rstpay!ProfCenter
      nCostCenter = rstpay!costcenter
      nDateDue = rstpay!DateDue
      taxCr = rstpay!taxCrdit
      nDebitnote = rstpay!debitnote
      nCreditnote = rstpay!creditnote
     'nStoreEntNo = rstpay!StoreEntNo
      'nPoNumber = rstpay!PoNumber
     ' nInvoiceno = rstpay!invoiceno
      On Error Resume Next
    '  nInvoiceDate = Format(rstpay!invoicedate, "mm/dd/yyyy")
      nInvAmt = rstpay!invAmt
      nXDate = rstpay!Xdate
      nAmtPaidBefore = rstpay!AmtPaidBefore
      nAmtRequested = rstpay!amtreqested
      nOutBal = rstpay!outbal
      'nProfitTax = rstpay!ProfitTax
      nRemarks = rstpay!remarks
     ' nPercentage = rstpay!Percentage
      nNoOfAttech = rstpay!NoOfAttech
      nDebitAC = rstpay!DebitAC
      nAmtDue = rstpay!AmtDue
      nRefNo = rstpay!RefNo
      nRemainingBal = Val(nOutBal) - Val(nAmtRequested)
      nPaymentFC = rstpay!FcAmount
      nPaymentLevel = rstpay!PaymentLevels
      
                Meena = 1
                If rstTerms.EOF = False Then
                rstTerms.MoveFirst
                While rstTerms.EOF = False
                If Meena = 1 Then
                Desc1 = rstTerms!descr
                Rate1 = rstTerms!rate
                Mode1 = rstTerms!Mode
                Days1 = rstTerms!days
                xlevel1 = rstTerms!xlevel
                ElseIf Meena = 2 Then
                Desc2 = rstTerms!descr
                Rate2 = rstTerms!rate
                Mode2 = rstTerms!Mode
                Days2 = rstTerms!days
                xlevel2 = rstTerms!xlevel
                ElseIf Meena = 3 Then
                Desc3 = rstTerms!descr
                Rate3 = rstTerms!rate
                Mode3 = rstTerms!Mode
                Days3 = rstTerms!days
                xlevel3 = rstTerms!xlevel
                ElseIf Meena = 4 Then
                Desc4 = rstTerms!descr
                Rate4 = rstTerms!rate
                Mode4 = rstTerms!Mode
                Days4 = rstTerms!days
                xlevel4 = rstTerms!xlevel
                ElseIf Meena = 5 Then
                Desc5 = rstTerms!descr
                Rate5 = rstTerms!rate
                Mode5 = rstTerms!Mode
                Days5 = rstTerms!days
                xlevel5 = rstTerms!xlevel
                
                
                
                End If
                 Meena = Meena + 1
                rstTerms.MoveNext
                Wend
                End If
                               If rstxpay.EOF = False Then
                                Ratha = 1
                                rstxpay.MoveFirst
                                While rstxpay.EOF = False
                                If Ratha = 1 Then
                                Accnox1 = rstxpay!AccNo
                                AccNamex1 = rstxpay!AccName
                                DebAmt1 = rstxpay!amount
                                ElseIf Ratha = 2 Then
                                Accnox2 = rstxpay!AccNo
                                AccNamex2 = rstxpay!AccName
                                DebAmt2 = rstxpay!amount
                                ElseIf Ratha = 3 Then
                                Accnox3 = rstxpay!AccNo
                                AccNamex3 = rstxpay!AccName
                                DebAmt3 = rstxpay!amount
                                ElseIf Ratha = 4 Then
                                Accnox4 = rstxpay!AccNo
                                AccNamex4 = rstxpay!AccName
                                DebAmt4 = rstxpay!amount
                                ElseIf Ratha = 5 Then
                                Accnox5 = rstxpay!AccNo
                                AccNamex5 = rstxpay!AccName
                                DebAmt5 = rstxpay!amount
                                ElseIf Ratha = 6 Then
                                Accnox6 = rstxpay!AccNo
                                AccNamex6 = rstxpay!AccName
                                DebAmt6 = rstxpay!amount
                                ElseIf Ratha = 7 Then
                                Accnox7 = rstxpay!AccNo
                                AccNamex7 = rstxpay!AccName
                                DebAmt7 = rstxpay!amount
                                ElseIf Ratha = 8 Then
                                Accnox8 = rstxpay!AccNo
                                AccNamex8 = rstxpay!AccName
                                DebAmt8 = rstxpay!amount
                                ElseIf Ratha = 9 Then
                                Accnox9 = rstxpay!AccNo
                                AccNamex9 = rstxpay!AccName
                                DebAmt9 = rstxpay!amount
                                ElseIf Ratha = 10 Then
                                Accnox10 = rstxpay!AccNo
                                AccNamex10 = rstxpay!AccName
                                DebAmt10 = rstxpay!amount
                                End If
                                Ratha = Ratha + 1
                                rstxpay.MoveNext
                                Wend
                                Labam = Val(DebAmt1) + Val(DebAmt2) + Val(DebAmt3) + Val(DebAmt4) + Val(DebAmt5) + Val(DebAmt6) + Val(DebAmt7) + Val(DebAmt8) + Val(DebAmt9) + Val(DebAmt10)
                                End If
                
               If rstxRec.EOF = False Then
                Kajol = 1
                
                rstxRec.MoveFirst
                While rstxRec.EOF = False
                
                If Kajol = 1 Then
                Accnoy1 = rstxRec!AccNo
                AccNamey1 = rstxRec!AccName
                CRAmt1 = rstxRec!amount
                ElseIf Kajol = 2 Then
                Accnoy2 = rstxRec!AccNo
                AccNamey2 = rstxRec!AccName
                CrAmt2 = rstxRec!amount
                ElseIf Kajol = 3 Then
                Accnoy3 = rstxRec!AccNo
                AccNamey3 = rstxRec!AccName
                CrAmt3 = rstxRec!amount
                ElseIf Kajol = 4 Then
                Accnoy4 = rstxRec!AccNo
                AccNamey4 = rstxRec!AccName
                CrAmt4 = rstxRec!amount
                ElseIf Kajol = 5 Then
                Accnoy5 = rstxRec!AccNo
                AccNamey5 = rstxRec!AccName
                CrAmt5 = rstxRec!amount
                ElseIf Kajol = 6 Then
                Accnoy6 = rstxRec!AccNo
                AccNamey6 = rstxRec!AccName
                CrAmt6 = rstxRec!amount
                ElseIf Kajol = 7 Then
                Accnoy7 = rstxRec!AccNo
                AccNamey7 = rstxRec!AccName
                CrAmt7 = rstxRec!amount
                ElseIf Kajol = 8 Then
                Accnoy8 = rstxRec!AccNo
                AccNamey8 = rstxRec!AccName
                CrAmt8 = rstxRec!amount
                ElseIf Kajol = 9 Then
                Accnoy9 = rstxRec!AccNo
                AccNamey9 = rstxRec!AccName
                CrAmt9 = rstxRec!amount
                ElseIf Kajol = 10 Then
                Accnoy10 = rstxRec!AccNo
                AccNamey10 = rstxRec!AccName
                CrAmt10 = rstxRec!amount

                
                End If
                Kajol = Kajol + 1
                
                rstxRec.MoveNext
                Wend
                
                Nattam = Val(CRAmt1) + Val(CrAmt2) + Val(CrAmt3) + Val(CrAmt4) + Val(CrAmt5) + Val(CrAmt6) + Val(CrAmt7) + Val(CrAmt8) + Val(CrAmt9) + Val(CrAmt10)
                End If
 
'-----------------------
maina3 = 1
 Maina2 = 1
            If xRstInv.EOF = False Then
                Maina = 1
                
                xRstInv.MoveFirst
                While xRstInv.EOF = False
                
                If Maina = 1 Then
                invno1 = xRstInv!InvNo
                InvDate1 = xRstInv!InvDate
                DueDate1 = xRstInv!duedate
                ElseIf Maina = 2 Then
                invno2 = xRstInv!InvNo
                InvDate2 = xRstInv!InvDate
                DueDate2 = xRstInv!duedate
                ElseIf Maina = 3 Then
                InvNo3 = xRstInv!InvNo
                InvDate3 = xRstInv!InvDate
                DueDate3 = xRstInv!duedate
                ElseIf Maina = 4 Then
                InvNo4 = xRstInv!InvNo
                InvDate4 = xRstInv!InvDate
                DueDate4 = xRstInv!duedate
                ElseIf Maina = 5 Then
                InvNo5 = xRstInv!InvNo
                InvDate5 = xRstInv!InvDate
                DueDate5 = xRstInv!duedate
                ElseIf Maina = 6 Then
                InvNo6 = xRstInv!InvNo
                InvDate6 = xRstInv!InvDate
                DueDate6 = xRstInv!duedate

                End If
                Maina = Maina + 1
                
                      
                        If Maina2 = 1 Then
                        Poinvno1 = xRstInv!PoNumber
                        PoinvDate1 = xRstInv!PODate
                        ElseIf Maina2 = 2 Then
                        Poinvno2 = xRstInv!PoNumber
                        PoinvDate2 = xRstInv!PODate
                        ElseIf Maina2 = 3 Then
                        Poinvno3 = xRstInv!PoNumber
                        PoinvDate3 = xRstInv!PODate
                        ElseIf Maina2 = 4 Then
                        Poinvno4 = xRstInv!PoNumber
                        PoinvDate4 = xRstInv!PODate
                        ElseIf Maina2 = 5 Then
                        Poinvno5 = xRstInv!PoNumber
                        PoinvDate5 = xRstInv!PODate
                        ElseIf Maina2 = 6 Then
                        Poinvno6 = xRstInv!PoNumber
                        PoinvDate6 = xRstInv!PODate
                        ElseIf Maina2 = 7 Then
                        Poinvno7 = xRstInv!PoNumber
                        PoinvDate7 = xRstInv!PODate
                        ElseIf Maina2 = 8 Then
                        Poinvno8 = xRstInv!PoNumber
                        PoinvDate8 = xRstInv!PODate

                        End If
                      Maina2 = Maina2 + 1
                      
          
        If maina3 = 1 Then
        SEInvNo1 = xRstInv!SENumber
        SEinvdate1 = xRstInv!SEDate
        ElseIf maina3 = 2 Then
        SEInvNo2 = xRstInv!SENumber
        SEinvdate2 = xRstInv!SEDate
        ElseIf maina3 = 3 Then
        SEInvNo3 = xRstInv!SENumber
        SEinvdate3 = xRstInv!SEDate
        ElseIf maina3 = 4 Then
        SEInvNo4 = xRstInv!SENumber
        SEinvdate4 = xRstInv!SEDate
        ElseIf maina3 = 5 Then
        SEInvNo5 = xRstInv!SENumber
        SEinvdate5 = xRstInv!SEDate
        ElseIf maina3 = 6 Then
        SEInvNo6 = xRstInv!SENumber
        SEinvdate6 = xRstInv!SEDate
        ElseIf maina3 = 7 Then
        SEInvNo7 = xRstInv!SENumber
        SEinvdate7 = xRstInv!SEDate
        ElseIf maina3 = 8 Then
        SEInvNo8 = xRstInv!SENumber
        SEinvdate8 = xRstInv!SEDate

      End If
      maina3 = maina3 + 1
                      
                xRstInv.MoveNext
                Wend
                
                 End If

xRstInv.close
rstTerms.close
rstxpay.close
rstxRec.close
rstpay.close
 
'On Error GoTo 0

Dim rstRepo As New ADODB.Recordset
 rstRepo.Open "ReportPayReq", constring, adOpenKeyset, adLockPessimistic, adCmdTable
 
 With rstRepo
 .addnew
  
        rstRepo!SerialNo = PSSno
'        RstRepo!code = ncode
        rstRepo!payto = npayto
        rstRepo!payfor = npayfor
        rstRepo!paymode = npaymode
        rstRepo!vencode = nvencode
        'RstRepo!payee2 = npayee2
        rstRepo!currency = nCurrency
        rstRepo!Payee = nPayee
        rstRepo!DocNo = nDocNo
        rstRepo!branch = nBranch
        rstRepo!Explanation = nExplanation
        rstRepo!ProfCenter = nProfitCenter
        rstRepo!costcenter = nCostCenter
        rstRepo!DateDue = nDateDue
        rstRepo!StoreEntDate = nStoreEntDate
        rstRepo!taxcredit = taxCr
        rstRepo!debitnote = nDebitnote
        rstRepo!creditnote = nCreditnote
        rstRepo!StoreEntNo = nStoreEntNo
        rstRepo!PoNumber = nPoNumber
        rstRepo!invoiceno = nInvoiceno
        rstRepo!invoicedate = Format(nInvoiceDate, "mm/dd/yyyy")
        rstRepo!invAmt = nInvAmt
        rstRepo!Xdate = nXDate
        rstRepo!AmtPaidBefore = nAmtPaidBefore
        rstRepo!amtreqested = nAmtRequested
        rstRepo!outbal = nOutBal
        rstRepo!ProfitTax = nProfitTax
        rstRepo!remarks = nRemarks
        rstRepo!Percentage = nPercentage
        rstRepo!NoOfAttech = nNoOfAttech
        rstRepo!DebitAC = nDebitAC
        rstRepo!AmtDue = nAmtDue
        rstRepo!RefNo = nRefNo
        rstRepo!RemainingBal = nRemainingBal
        rstRepo!PODate = Format(npoDate, "mm/dd/yyyy")
       ' rstRepo!TaxCredit = nTaxCredit
        rstRepo!FcAmount = nPaymentFC
        rstRepo!PaymentLevels = nPaymentLevel
            rstRepo!DescrA = Desc1
            rstRepo!ratea = Rate1
            rstRepo!Daysa = Days1
            rstRepo!xlevelA = xlevel1
            rstRepo!ModeA1 = Mode1
            
            rstRepo!DescrB = Desc2
            rstRepo!rateB = Rate2
            rstRepo!DaysB = Days2
            rstRepo!xlevelB = xlevel2
            rstRepo!ModeB = Mode2
                    
            rstRepo!DescrC = Desc3
            rstRepo!rateC = Rate3
            rstRepo!DaysC = Days3
            rstRepo!xlevelC = xlevel3
            rstRepo!Modec = Mode3

            rstRepo!DescrD = Desc4
            rstRepo!rateD = Rate4
            rstRepo!DaysD = Days4
           ' rstRepo!xlevelD = xlevel4
            rstRepo!ModeD = Mode4
 
            rstRepo!DescrE = Desc5
            rstRepo!rateE = Rate5
            rstRepo!DaysE = Days5
           ' rstRepo!xlevelE = xlevel5
            rstRepo!ModeE = Mode5

rstRepo!AccNoxA = Accnox1
rstRepo!AccNamexA = AccNamex1
rstRepo!DebAmtxA = DebAmt1

rstRepo!AccNoxB = Accnox2
rstRepo!AccNamexB = AccNamex2
rstRepo!DebAmtxB = DebAmt2

rstRepo!AccNoxC = Accnox3
rstRepo!AccNamexC = AccNamex3
rstRepo!DebAmtxC = DebAmt3

rstRepo!AccNoxD = Accnox4
rstRepo!AccNamexD = AccNamex4
rstRepo!DebAmtxD = DebAmt4
 
rstRepo!AccNoxE = Accnox5
rstRepo!AccNamexE = AccNamex5
rstRepo!DebAmtxE = DebAmt5
 
rstRepo!AccNoxF = Accnox6
rstRepo!AccNamexF = AccNamex6
rstRepo!DebAmtxF = DebAmt6
 
rstRepo!AccNoxG = Accnox7
rstRepo!AccNamexG = AccNamex7
rstRepo!DebAmtxG = DebAmt7
 
rstRepo!AccNoxH = Accnox8
rstRepo!AccNamexH = AccNamex8
rstRepo!DebAmtxH = DebAmt8
 
rstRepo!AccNoxI = Accnox9
rstRepo!AccNamexI = AccNamex9
rstRepo!DebAmtxI = DebAmt9
 
rstRepo!AccNoxJ = Accnox10
rstRepo!AccNamexJ = AccNamex10
rstRepo!DebAmtxJ = DebAmt10
 
 
rstRepo!TotalDb = Labam
 
 
rstRepo!AccNoyA = Accnoy1
rstRepo!AccNameyA = AccNamey1
rstRepo!CrAmtxA = CRAmt1

rstRepo!AccNoyB = Accnoy2
rstRepo!AccNameyB = AccNamey2
rstRepo!CrAmtxB = CrAmt2

rstRepo!AccNoyC = Accnoy3
rstRepo!AccNameyC = AccNamey3
rstRepo!CrAmtxC = CrAmt3

rstRepo!AccNoyD = Accnoy4
rstRepo!AccNameyD = AccNamey4
rstRepo!CrAmtxD = CrAmt4

rstRepo!AccNoyE = Accnoy5
rstRepo!AccNameyE = AccNamey5
rstRepo!CrAmtxE = CrAmt5

rstRepo!AccnoyF = Accnoy6
rstRepo!AccNameyF = AccNamey6
rstRepo!CrAmtxF = CrAmt6

rstRepo!AccnoyG = Accnoy7
rstRepo!AccNameyG = AccNamey7
rstRepo!CrAmtxG = CrAmt7

rstRepo!Accnoyh = Accnoy8
rstRepo!AccNameyH = AccNamey8
rstRepo!CrAmtxH = CrAmt8

rstRepo!AccnoyI = Accnoy9
rstRepo!AccNameyI = AccNamey9
rstRepo!CrAmtxI = CrAmt9

rstRepo!AccnoyJ = Accnoy10
rstRepo!AccNameyJ = AccNamey10
rstRepo!CrAmtxJ = CrAmt10

rstRepo!TotalCr = Nattam
 
'                RstRepo!invoiceno = invno1
'                RstRepo!invoicedate = InvDate1
'                RstRepo!DueDate = DueDate1
                
                rstRepo!invoiceno1 = invno1
                rstRepo!invoicedate1 = InvDate1
                rstRepo!DueDate1 = DueDate1
                
                rstRepo!invoiceno2 = invno2
                rstRepo!invoicedate2 = InvDate2
                rstRepo!DueDate2 = DueDate2

                rstRepo!invoiceno3 = InvNo3
                rstRepo!invoicedate3 = InvDate3
                rstRepo!DueDate3 = DueDate3
                
                rstRepo!invoiceno4 = InvNo4
                rstRepo!invoicedate4 = InvDate4
                rstRepo!DueDate4 = DueDate4
                
                rstRepo!invoiceno5 = InvNo5
                rstRepo!invoicedate5 = InvDate5
                rstRepo!DueDate5 = DueDate5
                
                rstRepo!invoiceno6 = InvNo6
                rstRepo!invoicedate6 = InvDate6
                rstRepo!DueDate6 = DueDate6
                
                rstRepo!JOurnalNo = JKL
            
                
                
                
                
rstRepo!POnum1 = Poinvno1
rstRepo!POnum2 = Poinvno2
rstRepo!POnum3 = Poinvno3
rstRepo!POnum4 = Poinvno4
rstRepo!POnum5 = Poinvno5
rstRepo!POnum6 = Poinvno6
rstRepo!POnum7 = Poinvno7
rstRepo!POnum8 = Poinvno8

rstRepo!PODate1 = PoinvDate1
rstRepo!PODate2 = PoinvDate2
rstRepo!PODate3 = PoinvDate3
rstRepo!PODate4 = PoinvDate4
rstRepo!PODate5 = PoinvDate5
rstRepo!PODate6 = PoinvDate6
rstRepo!PODate7 = PoinvDate7
rstRepo!PODate8 = PoinvDate8

rstRepo!SEnum1 = SEInvNo1
rstRepo!SEnum2 = SEInvNo2
rstRepo!SEnum3 = SEInvNo3
rstRepo!SEnum4 = SEInvNo4
rstRepo!SEnum5 = SEInvNo5
rstRepo!SEnum6 = SEInvNo6
rstRepo!SEnum7 = SEInvNo7
rstRepo!SEnum8 = SEInvNo8


rstRepo!SEdate1 = SEinvdate1
rstRepo!SEdate2 = SEinvdate2
rstRepo!SEdate3 = SEinvdate3
rstRepo!SEdate4 = SEinvdate4
rstRepo!SEdate5 = SEinvdate5
rstRepo!SEdate6 = SEinvdate6
rstRepo!SEdate7 = SEinvdate7
rstRepo!SEdate8 = SEinvdate8
                
                
                
 '.Update
 rstRepo.Update
 End With
On Error Resume Next

procePrintBy1 PaymentRequestLast.Sections(3).Controls("lblClog")

DataEnvironment1.rsReportPayRequest.close
DataEnvironment1.rsReportPayRequest.Requery
PaymentRequestLast.Show 1
On Error GoTo 0


'-------------------------------------
'THIS IS FOR THE DERECT PRINT
uifo = MsgBox("Did you print it?", vbYesNo + vbQuestion, "Please confirm")
If uifo = vbYes Then
   
Dim rstconPr As New ADODB.Recordset
rstconPr.Open "Select * from Payablesetup where serialno= " & "'" & PSSno & "'" & "", conx, adOpenDynamic, adLockOptimistic, adCmdText

  
rstconPr!printmark = "Printed"
rstconPr.Update
End If
'---------------------------------
cmdExit1_Click




Exit Sub
 
 
 rstpay.MoveNext
 Wend
End If
End If


End Sub

Private Sub cmdSearch_Click()
ListView1.BorderStyle = ccFixedSingle
ListView1.MultiSelect = False

   ' FindItem method returns a reference to the found item, so
   ' you must create an object variable and set the found item
   ' to it.
   Dim itmFound As ListItem   ' FoundItem variable.
   
   strFindMe = txtSearch1.Text

   Set itmFound = ListView1.Finditem(strFindMe, intSelectedOption, , lvwText)
   
   ' If no ListItem is found, then inform user and exit. If a
   ' ListItem is found, scroll the control using the EnsureVisible
   ' method, and select the ListItem.
   If itmFound Is Nothing Then  ' If no match, inform user and exit.
      MsgBox "No match found"
      Exit Sub
   Else
       itmFound.EnsureVisible ' Scroll ListView to show found ListItem.
       itmFound.Selected = True   ' Select the ListItem.
      ' Return focus to the control to see selection.
       ListView1.SetFocus
   End If
ListView1.MultiSelect = True

End Sub
Private Sub procePrintBy1(PrBy1 As RptLabel)

PrBy1.caption = cLogUser
End Sub


Private Sub cmdSearch2_Click()

ListView4.BorderStyle = ccFixedSingle
ListView4.MultiSelect = False

   ' FindItem method returns a reference to the found item, so
   ' you must create an object variable and set the found item
   ' to it.
   Dim itmFound As ListItem   ' FoundItem variable.
   
   strFindMe = txtSearch2.Text

   Set itmFound = ListView4.Finditem(strFindMe, intSelectedOption, , lvwText)
   
   ' If no ListItem is found, then inform user and exit. If a
   ' ListItem is found, scroll the control using the EnsureVisible
   ' method, and select the ListItem.
   If itmFound Is Nothing Then  ' If no match, inform user and exit.
      MsgBox "No match found"
      Exit Sub
   Else
       itmFound.EnsureVisible ' Scroll ListView to show found ListItem.
       itmFound.Selected = True   ' Select the ListItem.
      ' Return focus to the control to see selection.
       ListView4.SetFocus
   End If
ListView4.MultiSelect = True

End Sub

Private Sub cmdSearch3_Click()
ListView5.BorderStyle = ccFixedSingle
ListView5.MultiSelect = False

   ' FindItem method returns a reference to the found item, so
   ' you must create an object variable and set the found item
   ' to it.
   Dim itmFound As ListItem   ' FoundItem variable.
   
   strFindMe = txtSearch2.Text

   Set itmFound = ListView5.Finditem(strFindMe, intSelectedOption, , lvwText)
   
   ' If no ListItem is found, then inform user and exit. If a
   ' ListItem is found, scroll the control using the EnsureVisible
   ' method, and select the ListItem.
   If itmFound Is Nothing Then  ' If no match, inform user and exit.
      MsgBox "No match found"
      Exit Sub
   Else
       itmFound.EnsureVisible ' Scroll ListView to show found ListItem.
       itmFound.Selected = True   ' Select the ListItem.
      ' Return focus to the control to see selection.
       ListView5.SetFocus
   End If
ListView5.MultiSelect = True

End Sub

'Private Sub Combo5_Click()
' strVen = Trim(Combo5.Text)
'
'rstrate.MoveFirst
'While rstrate.EOF = False
'If strVen = Trim(rstrate!venname) Then
'      txtvencode.Text = Trim(rstrate!vendorcode)
'     Set mitem = Me.ListView2.ListItems.Add(, , Trim(rstrate!rate))
'     mitem.SubItems(1) = Trim(rstrate!Description)
'     mitem.SubItems(2) = Trim(rstrate!days)
'     mitem.SubItems(3) = Trim(rstrate!Mode)
'     mitem.SubItems(4) = Trim(rstrate!Levels)
'End If
'rstrate.MoveNext
'Wend

''This is to add the details to the list box
'rstVen.MoveFirst
'While rstVen.EOF = False
'If strVen = Trim(rstVen!vennameeng) Then
'On Error Resume Next
'txtvenCode.Text = rstVen!vencode
'List1.AddItem rstVen!venhomeper
'List1.AddItem rstVen!vencorptel
'List1.AddItem rstVen!venhomecty
'End If
'rstVen.MoveNext
'Wend

'End Sub


'Private Sub Combo5_GotFocus()
'Const CB_SHOWDROPDOWN = &H14F
'   Dim Tmp
'   Tmp = SendMessage(Combo5.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
'
'End Sub

'Private Sub Combo5_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyPageUp Or KeyCode = vbKeyUp Then
'cmbPaymentFor.SetFocus
'End If
'
'If KeyCode = vbKeyPageDown Or KeyCode = vbKeyDown Then
'txtDocuNo.SetFocus
'End If
'
'End Sub

'Private Sub Combo5_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'cmbPayee2.SetFocus
'End If
'
'End Sub

Public Sub PayableCancellation()
Dim CON1 As New ADODB.Connection
Dim rstpa As New ADODB.Recordset
conStr = "Provider=MSDASQL;DSN=Finance;UID=sa; PWD=;"
CON1.Open conStr

Unload frmPassword
Dim BufSerNo As Long
GotYaa = 0
'This is the main Program to cancel the the Payable Setup
rstpa.Open "Select * from Payablesetup where deletemark = '0' ", CON1, adOpenDynamic, adLockOptimistic
'this is to EDIT XPayment(Inside List3)
        BufSerNo = FrmPayableSetup.ListView1.SelectedItem.Text
If rstpa.EOF = False Then
rstpa.MoveFirst
End If
While rstpa.EOF = False
If BufSerNo = rstpa!SerialNo Then
rstpa!cancelledmark = 1
GotYaa = 1
End If
rstpa.Update
rstpa.MoveNext
Wend
rstpa.close

If GotYaa = 1 Then
MsgBox "Yor record was Cancelled Successfully", vbInformation, "confirmation"
End If


Dim Soothu
'This is to REfresh ListView 1
Dim Pay222 As New ADODB.Recordset
Soothu = "Select * from Payablesetup where cancelledmark = '0' and deletemark = '0' and confirmedmark = '0' "
Pay222.Open Soothu, CON1, adOpenDynamic, adLockOptimistic

'this is to add Listview 1
Me.ListView1.ListItems.clear

        If Pay222.EOF = False Then
        Pay222.MoveFirst
        End If


  While Pay222.EOF = False
   'If Trim(Pay222!cancelledmark) = 0 And Trim(Pay222!ConfirmedMark) = 0 And Trim(Pay222!Paidmark) = 0 Then    'This is not Cancelled
  


  
     Set MItem = Me.ListView1.ListItems.Add(, , Format(Pay222!SerialNo))
     MItem.SubItems(1) = Format(Pay222!Xdate, "dd/mm/yyyy")
     MItem.SubItems(2) = Format(Pay222!Requester)
     MItem.SubItems(3) = Format(Pay222!DateDue, "dd/mm/yyyy")
     MItem.SubItems(4) = Format(Pay222!RefNo)
     MItem.SubItems(5) = Format(Pay222!journaledmark)
     MItem.SubItems(6) = Format(Pay222!amtreqested, "#############.#0")
        
        
     On Error Resume Next
       Totlist1 = Val(Totlist1) + Val(Trim(Pay222!TotCrAmt)) 'This is for the Total of the List
     On Error GoTo 0
  '  End If
     Pay222.MoveNext
     Wend
 Pay222.close

txtTotList1.Text = Totlist1

Sunnie = "Select * from Payablesetup where cancelledmark = '1' and deletemark = '0' and confirmedmark = '0' "
Pay222.Open Sunnie, CON1, adOpenDynamic, adLockOptimistic

 
'This is  Refresh ListView4
Me.ListView4.ListItems.clear
If Pay222.EOF = False Then
 Pay222.MoveFirst
End If
  While Pay222.EOF = False
    'If Trim(Pay222!cancelledmark) = True Then

     Set MItem = Me.ListView4.ListItems.Add(, , Format(Pay222!SerialNo))
     MItem.SubItems(1) = Format(Pay222!Xdate, "dd/mm/yyyy")
     MItem.SubItems(2) = Format(Pay222!Payee)
     MItem.SubItems(3) = Format(Pay222!DocNo)
     MItem.SubItems(4) = Format(Pay222!amtreqested, "#############.#0")
     
    ' End If
   Totlist4 = Val(Totlist4) + Val(Trim(Pay222!outbal)) 'This is for the Total of the List
    
     Pay222.MoveNext
     Wend
Me.txtTotCanceleldBal = Trim(Totlist4)

 Pay222.close



End Sub

Private Sub Command1_Click()
cmdNew_Click
SSTab1.SetFocus
SendKeys "{Left}"
End Sub


Private Sub Command11_Click()
  FrmPayableSetup.SSTab1.SetFocus
  SendKeys "{right}"

End Sub

Private Sub Command12_Click()
  FrmPayableSetup.SSTab1.SetFocus
  SendKeys "{Left}"

End Sub

Private Sub Command2_Click()
If Me.cmdNew.caption = "&New" Then
MsgBox "You have to do the Payable setup to do the Payement Analysis", vbInformation, "Please Do the Payable Setup"
Exit Sub
End If

''''''''''''''''''''''
If Me.CmdEdit.caption = "&Update" Then
Dim ZXC
ZXC = Me.txtserialNo.Text
Dim rstList3XDirect As New ADODB.Recordset
Dim VarrstList3X1Direct
VarrstList3X1Direct = "Select * from xpayment where SerialNo = " & "'" & ZXC & "'" & ""
 rstList3XDirect.Open VarrstList3X1Direct, constring, adOpenDynamic, adLockOptimistic
 
        On Error Resume Next
        Dim stuDirect
        stuDirect = 0
        If rstList3XDirect.EOF = False Then
        rstList3XDirect.MoveFirst
        End If
        While rstList3XDirect.EOF = False
       ' If Trim(rstList3XDirect!SerialNo) = Trim(xvar) Then
        Set MItem = FrmPaymentAnalysis.ListView3.ListItems.Add(, , Trim(rstList3XDirect!AccNo))
        MItem.SubItems(1) = IIf(IsNull(rstList3XDirect!AccName) = True, "", Trim(rstList3XDirect!AccName))
        MItem.SubItems(2) = Trim(rstList3XDirect!amount)
        stuDirect = Val(stuDirect) + Val(rstList3XDirect!amount)
       ' End If
        rstList3XDirect.MoveNext
         Wend
rstList3XDirect.close




 'This is to Edit the XReceipt and Add Details to List3 Details From ListView1(Main)
 
 
 
Dim rstReceiptX1Direct As New ADODB.Recordset
Dim VarrstReceiptX1Direct
VarrstReceiptX1Direct = "Select * from xReceipt where SerialNo = " & "'" & ZXC & "'" & ""
rstReceiptX1Direct.Open VarrstReceiptX1Direct, constring, adOpenDynamic, adLockOptimistic
 
 
        Dim ijkDirect
        ijkDirect = 0
        If rstReceiptX1Direct.EOF = False Then
        rstReceiptX1Direct.MoveFirst
        End If
        While rstReceiptX1Direct.EOF = False
 '       If Trim(rstReceipt!SerialNo) = Trim(xvar) Then
        Set MItem = FrmPaymentAnalysis.ListView6.ListItems.Add(, , Trim(rstReceiptX1Direct!AccNo))
        MItem.SubItems(1) = Trim(rstReceiptX1Direct!AccName)
        MItem.SubItems(2) = Trim(rstReceiptX1Direct!amount)
        ijkDirect = Val(ijkDirect) + Val(rstReceiptX1Direct!amount)
    '    End If
        rstReceiptX1Direct.MoveNext
         Wend
         
FrmPaymentAnalysis.txtTotList6 = ijkDirect
FrmPaymentAnalysis.txtTotList3 = stuDirect

FrmPaymentAnalysis.txtAccNo.Enabled = True
FrmPaymentAnalysis.txtPartic.Enabled = True
FrmPaymentAnalysis.txtAmo.Enabled = True

FrmPaymentAnalysis.txtDBAccNo.Enabled = True
FrmPaymentAnalysis.txtDBPartic.Enabled = True
FrmPaymentAnalysis.txtDBAmo.Enabled = True

FrmPayableSetup.txtTMDays.Visible = True
FrmPayableSetup.txtTMDes.Visible = True
FrmPayableSetup.txtTMlevel.Visible = True
FrmPayableSetup.txtTMmode.Visible = True
FrmPayableSetup.txtTmRate.Visible = True
'Me.sEdit.Enabled = False
FrmPayableSetup.cmdNew.caption = "&Update"
FrmPayableSetup.CmdEdit.caption = "&Update"

FrmPaymentAnalysis.CmdEdit.caption = "&Update"
FrmPaymentAnalysis.cmdNew.caption = "&Update"
FrmPaymentAnalysis.cmdNew.Visible = False
FrmPayableSetup.cmdExit1.caption = "&Cancel"
frmMenu.shedit.Enabled = False
FrmPayableSetup.txtserialNo.Enabled = False


End If
'----------------------------------------



If Me.cmdNew.caption = "&Save" And Me.cmdPrintPayReq.Visible = False Then
    Dave = MsgBox("You have to Save the Payable setup to do the Payement Analysis, Do You  want to Save it ?", vbInformation + vbYesNo, "Confirmation")
    If Dave = vbNo Then
    Exit Sub
    Else

        If txtAmtReq.Text = "" Then
        MsgBox "You have to Fill the Text 'AmountRequested'", vbInformation, "Couldn't Shift to Payment Request"
        Exit Sub
        End If
 
        If Me.cmbPaymentFor = "" Then
        MsgBox "You have to Fill the Combo 'PaymentFor' ", vbInformation, "Couldn't Shift to Payment Request"
        Exit Sub
        End If
         
        If cmbPaymentLevel = "" Then
        MsgBox "You have to Fill the Combo 'Payment Level' ", vbInformation, "Couldn't Shift to Payment Request"
        Exit Sub
        End If
 
         
        If Me.cmbCostCenter.Text = "" Then
        MsgBox "You have to Fill the Combo  'CostCenter'", vbInformation, "Couldn't Shift to Payment Request"
        Exit Sub
        End If
 
        If Me.CmbPrepBy = "" Then
        MsgBox "You have to Fill the Combo  'PrepBy'", vbInformation, "Couldn't Shift to Payment Request"
        Exit Sub
        End If
    
        If mskDateDue = "__/__/____" Then
        MsgBox "You have to Fill the Box DateDue", vbInformation, "Couldn't Shift to Payment Request"
        Exit Sub
        End If
    
        If Me.txtInvAmt = "" Then
        MsgBox "Invoice Amount is Empty , You have No Rights to Change or Delete this Text, It should Automatically be Filled, So  First go to the Purchase Setup and Do the Transactions to Enter the Invoice Details", vbExclamation, "Operation Cancelled"
        'Here i can Validate the Invoice Amount in the Future
        Exit Sub
        End If


    End If

'Here I call the Proc to Save the Payable Setup
Call saveme


End If

        If txtAmtReq.Text = "" Then
        MsgBox "You have to Fill the Text 'AmountRequested'", vbInformation, "Couldn't Shift to Payment Request"
        Exit Sub
        End If


FrmPaymentAnalysis.txtAmtReq.Text = txtAmtReq.Text
FrmPaymentAnalysis.Text1.Text = txtInvAmt.Text
FrmPaymentAnalysis.Text2.Text = txtAmtPaid.Text
FrmPaymentAnalysis.Text3.Text = txtOutBal.Text
FrmPaymentAnalysis.txtPaymentLevel.Text = cmbPaymentLevel.Text

'FrmPaymentAnalysis.txtTotList3.Text = txtAmtreq.Text
FrmPaymentAnalysis.txtIdentifyNewData.Text = "New"
SNom = Me.txtserialNo
MskDa = Me.MskDate


FrmPaymentAnalysis.txtserialNo = SNom
FrmPaymentAnalysis.MaskEdBox1 = MskDa


FrmPaymentAnalysis.Show

End Sub
Private Sub SaveMeBForAnalisis()
FrmPayableSetup.txtIdentificationForPaymentAnalisis = "FlagSelected"







End Sub

Private Sub Command3_Click()


If Me.txtserialNo.Text = "" Then
MsgBox "YOu are Trying to Click the Button without any Reason", vbInformation, "Click New to Start Payable setup"
Exit Sub
End If



If Me.cmdPrintPayReq.Visible = True Then
Me.cmdPrintPayReq.Visible = False
cmdNew.Visible = True
cmdNew.caption = "&Update"
End If



'THIS IS TO DELETE THE RELATED RECORDS IN THE  TEMP TABLE
'If they Go back to the table the existing records will be Repeated
'So this is the Idea to do it

Dim Thisx As Long



Dim xItemFOrDelete As Long
xItemFOrDelete = Me.txtserialNo.Text
       ' FrmPayableSetup.ListView1.ListItems.Remove xItemFOrDel
Dim rstDelTempPayInvoiceDetails As New ADODB.Recordset
rstDelTempPayInvoiceDetails.Open "delete from PayTempInvoiceDetails where serialno = " & "'" & xItemFOrDelete & "'" & "", constring, adOpenDynamic, adLockOptimistic




'On Error Resume Next
frmNewPurchse.Show 1
On Error GoTo 0
End Sub

Private Sub Command4_Click()
ListView8.BorderStyle = ccFixedSingle
ListView8.MultiSelect = False
   Dim itmFound As ListItem
   strFindMe = txtSearch8.Text
   Set itmFound = ListView8.Finditem(strFindMe, intSelectedOption, , lvwText)
   If itmFound Is Nothing Then
      MsgBox "No match found"
      Exit Sub
   Else
       itmFound.EnsureVisible
       itmFound.Selected = True
       ListView8.SetFocus
   End If
ListView8.MultiSelect = True
End Sub

Private Sub Command5_Click()
  FrmPayableSetup.SSTab1.SetFocus
  SendKeys "{left}"

If SSTab1.caption = "Payable Entry" Then
SSTab1.ToolTipText = "Step1"
ElseIf SSTab1.caption = "List View" Then
SSTab1.ToolTipText = "Step2 & Step4"
ElseIf SSTab1.caption = "Cancelled List" Then
SSTab1.ToolTipText = "Step3"
ElseIf SSTab1.caption = "Paid Voucher" Then
SSTab1.ToolTipText = "Step5 / Step4"

ElseIf SSTab1.caption = "Journal" Then
SSTab1.ToolTipText = "Step6"
End If


End Sub

Private Sub Command6_Click()
  FrmPayableSetup.SSTab1.SetFocus
  SendKeys "{right}"
If SSTab1.caption = "Payable Entry" Then
SSTab1.ToolTipText = "Step1"
ElseIf SSTab1.caption = "List View" Then
SSTab1.ToolTipText = "Step2 & Step4"
ElseIf SSTab1.caption = "Cancelled List" Then
SSTab1.ToolTipText = "Step3"
ElseIf SSTab1.caption = "Paid Voucher" Then
SSTab1.ToolTipText = "Step5 /Step 4"

ElseIf SSTab1.caption = "Journal" Then
SSTab1.ToolTipText = "Step6"
End If


End Sub



Private Sub Form_Load()


Dim CON1 As ADODB.Connection
Set CON1 = New ADODB.Connection
Set rstJournal = New ADODB.Recordset
Set rstrate = New ADODB.Recordset
Set rstPAyFor = New ADODB.Recordset
Set rstEmp = New ADODB.Recordset
Set rstSource = New ADODB.Recordset
Set rstVen = New ADODB.Recordset
Set rstCosPro = New ADODB.Recordset
Set rstPaySetup = New ADODB.Recordset
Set rstterm = New ADODB.Recordset
Set rstPayment = New ADODB.Recordset
Set rstReceipt = New ADODB.Recordset
Set rstChart = New ADODB.Recordset
Set RstPaid = New ADODB.Recordset

Dim MYdep As New ADODB.Recordset

conStr = "Provider=MSDASQL;DSN=Finance;UID=sa; PWD=;"

Set xcol = Me.ListView2.ColumnHeaders.Add(, , "Rate ÇáÇíÌÇÑ", 800)
Set xcol = Me.ListView2.ColumnHeaders.Add(, , "Description ÇáæÕÝ  ", 2050)
Set xcol = Me.ListView2.ColumnHeaders.Add(, , "Days ÇáÇíÇã ", 800)
Set xcol = Me.ListView2.ColumnHeaders.Add(, , "Mode ÇáäæÚ  ", 850)
Set xcol = Me.ListView2.ColumnHeaders.Add(, , "Level ÇáãÓÊæí ", 850)


Set xcol = Me.ListView1.ColumnHeaders.Add(, , "Number ÇáÑÞã ", 1200)
Set xcol = Me.ListView1.ColumnHeaders.Add(, , "Date ÇáÊÇÑíÎ  ", 1200)
Set xcol = Me.ListView1.ColumnHeaders.Add(, , "Payee ÇáÏÝÚ ", 3000)
Set xcol = Me.ListView1.ColumnHeaders.Add(, , "Due Date íæã ÇáãÓÊÍÞ ", 1200)
Set xcol = Me.ListView1.ColumnHeaders.Add(, , "RequesterÇáãØáæÈ ", 1500)
Set xcol = Me.ListView1.ColumnHeaders.Add(, , "Status", 900)
Set xcol = Me.ListView1.ColumnHeaders.Add(, , "Journalised", 1000)
Set xcol = Me.ListView1.ColumnHeaders.Add(, , "Amount ÇáÇÌãÇáí ", 1100, lvwColumnRight)

Me.ListView1.ColumnHeaders(5).Alignment = lvwColumnRight

Set xcol = Me.ListView7.ColumnHeaders.Add(, , "Number ÇáÑÞã ", 1200)
Set xcol = Me.ListView7.ColumnHeaders.Add(, , "Date ÇáÊÇÑíÎ  ", 1300)
Set xcol = Me.ListView7.ColumnHeaders.Add(, , "Supplier ÇáãæÑÏ ", 4300)
Set xcol = Me.ListView7.ColumnHeaders.Add(, , "Due Date íæã áãÓÊÍÞ ", 1400)
Set xcol = Me.ListView7.ColumnHeaders.Add(, , "Requester ÇáãØáæÈ", 1500)
Set xcol = Me.ListView7.ColumnHeaders.Add(, , "Amount ÇáÇÌãÇáí ", 1400, lvwColumnRight)

Me.ListView7.ColumnHeaders(5).Alignment = lvwColumnRight


Set xcol = Me.ListView4.ColumnHeaders.Add(, , "Number ÇÑÞã ", 1400)
Set xcol = Me.ListView4.ColumnHeaders.Add(, , "Date ÇáÊÇÑíÎ ", 1300)
Set xcol = Me.ListView4.ColumnHeaders.Add(, , "Supplier ÇáãæÑÏ ", 4000)
Set xcol = Me.ListView4.ColumnHeaders.Add(, , "Account ÇáÍÓÇÈ ", 2600)
Set xcol = Me.ListView4.ColumnHeaders.Add(, , "Amount ÇáÇÌãÇáí ", 1800, lvwColumnRight)

Set xcol = Me.ListView5.ColumnHeaders.Add(, , "Number ÇáÑÞã ", 1400)
Set xcol = Me.ListView5.ColumnHeaders.Add(, , "Date ÇáÊÇÑíÎ ", 1300)
Set xcol = Me.ListView5.ColumnHeaders.Add(, , "Supplier ÇáãæÑÏ ", 3900)
Set xcol = Me.ListView5.ColumnHeaders.Add(, , "Requester ÇáãØáæÈ", 2300)
Set xcol = Me.ListView5.ColumnHeaders.Add(, , "Amount ÇáÇÌãÇáí ", 1400, lvwColumnRight)

Set xcol = Me.ListView8.ColumnHeaders.Add(, , "JounalNo ÑÞã ÇáíæãíÉ ", 1400)
Set xcol = Me.ListView8.ColumnHeaders.Add(, , "Ticket No ÑÞã ÇáÞíÏ ", 1300)
Set xcol = Me.ListView8.ColumnHeaders.Add(, , "Date Confirmed ÊÇÑíÎ ÇáÊÇßíÏ ", 1300)
Set xcol = Me.ListView8.ColumnHeaders.Add(, , "AccountcodeßæÏ ÇáÍÓÇÈ ", 1300)
Set xcol = Me.ListView8.ColumnHeaders.Add(, , "AccountName ÇÓã ÇáÍÓÇÈ ", 3000)
Set xcol = Me.ListView8.ColumnHeaders.Add(, , "Debit ÇáãÏíä ", 1400, lvwColumnRight)
Set xcol = Me.ListView8.ColumnHeaders.Add(, , "Credit ÇáÏÇÆä ", 1400, lvwColumnRight)

txtTMDays.Visible = False
txtTMDes.Visible = False
txtTMlevel.Visible = False
txtTMmode.Visible = False
txtTmRate.Visible = False
Dim Apple

CON1.Open conStr
rstEmp.Open "Select * from Newemployee", CON1, adOpenDynamic, adLockOptimistic
rstSource.Open "Select * from NewForSource", CON1, adOpenDynamic, adLockOptimistic
rstVen.Open "Select * from vendor", CON1, adOpenDynamic, adLockOptimistic
rstPAyFor.Open "Select * from newPaymentFor", CON1, adOpenDynamic, adLockOptimistic
rstrate.Open "Select * from Newrate", CON1, adOpenDynamic, adLockOptimistic
rstCosPro.Open "Select * from CosProCenters", CON1, adOpenDynamic, adLockOptimistic
Apple = "Select * from PayableSetup where deletemark = '0'  and Post = 'No' order by SerialNo"

rstPaySetup.Open Apple, CON1, adOpenDynamic, adLockOptimistic
rstPayment.Open "Select * from xPayment", CON1, adOpenDynamic, adLockOptimistic
rstterm.Open "Select * from term", CON1, adOpenDynamic, adLockOptimistic
rstReceipt.Open "Select * from xReceipt", CON1, adOpenDynamic, adLockOptimistic
'rstReceipt.Open "Select * from xPayment where cramount <> 0", con1, adOpenDynamic, adLockOptimistic
rstChart.Open "Select * from level6", CON1, adOpenDynamic, adLockOptimistic
RstPaid.Open "Select * from payablesetup where Paidmark = 1", CON1, adOpenDynamic, adLockOptimistic
MYdep.Open "Select * from Department", CON1, adOpenDynamic, adLockOptimistic


Dim kkk
kkk = Unposted
xs = "SELECT * From PayJournal where status='unposted' order by serialno"
rstJournal.Open xs, CON1, adOpenDynamic, adLockOptimistic, adCmdText

cmbPayment.AddItem "Local Currncy"
cmbPayment.AddItem "Forign Currency"

cmbPaymentLevel.AddItem "Accrued Liabilties"
cmbPaymentLevel.AddItem "Full"
cmbPaymentLevel.AddItem "First"
cmbPaymentLevel.AddItem "Second"
cmbPaymentLevel.AddItem "Third"
cmbPaymentLevel.AddItem "Fourth"
cmbPaymentLevel.AddItem "Final"

txtBranch.AddItem "Obuor ÚÈæÑ"
txtBranch.AddItem "Arcadia ÇÑßÇÏíÇ "

cmbPaymode.AddItem "Cash ÇáäÞÏí"
cmbPaymode.AddItem "Check ÇáÔíß "
cmbPaymode.AddItem "Bank Transfer "

txtTMmode.AddItem "Cash ÇáäÞÏí "
txtTMmode.AddItem "Check ÇáÔíß "

txtTMDes.AddItem "On Delevary ÚäÏ ÇáÊÓáíã "
txtTMDes.AddItem "After Delevary ÈÚÏ ÇáÊÓáíã "
txtTMDes.AddItem "When Shipping ÚäÏãÇ ÇáÊÓæíÞ "
txtTMDes.AddItem "Invoice Time æÞÊ ÇáÝÇÊæÑÉ "
txtTMDes.AddItem "Payment With P.O"

'-0---0--0-0-0-0

Dim recpayee As New ADODB.Recordset
recpayee.Open "select * from financemaster where substring(accountcode,1,3) = '131'", CON1, adOpenKeyset, adLockOptimistic
While recpayee.EOF = False
If Trim(recpayee!accountnamearab) = "" Then
    anu = recpayee!accountnameeng
Else
    anu = recpayee!accountnamearab & "\" & recpayee!accountnameeng
End If
    cmbPayee2.AddItem anu
    recpayee.MoveNext
Wend
recpayee.close


recpayee.Open "payee", CON1, adOpenKeyset, adLockOptimistic
While recpayee.EOF = False
anu = recpayee!payeecode & "     " & recpayee!payeenameeng
    cmbPayee2.AddItem anu
    recpayee.MoveNext
Wend
recpayee.close
'end received from


'-0---0-0-0-0-
MYdep.MoveFirst
While MYdep.EOF = False
    cmbCostCenter.AddItem MYdep!Denameara & "     " & MYdep!DenameEng
MYdep.MoveNext
Wend
MYdep.close


txtEnter1.Enabled = False
TxtEnter2.Enabled = False
' mskInvDate.Enabled = False
 txtCreditNote.Enabled = False
 txtDebitNote.Enabled = False
 txtOutBal.Enabled = False
 'txtPoNo.Enabled = False
' txtProTax.Enabled = False
' txtPercentage.Enabled = False
txtAmtPaid.Enabled = False
cmbPayee2.Enabled = False
txtAmtReq.Enabled = False

txtFCamount.Enabled = False
txtAmountDue.Enabled = False
'txtApprBy.Enabled = False
txtBranch.Enabled = False
txtDocuNo.Enabled = False
mskDateDue.Enabled = False
txtRefNo.Enabled = False
CmbApprovedBy.Enabled = False
cmbCostCenter.Enabled = False
CmbNotedBy.Enabled = False
CmbSource.Enabled = False
 'Combo5.Enabled = False
cmbPaymentFor.Enabled = False
cmbProfCenter.Enabled = False

MskDate.Enabled = False
List1.Enabled = False
'txtStorEntryNo.Enabled = False
'mskStoreEnDate.Enabled = False
'txtinvNo.Enabled = False
Option1.Enabled = False
Option2.Enabled = False

If rstPaySetup.BOF Then
On Error Resume Next

End If


'this is to add Listview 1
If rstPaySetup.EOF = False Then
 rstPaySetup.MoveFirst
End If
  
  While rstPaySetup.EOF = False
  If Trim(rstPaySetup!cancelledmark) = 0 And Trim(rstPaySetup!ConfirmedMark) = 0 And Trim(rstPaySetup!Paidmark) = 0 Then    'This is not Cancelled
     Set MItem = Me.ListView1.ListItems.Add(, , Format(rstPaySetup!SerialNo))
     MItem.SubItems(1) = Format(rstPaySetup!Xdate, "dd/mm/yyyy")
     MItem.SubItems(2) = Format(rstPaySetup!Requester)
     MItem.SubItems(3) = Format(rstPaySetup!DateDue, "dd/mm/yyyy")
     MItem.SubItems(4) = Format(rstPaySetup!RefNo)
     MItem.SubItems(5) = Format(rstPaySetup!printmark)
     MItem.SubItems(6) = Format(rstPaySetup!journaledmark)
     MItem.SubItems(7) = Format(rstPaySetup!amtreqested, "#############.#0")
        
     On Error Resume Next
       Totlist1 = Val(Totlist1) + Val(Trim(rstPaySetup!amtreqested)) 'This is for the Total of the List
     On Error GoTo 0
 End If
     rstPaySetup.MoveNext
     Wend


txtTotList1.Text = Totlist1


On Error GoTo 0

If RstPaid.BOF Then
On Error Resume Next
End If

'this is to add Listview 5
If RstPaid.EOF = False Then
 RstPaid.MoveFirst
End If


  While RstPaid.EOF = False
     Set MItem = Me.ListView5.ListItems.Add(, , Format(RstPaid!SerialNo))
     MItem.SubItems(1) = Format(RstPaid!Xdate, "dd/mm/yyyy")
     MItem.SubItems(2) = Format(RstPaid!payto)
     MItem.SubItems(3) = Format(RstPaid!RefNo)
     MItem.SubItems(4) = Format(RstPaid!amtreqested, "#############.#0")
     totlist5 = Val(totlist5) + IIf(IsNull(RstPaid!amtreqested) = True, "0", Val(Trim(RstPaid!amtreqested))) 'This is for the Total of the List
 
     RstPaid.MoveNext
     Wend
  txtTotVoch.Text = totlist5


If rstJournal.EOF <> True Then
txtTotList1.Text = Trim(Totlist1)

On Error GoTo 0

If rstJournal.BOF Then
On Error Resume Next
End If




'this is to add Listview8
If rstJournal.EOF = False Then
 rstJournal.MoveFirst
End If
  While rstJournal.EOF = False
   If rstJournal!cancelledmark = "0" Then
     Set MItem = Me.ListView8.ListItems.Add(, , Format(rstJournal!SerialNo))
     MItem.SubItems(1) = Format(rstJournal!ticket)
     MItem.SubItems(2) = Format(rstJournal!confirmeddate, "dd/mm/yyyy")
     MItem.SubItems(3) = Format(rstJournal!AccNo)
     MItem.SubItems(4) = Format(rstJournal!AccName)
     MItem.SubItems(5) = Format(rstJournal!DBamount, "#############.#0")
     MItem.SubItems(6) = Format(rstJournal!CRamount, "#############.#0")
       TotList8Db = Val(TotList8Db) + Val(Trim(rstJournal!DBamount)) 'This is for the Total of the List
       totlist8cr = Val(totlist8cr) + Val(Trim(rstJournal!CRamount)) 'This is for the Total of the List
    End If
     rstJournal.MoveNext
     Wend
txtJdb.Text = Trim(TotList8Db)
txtJCr.Text = Trim(totlist8cr)
Else
End If

On Error GoTo 0

If rstPaySetup.BOF Then
On Error Resume Next
End If


'this is to add Listview7

 rstPaySetup.MoveFirst

  While rstPaySetup.EOF = False
  Dim Yaas
  Yaas = "Yes"
  If Trim(rstPaySetup!cancelledmark) = False And Trim(rstPaySetup!deletemark) = False And Trim(rstPaySetup!ConfirmedMark) = True And Trim(rstPaySetup!Paidmark) = False And Trim(rstPaySetup!Post) = "No" Then   'This is not Cancelled

     Set MItem = Me.ListView7.ListItems.Add(, , Format(rstPaySetup!SerialNo))
     MItem.SubItems(1) = Format(rstPaySetup!Xdate, "dd/mm/yyyy")
     MItem.SubItems(2) = Format(rstPaySetup!Payee)
     MItem.SubItems(3) = Format(rstPaySetup!DateDue, "dd/mm/yyYy")
     MItem.SubItems(4) = Format(rstPaySetup!RefNo)
     MItem.SubItems(5) = Format(rstPaySetup!amtreqested, "#############.#0")

       TotList7 = Val(TotList7) + Val(Trim(rstPaySetup!amtreqested)) 'This is for the Total of the List
 End If
     rstPaySetup.MoveNext
     Wend
txtTotList7.Text = Trim(TotList7)

On Error GoTo 0
'rstJournal

''This is to add the details to the list box
'rstVen.MoveFirst
'While rstVen.EOF = False
'txtvencode.AddItem rstVen!vencode
'cmbPayee2.AddItem rstVen!vennameeng
'rstVen.MoveNext
'Wend


If rstPaySetup.BOF Then
On Error Resume Next
End If



'This is  for ListView4
 rstPaySetup.MoveFirst
  While rstPaySetup.EOF = False
    If Trim(rstPaySetup!cancelledmark) = True Then
     Set MItem = Me.ListView4.ListItems.Add(, , Format(rstPaySetup!SerialNo))
     MItem.SubItems(1) = Format(rstPaySetup!Xdate, "dd/mm/yyyy")
     MItem.SubItems(2) = Format(rstPaySetup!Payee)
     MItem.SubItems(3) = Format(rstPaySetup!DocNo)
     MItem.SubItems(4) = Format(rstPaySetup!amtreqested, "#############.#0")
 On Error Resume Next
   Totlist4 = Val(Totlist4) + Val(Trim(rstPaySetup!TotCrAmt)) 'This is for the Total of the List
    End If
     rstPaySetup.MoveNext
     Wend
Me.txtTotCanceleldBal = Trim(Totlist4)
On Error GoTo 0

If rstSource.EOF = False Then
rstSource.MoveFirst
End If


Do Until rstSource.EOF
       CmbSource.AddItem rstSource!Name
   rstSource.MoveNext
  Loop
rstSource.close

If rstEmp.EOF = False Then
rstEmp.MoveFirst
End If

Do Until rstEmp.EOF
       CmbPrepBy.AddItem rstEmp!Name
       CmbNotedBy.AddItem rstEmp!Name
       CmbApprovedBy.AddItem rstEmp!Name
   rstEmp.MoveNext
  Loop

If rstPAyFor.EOF = False Then
rstPAyFor.MoveFirst
End If


Do Until rstPAyFor.EOF
       cmbPaymentFor.AddItem rstPAyFor!Name
   rstPAyFor.MoveNext
    Loop

If rstCosPro.EOF = False Then
rstCosPro.MoveFirst
End If


Do Until rstCosPro.EOF
        If serial <> rstCosPro!profitcenter Then
       cmbProfCenter.AddItem rstCosPro!profitcenter
       End If
        serial = rstCosPro!profitcenter
         rstCosPro.MoveNext
  Loop



Dim RATEME As New ADODB.Recordset
RATEME.Open "select * from Exrate", CON1, adOpenDynamic, adLockOptimistic
If RATEME.EOF = False Then
RATEME.MoveFirst
End If
Do Until RATEME.EOF
cmbCurrency.AddItem RATEME!zcurrency & "-" & RATEME!zrate
RATEME.MoveNext
Loop





txtToday.Text = Date


End Sub



Private Sub Label61_Click()
frmInvoiceDetails.Show
End Sub


Private Sub List1_Click()
'x = Me.List1.
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  ListView1.SortKey = ColumnHeader.Index - 1
 ' Set Sorted to True to sort the list.
   ListView1.Sorted = True

End Sub

Private Sub ListView1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 93 Then
  PopupMenu frmMenu.ListV
End If

End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = 2 Then
  PopupMenu frmMenu.ListV
End If

End Sub



Private Sub ListView2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  ListView2.SortKey = ColumnHeader.Index - 1
 ' Set Sorted to True to sort the list.
   ListView2.Sorted = True

End Sub

Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)
  
  If frmMenu.edit.caption = "Edit" Then  'This is to Edit So it will bring all the ListView Datas to the Particular TextBoxes
        FrmPayableSetup.txtTmRate.Text = FrmPayableSetup.ListView2.SelectedItem.Text
        FrmPayableSetup.txtTMDes.Text = FrmPayableSetup.ListView2.SelectedItem.SubItems(1)
        FrmPayableSetup.txtTMDays.Text = FrmPayableSetup.ListView2.SelectedItem.SubItems(2)
        FrmPayableSetup.txtTMmode.Text = FrmPayableSetup.ListView2.SelectedItem.SubItems(3)
        FrmPayableSetup.txtTMlevel.Text = FrmPayableSetup.ListView2.SelectedItem.SubItems(4)
        
        FrmPayableSetup.txtTmp1.Text = FrmPayableSetup.txtTmRate.Text
        FrmPayableSetup.txtTmp2.Text = FrmPayableSetup.txtTMDays.Text
      
      frmMenu.edit.caption = "Update"
      frmMenu.clear.caption = "Cancel" 'This will Enable Internal Edit Again"

FrmPayableSetup.txtTMDays.Visible = True
FrmPayableSetup.txtTMDes.Visible = True
FrmPayableSetup.txtTMlevel.Visible = True
FrmPayableSetup.txtTMmode.Visible = True
FrmPayableSetup.txtTmRate.Visible = True
End If

End Sub

Private Sub ListView2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 93 Then
  PopupMenu frmMenu.Terms
End If

End Sub

Private Sub ListView2_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If cmdNew.caption = "&Save" Or cmdNew.caption = "&Update" Then
If Button = 2 Then
  PopupMenu frmMenu.Terms
End If
End If
End Sub


Private Sub ListView3_KeyUp(KeyCode As Integer, Shift As Integer)
If cmdNew.caption = "&Save" Or cmdNew.caption = "&Update" Then

If KeyCode = 93 Then
  PopupMenu frmMenu.rate
End If
'If cmdNew.Caption = "&Save" Then
End If
End Sub

Private Sub ListView3_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
'If cmdNew.Caption = "&Save" Or cmdNew.Caption = "&Update" Then
'
'If Button = 2 Then
'  PopupMenu frmMenu.rate
'End If
'End If
End Sub

Private Sub Text9_Change()

End Sub



Private Sub ListView4_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  ListView4.SortKey = ColumnHeader.Index - 1
 ' Set Sorted to True to sort the list.
   ListView4.Sorted = True

End Sub

Private Sub ListView4_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = 2 Then
  PopupMenu frmMenu.Conf
End If

End Sub


Private Sub ListView5_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  ListView5.SortKey = ColumnHeader.Index - 1
 ' Set Sorted to True to sort the list.
   ListView5.Sorted = True

End Sub

Private Sub ListView5_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = 2 Then
  PopupMenu frmMenu.xPost
End If

End Sub


Private Sub ListView6_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
'If cmdNew.Caption = "&Save" Or cmdNew.Caption = "&Update" Then
'
'If Button = 2 Then
'  PopupMenu frmMenu.Rec
'End If
'End If

End Sub

Private Sub ListView7_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  ListView7.SortKey = ColumnHeader.Index - 1
 ' Set Sorted to True to sort the list.
   ListView7.Sorted = True

End Sub

Private Sub ListView7_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = 2 Then
  PopupMenu frmMenu.List7
End If

End Sub

Private Sub ListView8_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  ListView8.SortKey = ColumnHeader.Index - 1
 ' Set Sorted to True to sort the list.
   ListView8.Sorted = True

End Sub

Private Sub ListView8_DblClick()
ItsForRefrenceForm = ListView8.SelectedItem
Me.txtForFrmRef = ItsForRefrenceForm
frmREfrence.Show
End Sub

Private Sub ListView8_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = 2 Then
  PopupMenu frmMenu.Journal
End If

End Sub

Private Sub ListView9_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = 2 Then
  PopupMenu frmMenu.Jounal1
End If

End Sub

Private Sub MskDate_LostFocus()
If Me.MskDate.Text = "__/__/____" Then
    Exit Sub
End If
Me.MskDate.Text = Format(Me.MskDate.Text, "dd/mm/yyyy")
cDay = Val(Left(Me.MskDate.Text, 2))
cMonth = Val(Mid(Me.MskDate.Text, 4, 2))
cYear = Val(Right(Me.MskDate.Text, 4))
If cDay > 31 Or cDay < 1 Then
    mess = MsgBox("Invalid Date ÊÇÑíÎ ÎØÇð", vbInformation + vbOKOnly, "Message")
    Me.MskDate.SetFocus
  ElseIf cMonth > 12 Or cMonth < 1 Then
    mess = MsgBox("Invalid Month ÔåÑ ÎØÇð", vbInformation + vbOKOnly, "Message")
    Me.MskDate.SetFocus
ElseIf cYear < 1900 Or cYear > Year(Date) Then
    mess = MsgBox("Invalid YearÓäÉ ÎØÇð", vbInformation + vbOKOnly, "Message")
    Me.MskDate.SetFocus
End If

End Sub

Private Sub mskDateDue_LostFocus()
If Me.mskDateDue.Text = "__/__/____" Then
    Exit Sub
End If
Me.mskDateDue.Text = Format(Me.mskDateDue.Text, "dd/mm/yyyy")
cDay = Val(Left(Me.mskDateDue.Text, 2))
cMonth = Val(Mid(Me.mskDateDue.Text, 4, 2))
cYear = Val(Right(Me.mskDateDue.Text, 4))
If cDay > 31 Or cDay < 1 Then
    mess = MsgBox("Invalid Date ÊÇÑíÎ ÎØÇð", vbInformation + vbOKOnly, "Message")
    Me.mskDateDue.SetFocus
  ElseIf cMonth > 12 Or cMonth < 1 Then
    mess = MsgBox("Invalid Month ÔåÑ ÎØÇð", vbInformation + vbOKOnly, "Message")
    Me.mskDateDue.SetFocus
'ElseIf cYear < 1900 Or cYear > Year(Date) Then
'    mess = MsgBox("Invalid Year", vbInformation + vbOKOnly, "Message")
'    Me.mskDateDue.SetFocus
End If
'Dim reem As Date
'
'reem = Me.mskDateDue.Text
'If reem < Date Then
'MsgBox "Date Due can not be The Date Before the Current Date"
'Me.mskDateDue.SetFocus
'End If
End Sub

'Private Sub mskInvDate_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'txtPoNo.SetFocus
'End If
'
'End Sub

'Private Sub mskInvDate_LostFocus()
'If Me.mskInvDate.Text = "__/__/____" Then
'    Exit Sub
'End If
'Me.mskInvDate.Text = Format(Me.mskInvDate.Text, "dd/mm/yyyy")
'cDay = Val(Left(Me.mskInvDate.Text, 2))
'cMonth = Val(Mid(Me.mskInvDate.Text, 4, 2))
'cYear = Val(Right(Me.mskInvDate.Text, 4))
'If cDay > 31 Or cDay < 1 Then
'    mess = MsgBox("Invalid Date ÊÇÑíÎ ÎØÇð", vbInformation + vbOKOnly, "Message")
'    Me.mskInvDate.SetFocus
'  ElseIf cMonth > 12 Or cMonth < 1 Then
'    mess = MsgBox("Invalid Month ÔåÑ ÎØÇð", vbInformation + vbOKOnly, "Message")
'    Me.mskInvDate.SetFocus
'ElseIf cYear < 1900 Or cYear > Year(Date) Then
'    mess = MsgBox("Invalid Year ÓäÉ ÎØÇð", vbInformation + vbOKOnly, "Message")
'    Me.mskInvDate.SetFocus
'End If
'
'End Sub

'Private Sub mskPOdate_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'txtStorEntryNo.SetFocus
'End If
'End Sub

'Private Sub mskStoreEnDate_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'txtProTax.SetFocus
'End If
'
'End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbProfCenter.SetFocus
End If

End Sub


Private Sub txtAccNo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyPageUp Or KeyCode = vbKeyUp Then
MskDate.SetFocus
End If

If KeyCode = vbKeyPageDown Or KeyCode = vbKeyDown Then
txtPartic.SetFocus
End If
End Sub

Private Sub txtAccNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtPartic.SetFocus
End If
End Sub

Private Sub txtAmo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyPageUp Or KeyCode = vbKeyUp Then
txtPartic.SetFocus
End If

If KeyCode = vbKeyPageDown Or KeyCode = vbKeyDown Then
CmbPrepBy.SetFocus
End If

End Sub

Private Sub txtAmo_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'
'     Set MItem = Me.ListView3.ListItems.Add(, , Trim(txtAccNo.Text))
'     MItem.SubItems(1) = Trim(txtPartic.Text)
'     MItem.SubItems(2) = Trim(txtAmo.Text)
'TotList3 = Val(txtTotList3) + Val(Trim(txtAmo.Text)) 'This is for the Total of the List
'txtAccNo.Text = ""
'txtPartic.Text = ""
'txtAmo.Text = ""
'txtAccNo.SetFocus
''CmbPrepBy.SetFocus
'txtTotList3.Text = TotList3
'End If
End Sub

'Private Sub mskStoreEnDate_LostFocus()
'If Me.mskStoreEnDate.Text = "__/__/____" Then
'    Exit Sub
'End If
'Me.mskStoreEnDate.Text = Format(Me.mskStoreEnDate.Text, "dd/mm/yyyy")
'cDay = Val(Left(Me.mskStoreEnDate.Text, 2))
'cMonth = Val(Mid(Me.mskStoreEnDate.Text, 4, 2))
'cYear = Val(Right(Me.mskStoreEnDate.Text, 4))
'If cDay > 31 Or cDay < 1 Then
'    mess = MsgBox("Invalid Date", vbInformation + vbOKOnly, "Message")
'    Me.mskStoreEnDate.SetFocus
'  ElseIf cMonth > 12 Or cMonth < 1 Then
'    mess = MsgBox("Invalid Month", vbInformation + vbOKOnly, "Message")
'    Me.mskStoreEnDate.SetFocus
'ElseIf cYear < 1900 Or cYear > Year(Date) Then
'    mess = MsgBox("Invalid Year", vbInformation + vbOKOnly, "Message")
'    Me.mskStoreEnDate.SetFocus
'End If
'
'End Sub

Private Sub Option1_Click()
Command3.Enabled = True
txtInvAmt.Enabled = True
'txtInvNo.Enabled = True

End Sub

Private Sub Option1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command3.SetFocus
End If
End Sub

Private Sub Option2_Click()
Command3.Enabled = False
txtInvAmt.Locked = False
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.caption = "Payable Entry" Then
SSTab1.ToolTipText = "Step1"
ElseIf SSTab1.caption = "List View" Then
SSTab1.ToolTipText = "Step2 & Step4"
ElseIf SSTab1.caption = "Cancelled List" Then
SSTab1.ToolTipText = "Step3"
ElseIf SSTab1.caption = "Paid Voucher" Then
SSTab1.ToolTipText = "Step5"

ElseIf SSTab1.caption = "Journal" Then
SSTab1.ToolTipText = "Step6"
End If
End Sub


Private Sub Timer1_Timer()
  
    Dim serv As New ADODB.Recordset
    serv.Open "select * from SerialNom", constring, adOpenDynamic, adLockOptimistic
    If serv.EOF = False Then
    serv.MoveFirst
    End If
    While serv.EOF = False
    If Val(Me.txtserialNo) = Val(serv!SENo) Then
      txtserialNo.Text = serv!SENo
    End If
     serv.MoveNext
    Wend

End Sub

Private Sub txtAmountDue_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtAmtReq.SetFocus
End If

End Sub

Private Sub txtAmountDue_LostFocus()
txtAmountDue.Text = Format(txtAmountDue.Text, "############.#0")

End Sub

Private Sub txtAmtPaid_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtCreditNote.SetFocus
End If

End Sub

Private Sub txtAmtPaid_LostFocus()
txtAmtPaid = Format(txtAmtPaid.Text, "############.#0")

Dim uiu As Currency
Dim iii As Currency
On Error Resume Next
uiu = txtInvAmt.Text
iii = txtAmtPaid.Text

If Val(uiu) < Val(iii) Then
MsgBox "Amount Paid can not be more than the Invoice Amount ÇÌãÇáí ÇáãÏÝÚ áÇíÓÇæí ÇÌãÇáí ÇáÝÇÊæÑÉ ", vbCritical, "Error"
Exit Sub
End If
txtRefNo_LostFocus
End Sub



Private Sub txtAmtReq_GotFocus()
txtRefNo_LostFocus

End Sub

Private Sub txtAmtReq_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
mskDateDue.SetFocus
End If

End Sub

Private Sub txtAmtReq_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 17 Or KeyCode = 67 Or KeyCode = 86 Or KeyCode = 8 Or KeyCode = 18 Or KeyCode = 46 Or KeyCode = 27 Then

Exit Sub
End If
 
 If KeyCode <> 110 Then
 If KeyCode < 96 Or KeyCode > 105 Then

 If KeyCode < 48 Or KeyCode > 57 Then
    MsgBox "The Key you entered is not the Digit", vbExclamation, "Try Again"
    Me.txtAmtReq = ""
    Exit Sub
End If
End If
End If


End Sub

Private Sub txtAmtReq_LostFocus()
txtAmtReq.Text = Format(txtAmtReq.Text, "############.#0")
Dim xxx As Currency
Dim yyy As Currency

On Error Resume Next
xxx = txtOutBal.Text
yyy = txtAmtReq.Text
If cmbPaymentFor <> "P.O" Then
If Val(xxx) < Val(yyy) Then
MsgBox "Amount Requested can't be more than Outstanding Balance"
txtAmtReq.SetFocus
End If
End If
End Sub

Private Sub txtBranch_GotFocus()
Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(txtBranch.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)

End Sub

'



Private Sub txtBranch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtEnter1.SetFocus
End If

End Sub


Private Sub MskDate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
CmbSource.SetFocus
End If
End Sub

Private Sub txtDBtot_Change()

End Sub



Private Sub txtDBAccNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtDBPartic.SetFocus
End If

End Sub

Private Sub txtDBAmo_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'
'If frmMenu.edit.Caption = "Edit" And FrmPayableSetup.cmdNew.Caption = "&Update" Then
''This will save the additional field to the file insted of saving all the available datas in the listbox
'xok = MsgBox("Are you sure you want to save this Debit details", vbOKCancel, "SAVE")
'    If xok = vbOK Then
'    'Save
'
'     With rstReceipt
'        .AddNew
'            rstReceipt!Serialno = FrmPayableSetup.txtserialNo.Text
'            rstReceipt!accno = FrmPayableSetup.txtDBAccNo.Text
'            rstReceipt!accname = FrmPayableSetup.txtDBPartic.Text
'            rstReceipt!amount = FrmPayableSetup.txtDBAmo.Text
'            rstReceipt!total = FrmPayableSetup.txtTotList6.Text
'       '   frmMenu.edit.Caption = "&Edit"
'
'     .Update
'
'    MsgBox "Records for Receipt Details Saved Seccussfully"
'    End With
'    End If
''0-0-0-0--0--0--
'Else
'
'
'     Set MItem = Me.ListView6.ListItems.Add(, , Trim(txtDBAccNo.Text))
'     MItem.SubItems(1) = Trim(txtDBPartic.Text)
'     MItem.SubItems(2) = Trim(txtDBAmo.Text)
'TotList6 = Val(txtTotList6) + Val(Trim(txtDBAmo.Text)) 'This is for the Total of the List
'txtDBAccNo.Text = ""
'txtDBPartic.Text = ""
'txtDBAmo.Text = ""
'txtDBAccNo.SetFocus
''CmbPrepBy.SetFocus
'txtTotList6.Text = TotList6
'End If
'End If
End Sub



Private Sub txtCreditNote_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtDebitNote.SetFocus
End If

End Sub

Private Sub txtCreditNote_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode <> 13 Then
 If KeyCode <> 110 Then
 If KeyCode < 96 Or KeyCode > 105 Then

 If KeyCode < 48 Or KeyCode > 57 Then
    MsgBox "The Key you entered is not the Digit", vbExclamation, "Try Again"
    Me.txtCreditNote = ""
    Exit Sub
End If
End If
End If
End If

End Sub

Private Sub txtCreditNote_LostFocus()
txtCreditNote.Text = Format(txtCreditNote.Text, "############.#0")
txtRefNo_LostFocus
End Sub

Private Sub txtDebitNote_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtOutBal.SetFocus
End If

End Sub

Private Sub txtDebitNote_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode <> 13 Then
 If KeyCode <> 110 Then
 If KeyCode < 96 Or KeyCode > 105 Then

 If KeyCode < 48 Or KeyCode > 57 Then
    MsgBox "The Key you entered is not the Digit", vbExclamation, "Try Again"
    Me.txtDebitNote = ""
    Exit Sub
End If
End If
End If
End If
End Sub

Private Sub txtDebitNote_LostFocus()
txtDebitNote.Text = Format(txtDebitNote.Text, "############.#0")
txtRefNo_LostFocus
End Sub




Private Sub txtDocuNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtBranch.SetFocus
End If
End Sub


Private Sub mskDateDue_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtTmRate.SetFocus
End If

End Sub



Private Sub txtEnter1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyControl Then
cmbProfCenter.SetFocus
End If
End Sub

Private Sub txtEnter1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And txtEnter1 <> "" Then
List1.AddItem txtEnter1.Text
txtEnter1.Text = ""
txtEnter1.SetFocus
End If

End Sub

Private Sub TxtEnter2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyControl Then
txtRefNo.SetFocus
End If
End Sub

Private Sub TxtEnter2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And TxtEnter2 <> "" Then
txtNoOfAtt.AddItem TxtEnter2.Text
TxtEnter2.Text = ""
TxtEnter2.SetFocus
End If

End Sub

Private Sub txtInvAmt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtAmtPaid.SetFocus
End If

End Sub

Private Sub txtInvAmt_LostFocus()
txtInvAmt.Text = Format(txtInvAmt.Text, "############.#0")

End Sub

'Private Sub txtinvNo_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'txtInvAmt.SetFocus
'End If
'
'End Sub

Private Sub txtJCr_Change()
txtJCr = Format(txtJCr.Text, "############.#0")

End Sub

Private Sub txtJdb_Change()
txtJdb = Format(txtJdb.Text, "############.#0")

End Sub

Private Sub txtNoOfAtt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Option1.SetFocus
'txtinvNo.SetFocus
End If

End Sub

Private Sub txtOutBal_GotFocus()
txtRefNo_LostFocus
End Sub

Private Sub txtOutBal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtTaxCredit.SetFocus
End If

End Sub





'Private Sub txtPercentage_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'txtInvAmt.SetFocus
'End If
'
'End Sub

'Private Sub txtPoNo_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'mskPOdate.SetFocus
'End If
'
'End Sub

'Private Sub txtProTax_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'txtPercentage.SetFocus
'End If
'
'End Sub

'Private Sub txtProTax_LostFocus()
'txtProTax.Text = Format(txtProTax.Text, "############.#0")
'End Sub


Private Sub txtRefNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Option1.SetFocus
End If
End Sub


Private Sub txtRefNo_LostFocus()

If cmbPaymentFor.Text <> "P.O" Then
Dim q As Currency
Dim w As Currency
Dim r As Currency
Dim s As Currency
Dim zzz As Currency

  On Error Resume Next
    q = Format(txtInvAmt.Text, "###,###,###.#0")
    w = Format(txtDebitNote.Text, "###,###,###.#0")
    r = Format(txtCreditNote.Text, "###,###,###.#0")
    s = Format(txtAmtPaid.Text, "###,###,###.#0")
   zzz = Format(txtTaxCredit.Text, "###,###,###.#0")

Dim outs



outs = Val(q) + Val(w) - Val(r) - Val(s) - Val(zzz)

txtOutBal = Val(outs)
End If
End Sub


Private Sub txtSearch1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmdSearch_Click
End If

End Sub

Private Sub txtSearch2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmdSearch2_Click
End If
End Sub

Private Sub txtSearch3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmdSearch3_Click
End If

End Sub

Private Sub txtserialNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
'MskDate.SetFocus
End If

End Sub

'Private Sub txtStorEntryNo_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'mskStoreEnDate.SetFocus
'End If
'
'End Sub

Private Sub txtTaxCredit_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtAmountDue.SetFocus
End If

End Sub


Private Sub txtTaxCredit_LostFocus()
txtRefNo_LostFocus

End Sub

Private Sub txtTMDays_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtTMmode.SetFocus
End If

End Sub

Private Sub txtTMDes_GotFocus()
Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(txtTMDes.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)

End Sub

Private Sub txtTMDes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtTMDays.SetFocus
End If

End Sub

Private Sub txtTMlevel_KeyPress(KeyAscii As Integer)



If KeyAscii = 13 Then
  
  If frmMenu.edit.caption = "Edit" And FrmPayableSetup.cmdNew.caption = "&Update" Then
  'This will be when we save only this list and save shuld be done from the particular textboxes so it should not be cleared
   txtTMlevel_LostFocus
   Exit Sub
  End If

If frmMenu.edit.caption = "Update" And FrmPayableSetup.cmdNew.caption = "&Save" Then
 
 DeleteTermj = FrmPayableSetup.ListView2.SelectedItem.Index

 
 'DeleteTerm = TxtKallapundai.Text
 FrmPayableSetup.ListView2.ListItems.Remove DeleteTermj
     Set MItem = Me.ListView2.ListItems.Add(, , Trim(txtTmRate.Text))
     MItem.SubItems(1) = Trim(txtTMDes.Text)
     MItem.SubItems(2) = Trim(txtTMDays.Text)
     MItem.SubItems(3) = Trim(txtTMmode.Text)
     MItem.SubItems(4) = Trim(txtTMlevel.Text)
  ' End If
'Next
txtTMDes.Text = ""
txtTMDays.Text = ""
txtTMmode.Text = ""
txtTMlevel.Text = ""
txtTmRate.Text = ""
txtTmRate.SetFocus



Exit Sub
End If

     Set MItem = Me.ListView2.ListItems.Add(, , Trim(txtTmRate.Text))
     MItem.SubItems(1) = Trim(txtTMDes.Text)
     MItem.SubItems(2) = Trim(txtTMDays.Text)
     MItem.SubItems(3) = Trim(txtTMmode.Text)
     MItem.SubItems(4) = Trim(txtTMlevel.Text)
     
     
txtTMDes.Text = ""
txtTMDays.Text = ""
txtTMmode.Text = ""
txtTMlevel.Text = ""
txtTmRate.Text = ""
txtTmRate.SetFocus
End If

End Sub

Private Sub txtTMlevel_LostFocus()
  If frmMenu.edit.caption = "Edit" And FrmPayableSetup.cmdNew.caption = "&Save" Then
    If KeyAscii = 13 Then
     Set MItem = Me.ListView2.ListItems.Add(, , Trim(txtTmRate.Text))
     MItem.SubItems(1) = Trim(txtTMDes.Text)
     MItem.SubItems(2) = Trim(txtTMDays.Text)
     MItem.SubItems(3) = Trim(txtTMmode.Text)
     MItem.SubItems(4) = Trim(txtTMlevel.Text)

txtTmRate.Text = ""
txtTMDes.Text = ""
txtTMmode.Text = ""
txtTMDays.Text = ""
txtTMlevel.Text = ""
txtTmRate.SetFocus
End If


ElseIf frmMenu.edit.caption = "Edit" And FrmPayableSetup.cmdNew.caption = "&Update" Then
'This will save the additional field to the file insted of saving all the available datas in the listbox
xok = MsgBox("Do you want to Add the Term details åá ÊÑíÏ ÇÖÇÝÉ ÔÑæØ ÊÝÇÕíáíÉ ", vbOKCancel, "SAVE")
If xok = vbOK Then
'Save
      With rstterm
      .addnew
     
     rstterm!SerialNo = FrmPayableSetup.txtserialNo.Text
     rstterm!rate = Trim(txtTmRate.Text)
     rstterm!descr = Trim(txtTMDes.Text)
     rstterm!days = Trim(txtTMDays.Text)
     rstterm!Mode = Trim(txtTMmode.Text)
     rstterm!xlevel = Trim(txtTMlevel.Text)

           
    .Update
    End With
    
    
    
         Set MItem = Me.ListView2.ListItems.Add(, , Trim(txtTmRate.Text))
     MItem.SubItems(1) = Trim(txtTMDes.Text)
     MItem.SubItems(2) = Trim(txtTMDays.Text)
     MItem.SubItems(3) = Trim(txtTMmode.Text)
     MItem.SubItems(4) = Trim(txtTMlevel.Text)

txtTmRate.Text = ""
txtTMDes.Text = ""
txtTMmode.Text = ""
txtTMDays.Text = ""
txtTMlevel.Text = ""
txtTmRate.SetFocus

MsgBox "New Data saved succesfully"
Exit Sub

''..This is to REFRESH the ListView, after updating the Details
   FrmPayableSetup.ListView2.ListItems.clear
   
If rstterm.EOF = False Then
rstterm.MoveFirst
End If

   While rstterm.EOF = False
       
            If Trim(rstterm!SerialNo) = Trim(FrmPayableSetup.txtserialNo.Text) Then
        
     Set MItem = FrmPayableSetup.ListView2.ListItems.Add(, , Trim(rstterm!rate))
     MItem.SubItems(1) = Trim(rstterm!descr)
     MItem.SubItems(2) = Trim(rstterm!days)
     MItem.SubItems(3) = Trim(rstterm!Mode)
     MItem.SubItems(4) = Trim(rstterm!xlevel)
        End If
     rstterm.MoveNext
     Wend
'........................................
 


End If

      txtTMDes.Text = ""
      txtTMDays.Text = ""
      txtTMmode.Text = ""
      txtTMlevel.Text = ""
      txtTmRate.Text = ""
      
End If
txtTmRate.SetFocus
End Sub

Private Sub txtTMmode_GotFocus()
Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(txtTMmode.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)

End Sub

Private Sub txtTMmode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtTMlevel.SetFocus
End If

End Sub

Private Sub txtTmRate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeySpacen Then
CmbPrepBy.SetFocus
End If
End Sub

Private Sub txtTmRate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtTMDes.SetFocus
End If

End Sub

Private Sub txtTotCanceleldBal_Change()
txtTotCanceleldBal = Format(txtTotCanceleldBal.Text, "############.#0")

End Sub

Private Sub txtTotList1_Change()
txtTotList1 = Format(txtTotList1.Text, "############.#0")

End Sub

Private Sub txtTotList7_Change()
txtTotList7 = Format(txtTotList7.Text, "############.#0")

End Sub

Private Sub txtTotVoch_Change()
txtTotVoch = Format(txtTotVoch.Text, "############.#0")

End Sub





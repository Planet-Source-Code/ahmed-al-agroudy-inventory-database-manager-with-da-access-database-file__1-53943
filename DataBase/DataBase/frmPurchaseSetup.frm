VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPurchaseSetup 
   Caption         =   "Purchase Setup"
   ClientHeight    =   7770
   ClientLeft      =   615
   ClientTop       =   60
   ClientWidth     =   8880
   Icon            =   "frmPurchaseSetup.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7770
   ScaleWidth      =   8880
   Begin TabDlg.SSTab SSTab1 
      Height          =   7815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   13785
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Purchase Details"
      TabPicture(0)   =   "frmPurchaseSetup.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label13"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label7"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtStorEnDate"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label16"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label9"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label11"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label3"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label4"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label5"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label8"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label17"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label18"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label19"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label20"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label21"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label22"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label25"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label32"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label12"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "mskDateDue"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "MskDate"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "mskStoreEnDate"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "cmbCostCenter"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "cmbProfCenter"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtDocuNo"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtStorEntryNo"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtvenCode"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Combo5"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "List1"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "CmbSource"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "txtRefNo"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "txtAmountDue"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "txtAmtPaid"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "txtAmtReq"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "txtOutBal"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "txtProTax"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "txtPercentage"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "txtNoOfAtt"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "txtserialNo"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "Frame1"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "txtPoNo"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "Frame2"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "txtBranch"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "Frame4"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).ControlCount=   47
      TabCaption(1)   =   "List"
      TabPicture(1)   =   "frmPurchaseSetup.frx":045E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label23"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "ListView1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "txtTotList7"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "txtSearch1"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame3"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdSearch"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "CmdEdit"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      Begin VB.Frame Frame4 
         Height          =   735
         Left            =   -68160
         TabIndex        =   78
         Top             =   3480
         Width           =   615
         Begin VB.CommandButton Command3 
            BackColor       =   &H8000000A&
            Height          =   340
            Left            =   120
            Picture         =   "frmPurchaseSetup.frx":047A
            Style           =   1  'Graphical
            TabIndex        =   79
            ToolTipText     =   "Find Items"
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.ComboBox txtBranch 
         Height          =   315
         Left            =   -73320
         TabIndex        =   77
         Top             =   2520
         Width           =   2055
      End
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   -67440
         TabIndex        =   71
         Top             =   3480
         Width           =   3495
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
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   74
            Top             =   240
            Width           =   1000
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
            Height          =   325
            Left            =   1320
            TabIndex        =   73
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
            Left            =   2400
            TabIndex        =   72
            Top             =   240
            Width           =   1000
         End
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
         Left            =   240
         TabIndex        =   69
         Top             =   6240
         Width           =   1000
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
         Left            =   1320
         TabIndex        =   61
         Top             =   6240
         Width           =   1000
      End
      Begin VB.Frame Frame3 
         Height          =   975
         Left            =   120
         TabIndex        =   54
         Top             =   6600
         Width           =   11055
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
            Left            =   8280
            TabIndex        =   57
            Top             =   480
            Width           =   2655
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
            Left            =   4080
            TabIndex        =   56
            Top             =   480
            Width           =   2655
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
            Left            =   120
            TabIndex        =   55
            Top             =   480
            Width           =   2655
         End
         Begin VB.Label Label29 
            Caption         =   "Approved By"
            Height          =   255
            Left            =   8280
            TabIndex        =   60
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label28 
            Caption         =   "Noted By"
            Height          =   255
            Left            =   4080
            TabIndex        =   59
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label27 
            Caption         =   "Prepared By"
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.TextBox txtSearch1 
         Height          =   325
         Left            =   2400
         TabIndex        =   53
         Top             =   6240
         Width           =   1455
      End
      Begin VB.TextBox txtTotList7 
         Height          =   330
         Left            =   9480
         TabIndex        =   52
         Top             =   6240
         Width           =   1695
      End
      Begin VB.TextBox txtPoNo 
         Height          =   320
         Left            =   -67800
         TabIndex        =   27
         Top             =   1500
         Width           =   1215
      End
      Begin VB.Frame Frame1 
         Caption         =   "Item"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Left            =   -74880
         TabIndex        =   18
         Top             =   4320
         Width           =   10935
         Begin VB.ComboBox TxtItemDes 
            Height          =   315
            Left            =   1320
            TabIndex        =   76
            Top             =   240
            Width           =   2535
         End
         Begin VB.ComboBox txtItemCode 
            Height          =   315
            Left            =   120
            TabIndex        =   75
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txtItemTemp1List1 
            Alignment       =   1  'Right Justify
            Height          =   320
            Left            =   1320
            TabIndex        =   68
            Top             =   1080
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.TextBox txtTotalInvInn 
            Height          =   320
            Left            =   9600
            TabIndex        =   67
            Top             =   3000
            Width           =   1215
         End
         Begin VB.TextBox txtItemQty 
            Height          =   320
            Left            =   5565
            TabIndex        =   66
            Top             =   240
            Width           =   570
         End
         Begin VB.TextBox txtItemDiam 
            Height          =   320
            Left            =   4680
            TabIndex        =   65
            Top             =   240
            Width           =   930
         End
         Begin VB.TextBox txtItemModelNo 
            Height          =   320
            Left            =   3840
            TabIndex        =   64
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txtTmp1 
            Height          =   285
            Left            =   1080
            TabIndex        =   25
            Text            =   "Text2"
            Top             =   720
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox txtTmp2 
            Height          =   285
            Left            =   2040
            TabIndex        =   24
            Text            =   "Text2"
            Top             =   720
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox txtItemPrize 
            Height          =   320
            Left            =   6125
            TabIndex        =   23
            Top             =   240
            Width           =   930
         End
         Begin VB.TextBox txtVAt 
            Height          =   320
            Left            =   7040
            TabIndex        =   22
            Top             =   240
            Width           =   930
         End
         Begin VB.TextBox txtSurTax 
            Height          =   320
            Left            =   7950
            TabIndex        =   21
            Top             =   240
            Width           =   930
         End
         Begin VB.TextBox txtTaxCredit 
            Height          =   320
            Left            =   8860
            TabIndex        =   20
            Top             =   240
            Width           =   930
         End
         Begin VB.TextBox txtItemInvInn 
            Height          =   320
            Left            =   9800
            TabIndex        =   19
            Top             =   240
            Width           =   930
         End
         Begin MSComctlLib.ListView ListView2 
            Height          =   2295
            Left            =   120
            TabIndex        =   26
            Top             =   720
            Width           =   10650
            _ExtentX        =   18785
            _ExtentY        =   4048
            View            =   3
            LabelWrap       =   0   'False
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin VB.Label Label10 
            Caption         =   "Total Inventory In"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7440
            TabIndex        =   70
            Top             =   3000
            Width           =   1815
         End
      End
      Begin VB.TextBox txtserialNo 
         Alignment       =   1  'Right Justify
         Height          =   320
         Left            =   -73350
         TabIndex        =   17
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtNoOfAtt 
         Height          =   675
         Left            =   -73350
         TabIndex        =   16
         Top             =   3600
         Width           =   3015
      End
      Begin VB.TextBox txtPercentage 
         Height          =   320
         Left            =   -65280
         TabIndex        =   15
         Top             =   2580
         Width           =   735
      End
      Begin VB.TextBox txtProTax 
         Height          =   320
         Left            =   -65280
         TabIndex        =   14
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox txtOutBal 
         Height          =   320
         Left            =   -67830
         TabIndex        =   13
         Top             =   2580
         Width           =   1215
      End
      Begin VB.TextBox txtAmtReq 
         Height          =   320
         Left            =   -65280
         TabIndex        =   12
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox txtAmtPaid 
         Height          =   320
         Left            =   -67830
         TabIndex        =   11
         Top             =   2220
         Width           =   1215
      End
      Begin VB.TextBox txtAmountDue 
         Height          =   320
         Left            =   -65280
         TabIndex        =   10
         Top             =   3000
         Width           =   1335
      End
      Begin VB.TextBox txtRefNo 
         Height          =   320
         Left            =   -67830
         TabIndex        =   9
         Top             =   2880
         Width           =   1215
      End
      Begin VB.ComboBox CmbSource 
         Height          =   315
         Left            =   -73320
         TabIndex        =   8
         Top             =   1080
         Width           =   3015
      End
      Begin VB.ListBox List1 
         DataField       =   "Branch"
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         Height          =   645
         Left            =   -73350
         TabIndex        =   7
         Top             =   2880
         Width           =   3015
      End
      Begin VB.ComboBox Combo5 
         Height          =   315
         ItemData        =   "frmPurchaseSetup.frx":0A14
         Left            =   -73350
         List            =   "frmPurchaseSetup.frx":0A16
         OLEDropMode     =   1  'Manual
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   1485
         Width           =   3015
      End
      Begin VB.TextBox txtvenCode 
         Alignment       =   1  'Right Justify
         Height          =   320
         Left            =   -73350
         TabIndex        =   5
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox txtStorEntryNo 
         Height          =   320
         Left            =   -67830
         TabIndex        =   4
         Top             =   1860
         Width           =   1215
      End
      Begin VB.TextBox txtDocuNo 
         Height          =   320
         Left            =   -73350
         TabIndex        =   3
         Top             =   1800
         Width           =   3015
      End
      Begin VB.ComboBox cmbProfCenter 
         Height          =   315
         Left            =   -67830
         TabIndex        =   2
         Top             =   480
         Width           =   3855
      End
      Begin VB.ComboBox cmbCostCenter 
         Height          =   315
         Left            =   -67830
         TabIndex        =   1
         Top             =   810
         Width           =   3855
      End
      Begin MSMask.MaskEdBox mskStoreEnDate 
         Height          =   315
         Left            =   -65250
         TabIndex        =   28
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "mm/dd/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MskDate 
         Height          =   315
         Left            =   -73350
         TabIndex        =   29
         Top             =   780
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "mm/dd/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskDateDue 
         Height          =   315
         Left            =   -67830
         TabIndex        =   30
         Top             =   1110
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "mm/dd/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   5775
         Left            =   120
         TabIndex        =   62
         Top             =   360
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   10186
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
         NumItems        =   0
      End
      Begin VB.Label Label12 
         Caption         =   "Amount Due"
         Height          =   255
         Left            =   -66480
         TabIndex        =   80
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label Label23 
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
         Height          =   375
         Left            =   7320
         TabIndex        =   63
         Top             =   6240
         Width           =   2655
      End
      Begin VB.Label Label32 
         Caption         =   "Po Number"
         Height          =   255
         Left            =   -69240
         TabIndex        =   51
         Top             =   1500
         Width           =   1215
      End
      Begin VB.Label Label25 
         Caption         =   "Serial No"
         Height          =   255
         Left            =   -74730
         TabIndex        =   50
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label22 
         Caption         =   " Attechmets"
         Height          =   255
         Left            =   -74760
         TabIndex        =   49
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Label Label21 
         Caption         =   "Percentage"
         Height          =   255
         Left            =   -66480
         TabIndex        =   48
         Top             =   2580
         Width           =   1095
      End
      Begin VB.Label Label20 
         Caption         =   "Profit Taxes"
         Height          =   255
         Left            =   -66480
         TabIndex        =   47
         Top             =   2220
         Width           =   1215
      End
      Begin VB.Label Label19 
         Caption         =   "Outstanding Bal"
         Height          =   255
         Left            =   -69240
         TabIndex        =   46
         Top             =   2580
         Width           =   1335
      End
      Begin VB.Label Label18 
         Caption         =   "Amt Requested"
         Height          =   255
         Left            =   -66480
         TabIndex        =   45
         Top             =   1860
         Width           =   1215
      End
      Begin VB.Label Label17 
         Caption         =   "Amt Paid  Brfore"
         Height          =   255
         Left            =   -69240
         TabIndex        =   44
         Top             =   2220
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Reference No"
         Height          =   255
         Left            =   -69240
         TabIndex        =   43
         Top             =   2895
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Date"
         Height          =   255
         Left            =   -74730
         TabIndex        =   42
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Branch"
         Height          =   255
         Left            =   -74730
         TabIndex        =   41
         Top             =   2490
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Payee"
         Height          =   255
         Left            =   -74730
         TabIndex        =   40
         Top             =   1485
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Purchase From"
         Height          =   255
         Left            =   -74730
         TabIndex        =   39
         Top             =   1185
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "Explanation"
         Height          =   255
         Left            =   -74730
         TabIndex        =   38
         Top             =   2820
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Vendor Code"
         Height          =   255
         Left            =   -74730
         TabIndex        =   37
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label16 
         Caption         =   "Store Entry No"
         Height          =   255
         Left            =   -69240
         TabIndex        =   36
         Top             =   1860
         Width           =   1095
      End
      Begin VB.Label txtStorEnDate 
         Caption         =   " Date"
         Height          =   255
         Left            =   -66480
         TabIndex        =   35
         Top             =   1500
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Document No"
         Height          =   255
         Left            =   -74730
         TabIndex        =   34
         Top             =   1815
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Date Due"
         Height          =   255
         Left            =   -69210
         TabIndex        =   33
         Top             =   1140
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "Profit Center"
         Height          =   255
         Left            =   -69210
         TabIndex        =   32
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Cost Center"
         Height          =   255
         Left            =   -69210
         TabIndex        =   31
         Top             =   810
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmPurchaseSetup"
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
Dim RstItem As ADODB.Recordset
Dim rstChart As ADODB.Recordset
Dim rstVen As ADODB.Recordset
Dim rstPAyFor As ADODB.Recordset
Dim rstCosPro As ADODB.Recordset
Dim strVen As String
Dim serial As String
Dim RstPurchaseSetup As ADODB.Recordset
Dim RstINVcat As ADODB.Recordset

Dim rstterm As ADODB.Recordset
Dim TotList2
Dim Totlist1

Private Sub CmbApprovedBy_GotFocus()
Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(CmbApprovedBy.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cmbCostCenter_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
mskDateDue.SetFocus
End If

End Sub

Private Sub CmbNotedBy_GotFocus()
Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(CmbNotedBy.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub CmbPrepBy_GotFocus()
Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(CmbPrepBy.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)

End Sub

Private Sub cmbProfCenter_Click()
Dim serialpro As String
Dim serialdep As String
Dim serialxsec As String
cmbCostCenter.clear
rstCosPro.MoveFirst
Do Until rstCosPro.EOF
        
        If cmbProfCenter = rstCosPro!profitcenter Then
            If serialpro <> rstCosPro!productionUnit Then
            cmbCostCenter.AddItem rstCosPro!productionUnit
            End If
            serialpro = rstCosPro!productionUnit
 
             If serialdep <> rstCosPro!department Then
            cmbCostCenter.AddItem rstCosPro!department
            End If
            serialdep = rstCosPro!department

            If serialxsec <> rstCosPro!xsection Then
            cmbCostCenter.AddItem rstCosPro!xsection
            End If
            serialxsec = rstCosPro!xsection

 
        End If

         rstCosPro.MoveNext
  Loop

End Sub

Private Sub cmbProfCenter_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbCostCenter.SetFocus
End If

End Sub

Private Sub CmbSource_Click()
CmbSource_LostFocus
End Sub

Private Sub CmbSource_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
'Combo5.SetFocus
End If
End Sub

Private Sub CmbSource_LostFocus()
Combo5.clear
    If Trim(CmbSource.Text) = "Supplier" Then 'This is for Vendor
Combo5.Enabled = True
rstVen.MoveFirst
Do Until rstVen.EOF
       'Combo5.AddItem rstVen!vencode & "    " & rstVen!vennameeng
       Combo5.AddItem rstVen!Vennameeng
       rstVen.MoveNext
       Loop
    End If

End Sub

Private Sub cmdExit1_Click()
If cmdExit1.caption = "E&xit" Then  'This is to Exit From form
Unload Me

Else    'This is to Cancel the Job

FrmPayableSetup.CmdEdit.caption = Trim("E&dit")
FrmPayableSetup.ListView1.ListItems.clear
FrmPayableSetup.ListView2.ListItems.clear

'frmMenu.sEdit.Caption = "Edit" 'Once Press Cancel this Caption Agian will be "Edit"
'frmMenu.edit.Caption = "Edit" 'Once Press Cancel this Caption Agian will be "Edit"

cmdNew.caption = "&New"
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
 CmbPrepBy.Enabled = False
 CmbSource.Enabled = False
 Combo5.Enabled = False
txtStorEntryNo.Enabled = False
mskStoreEnDate.Enabled = False
'txtvenCode.Enabled = False
cmbProfCenter.Enabled = False

 CmbApprovedBy.Text = ""
 cmbCostCenter.Text = ""
 CmbNotedBy.Text = ""
 CmbPrepBy.Text = ""
 CmbSource.Text = ""
 'Combo5.text=""
cmbProfCenter.Text = ""
List1.Text = ""

For Each Control In Me
    If TypeOf Control Is TextBox Then
        Control.Text = ""
    End If
Next

Me.ListView2.ListItems.clear

TxtItemDes.Text = ""
txtItemCode.Text = ""

frmMenu.PurchEdit.Enabled = True
End If

End Sub

Private Sub cmdNew_Click()
Dim JOurnalNo As ADODB.Recordset
Set JOurnalNo = New ADODB.Recordset

If cmdNew.caption = "&New" Then

'This is to Add New Records
cmdNew.caption = "&Save"
cmdExit1.caption = "&Cancel"
txtAmountDue.Enabled = True
 'txtApprBy.Enabled = True
 txtBranch.Enabled = True
 txtDocuNo.Enabled = True
 mskDateDue.Enabled = True
' txtNotedBy.Enabled = True
 'txtPrepBy.Enabled = True
 txtRefNo.Enabled = True
 CmbApprovedBy.Enabled = True
 cmbCostCenter.Enabled = True
 CmbNotedBy.Enabled = True
 CmbPrepBy.Enabled = True
 CmbSource.Enabled = True
 'Combo5.Enabled = True
cmbProfCenter.Enabled = True
txtPoNo.Enabled = True

MskDate.Enabled = True
List1.Enabled = True
txtStorEntryNo.Enabled = True
mskStoreEnDate.Enabled = True
'txtvenCode.Enabled = True

'Dim strCount
'strCount = 0
'RstPurchaseSetup.MoveFirst
'While RstPurchaseSetup.EOF = False
''strCount = RstPurchaseSetup!serialno
'    If strCount < Trim(RstPurchaseSetup!serialno) Then
'  strCount = Trim(RstPurchaseSetup!serialno)
'    End If
' RstPurchaseSetup.MoveNext
' Wend
'
'
''This is to increase the serial number one by one
'txtSerialNo.Text = Trim(strCount) + 1


'-0-0-0-
JOurnalNo.Open "Setup", constring, adOpenKeyset, adLockPessimistic, adCmdTable


If Left(JOurnalNo!CurrentMoYr, 2) <> Format(Date, "mm") Then
   JOurnalNo!CurrentMoYr = Format(Date, "mm") & Trim(Right(Year(Date), 2))
   JOurnalNo!nextjn = "00001"
   JOurnalNo.Update
Else
   Jn = "PC" & "-" & Right(Year(Date), 2) & Trim(Format(Date, "mm")) & "-" & JOurnalNo!nextjn
   nextjn = Val(JOurnalNo!nextjn)
   If (Val(nextjn)) = 1 Then
    Zeros = "0000"
    ElseIf Len(nextjn) = 2 Then
    Zeros = "000"
    ElseIf Len(nextjn) = 3 Then
    Zeros = "00"
    ElseIf Len(nextjn) = 4 Then
    Zeros = "0"
    ElseIf Len(nextjn) = 5 Then
    Zeros = ""
   End If
   JOurnalNo!nextjn = Zeros & Trim(Val(nextjn) + 1)
   JOurnalNo.Update
   JOurnalNo.close
End If
'-0-0-0
 
 Me.txtserialNo.Text = Trim(Jn)



' This is to Save the Records
 ElseIf cmdNew.caption = "&Save" Then
    If Me.cmbCostCenter.Text = "" Or Me.CmbPrepBy = "" Then
    MsgBox "There is no Sufficient Records to Save ", vbCritical, "Not Saved"
   
    Exit Sub
    End If




        X = MsgBox("Are You sure Adding this Records ?", vbYesNo, "SAVE")
        
        If X = vbNo Then
          Exit Sub
          
         End If


'This will Immediately call the Password Form once we select "YES" to save from MSG Box
frmPassword.txtBuffer = "Save"
frmPassword.txtPrepBy.Text = Me.CmbPrepBy.Text 'Useful to varify the password for him
frmPassword.Show 1
On Error Resume Next
frmPassword.txtUserId.SetFocus 'From the Password it will call SaveMe

ElseIf cmdNew.caption = "&Update" Then
frmPassword.txtBuffer = "Update"
frmPassword.txtPrepBy.Text = Me.CmbPrepBy.Text 'Useful to varify the password for him
frmPassword.Show 1
On Error Resume Next

frmPassword.txtUserId.SetFocus 'From the Password it will call SaveMe


End If

End Sub

Private Sub cmdPrint_Click()
DataReport3.Show 1
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


Public Sub DeleteMe()
Dim X As String
xyes = MsgBox("Are You Sure You Want to Delete", vbYesNo + vbQuestion, "Deleting record")
If xyes = vbYes Then

'Delete items in the ListV
        If frmPurchaseSetup.ListView1.ListItems.Count = 0 Then
        frmMenu.IDelete.Enabled = False
        Exit Sub
        End If
  xindex = frmPurchaseSetup.ListView1.SelectedItem.Index
   X = frmPurchaseSetup.ListView1.SelectedItem
     
        frmPurchaseSetup.ListView1.ListItems.Remove xindex

'Delete From the File Permanently(Put the Deletemark LATER)
    
        RstPurchaseSetup.MoveFirst
        While RstPurchaseSetup.EOF = False
        
        If X = (RstPurchaseSetup!SerialNo) Then
        RstPurchaseSetup.Delete
        
        MsgBox "Records deleted"
        
        End If
        
  RstPurchaseSetup.MoveNext
  Wend
 Unload Me
End If
 'This is the end for Deletion

End Sub


Private Sub Combo5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtDocuNo.SetFocus
End If

End Sub

Private Sub Form_Load()
Set CON1 = New ADODB.Connection

Set rstrate = New ADODB.Recordset
Set rstChart = New ADODB.Recordset

Set rstPAyFor = New ADODB.Recordset
Set rstEmp = New ADODB.Recordset
Set rstSource = New ADODB.Recordset
Set rstVen = New ADODB.Recordset
Set rstCosPro = New ADODB.Recordset
Set RstPurchaseSetup = New ADODB.Recordset
Set rstterm = New ADODB.Recordset
Set rstPayment = New ADODB.Recordset
Set rstReceipt = New ADODB.Recordset
Set RstItem = New ADODB.Recordset
Set RstINVcat = New ADODB.Recordset

conStr = "Provider=MSDASQL;DSN=Finance;UID=sa; PWD=;"


Set xcol = Me.ListView1.ColumnHeaders.Add(, , "Number", 1000)
Set xcol = Me.ListView1.ColumnHeaders.Add(, , "Date", 1100)
Set xcol = Me.ListView1.ColumnHeaders.Add(, , "Supplier", 4000)
Set xcol = Me.ListView1.ColumnHeaders.Add(, , "Account", 2500)
Set xcol = Me.ListView1.ColumnHeaders.Add(, , "Amount", 1500)

Set xcol = Me.ListView2.ColumnHeaders.Add(, , "Item code", 1150)
Set xcol = Me.ListView2.ColumnHeaders.Add(, , "Description", 2200)
Set xcol = Me.ListView2.ColumnHeaders.Add(, , "Model No", 850)
Set xcol = Me.ListView2.ColumnHeaders.Add(, , "Diamention", 950)
Set xcol = Me.ListView2.ColumnHeaders.Add(, , "Qty", 850)
Set xcol = Me.ListView2.ColumnHeaders.Add(, , "Price", 1000)
Set xcol = Me.ListView2.ColumnHeaders.Add(, , "VAT", 850)
Set xcol = Me.ListView2.ColumnHeaders.Add(, , "Sur Tax", 850)
Set xcol = Me.ListView2.ColumnHeaders.Add(, , "Tax Credit", 900)
Set xcol = Me.ListView2.ColumnHeaders.Add(, , "Inventory Inn", 850)



CON1.Open conStr
rstEmp.Open "Select * from Newemployee", CON1, adOpenDynamic, adLockOptimistic
rstSource.Open "Select * from NewForSource", CON1, adOpenDynamic, adLockOptimistic
rstVen.Open "Select * from vendor", CON1, adOpenDynamic, adLockOptimistic
rstCosPro.Open "Select * from CosProCenters", CON1, adOpenDynamic, adLockOptimistic
RstPurchaseSetup.Open "Select * from PurchaseSetup", CON1, adOpenDynamic, adLockOptimistic
RstItem.Open "Select * from Purchaseitem", CON1, adOpenDynamic, adLockOptimistic
rstChart.Open "Select * from Level6", CON1, adOpenDynamic, adLockOptimistic
RstINVcat.Open "Select * from inventorycategory", CON1, adOpenDynamic, adLockOptimistic
 
 txtAmountDue.Enabled = False
'txtApprBy.Enabled = False
 txtBranch.Enabled = False
 txtDocuNo.Enabled = False
 mskDateDue.Enabled = False
 'txtNotedBy.Enabled = False
 'txtPrepBy.Enabled = False
 txtRefNo.Enabled = False
 CmbApprovedBy.Enabled = False
 cmbCostCenter.Enabled = False
 CmbNotedBy.Enabled = False
 CmbPrepBy.Enabled = False
 CmbSource.Enabled = False
 Combo5.Enabled = False
cmbProfCenter.Enabled = False
txtPoNo.Enabled = False
MskDate.Enabled = False
List1.Enabled = False
txtStorEntryNo.Enabled = False
mskStoreEnDate.Enabled = False
'txtvenCode.Enabled = False

'this is to add Listview 1
 RstPurchaseSetup.MoveFirst
  While RstPurchaseSetup.EOF = False
  If Trim(RstPurchaseSetup!cancelledmark) = False And Trim(RstPurchaseSetup!ConfirmedMark) = False Then 'This is not Cancelled
  
     Set MItem = Me.ListView1.ListItems.Add(, , Format(RstPurchaseSetup!SerialNo))
     MItem.SubItems(1) = Format(RstPurchaseSetup!Xdate)
     'mitem.SubItems(2) = Format(RstPurchaseSetup!payto)
     MItem.SubItems(3) = Format(RstPurchaseSetup!DocNo)
     MItem.SubItems(4) = Format(RstPurchaseSetup!outbal)
        
       Totlist1 = Val(Totlist1) + Val(Trim(RstPurchaseSetup!outbal)) 'This is for the Total of the List
 End If
     RstPurchaseSetup.MoveNext
     Wend
'txtTotList1.Text = Trim(Totlist1)





''this is to add Listview2
' RstItem.MoveFirst
'  While RstItem.EOF = False
' ' If Trim(RstItem!cancelledmark) = False And Trim(RstItem!ConfirmedMark) = True Then 'This is not Cancelled
'
'     Set mitem = Me.ListView2.ListItems.Add(, , Format(RstItem!itemcode))
'     mitem.SubItems(1) = Format(RstItem!itemdesc)
'     mitem.SubItems(2) = Format(RstItem!itemmodelno)
'     mitem.SubItems(3) = Format(RstItem!itemdiamention)
'     mitem.SubItems(4) = Format(RstItem!itemqty)
'     mitem.SubItems(5) = Format(RstItem!itemprice)
'     mitem.SubItems(6) = Format(RstItem!surtax)
'     mitem.SubItems(7) = Format(RstItem!vat)
'     mitem.SubItems(8) = Format(RstItem!taxcredit)
'     mitem.SubItems(9) = Format(RstItem!inventoryinn)
'
'
'       TotList2 = Val(TotList2) + Val(Trim(RstItem!totalInventoryInn)) 'This is for the Total of the List
' 'End If
'     RstItem.MoveNext
'     Wend
'txtTotalInvInn.Text = Trim(TotList2)





rstSource.MoveFirst
Do Until rstSource.EOF
       CmbSource.AddItem rstSource!Name
   rstSource.MoveNext
  Loop

rstEmp.MoveFirst
Do Until rstEmp.EOF
       CmbPrepBy.AddItem rstEmp!Name
       CmbNotedBy.AddItem rstEmp!Name
       CmbApprovedBy.AddItem rstEmp!Name
   rstEmp.MoveNext
  Loop

rstCosPro.MoveFirst
Do Until rstCosPro.EOF
        If serial <> rstCosPro!profitcenter Then
       cmbProfCenter.AddItem rstCosPro!profitcenter
       End If
        serial = rstCosPro!profitcenter
         rstCosPro.MoveNext
  Loop


'rstChart.MoveFirst
'Do Until rstChart.EOF
'       txtItemCode.AddItem rstChart!Accountcode
'       TxtItemDes.AddItem rstChart!Accountnameeng
'         rstChart.MoveNext
'Loop
  
RstINVcat.MoveFirst
Do Until RstINVcat.EOF
       txtItemCode.AddItem RstINVcat!iccode
       TxtItemDes.AddItem RstINVcat!icnameeng
         RstINVcat.MoveNext
Loop
  
  
  
End Sub




Private Sub List1_Click()
If KeyAscii = 13 Then
txtNoOfAtt.SetFocus
End If

End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)

cmdNew.caption = "&Save"
cmdExit1.caption = "&Cancel"
txtAmountDue.Enabled = True
 'txtApprBy.Enabled = True
 txtBranch.Enabled = True
 txtDocuNo.Enabled = True
 mskDateDue.Enabled = True
' txtNotedBy.Enabled = True
 'txtPrepBy.Enabled = True
 txtRefNo.Enabled = True
 CmbApprovedBy.Enabled = True
 cmbCostCenter.Enabled = True
 CmbNotedBy.Enabled = True
 CmbPrepBy.Enabled = True
 CmbSource.Enabled = True
 'Combo5.Enabled = True
cmbProfCenter.Enabled = True
txtPoNo.Enabled = True

MskDate.Enabled = True
List1.Enabled = True
txtStorEntryNo.Enabled = True
mskStoreEnDate.Enabled = True
txtvenCode.Enabled = True


If cmdNew.caption = "&Save" Or cmdNew.caption = "&Update" Then
If Button = 2 Then
  PopupMenu frmMenu.PurcSetup
End If
End If

End Sub


Private Sub ListView2_Click()
        
      If ListView2.ListItems.Count = 0 Then
      Exit Sub
      End If
        
        
        
        frmPurchaseSetup.txtItemCode.Text = frmPurchaseSetup.ListView2.SelectedItem.Text
        frmPurchaseSetup.TxtItemDes.Text = frmPurchaseSetup.ListView2.SelectedItem.SubItems(1)
        frmPurchaseSetup.txtItemModelNo.Text = frmPurchaseSetup.ListView2.SelectedItem.SubItems(2)
        frmPurchaseSetup.txtItemDiam.Text = frmPurchaseSetup.ListView2.SelectedItem.SubItems(3)
        frmPurchaseSetup.txtItemQty.Text = frmPurchaseSetup.ListView2.SelectedItem.SubItems(4)
        frmPurchaseSetup.txtItemPrize.Text = frmPurchaseSetup.ListView2.SelectedItem.SubItems(5)
        frmPurchaseSetup.txtSurTax.Text = frmPurchaseSetup.ListView2.SelectedItem.SubItems(6)
        frmPurchaseSetup.txtVAt.Text = frmPurchaseSetup.ListView2.SelectedItem.SubItems(7)
        frmPurchaseSetup.txtTaxCredit.Text = frmPurchaseSetup.ListView2.SelectedItem.SubItems(8)
        frmPurchaseSetup.txtItemInvInn.Text = frmPurchaseSetup.ListView2.SelectedItem.SubItems(9)

        
        frmMenu.IEdit.caption = "Update"
        frmMenu.IClear.caption = "Cancel" 'This will Enable Internal Edit Again"

frmPurchaseSetup.txtItemTemp1List1.Text = frmPurchaseSetup.txtItemCode.Text
'frmPurchaseSetup.txtTemp2List3.Text = frmPurchaseSetup.txtPartic.Text

End Sub

Private Sub ListView2_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
cmdNew.caption = "&Save"
cmdExit1.caption = "&Cancel"
txtAmountDue.Enabled = True
 'txtApprBy.Enabled = True
 txtBranch.Enabled = True
 txtDocuNo.Enabled = True
 mskDateDue.Enabled = True
' txtNotedBy.Enabled = True
 'txtPrepBy.Enabled = True
 txtRefNo.Enabled = True
 CmbApprovedBy.Enabled = True
 cmbCostCenter.Enabled = True
 CmbNotedBy.Enabled = True
 CmbPrepBy.Enabled = True
 CmbSource.Enabled = True
 'Combo5.Enabled = True
cmbProfCenter.Enabled = True
txtPoNo.Enabled = True

MskDate.Enabled = True
List1.Enabled = True
txtStorEntryNo.Enabled = True
mskStoreEnDate.Enabled = True
txtvenCode.Enabled = True



If cmdNew.caption = "&Save" Or cmdNew.caption = "&Update" Then
If Button = 2 Then
  PopupMenu frmMenu.xItem
End If
End If
End Sub

Private Sub MskDate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
CmbSource.SetFocus
End If

End Sub

Private Sub mskDateDue_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtPoNo.SetFocus
End If

End Sub

Private Sub mskStoreEnDate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtStorEntryNo.SetFocus
End If

End Sub

Private Sub txtAmtPaid_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtProTax.SetFocus
End If

End Sub

Private Sub txtAmtReq_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtAmtPaid.SetFocus
End If

End Sub

Private Sub txtBranch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
List1.SetFocus
End If

End Sub

Private Sub txtDocuNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtvenCode.SetFocus
End If

End Sub

Private Sub txtItemCode_Click()
RstINVcat.MoveFirst
While RstINVcat.EOF = False
If txtItemCode.Text = RstINVcat!iccode Then
TxtItemDes.Text = RstINVcat!icnameeng
End If
RstINVcat.MoveNext
Wend

End Sub

Private Sub txtItemCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
TxtItemDes.SetFocus
End If
End Sub

Private Sub txtItemCode_LostFocus()
rstChart.MoveFirst
While rstChart.EOF = False
If txtItemCode.Text = rstChart!AccountCode Then
TxtItemDes.Text = rstChart!accountnameeng
End If
rstChart.MoveNext
Wend

End Sub

Private Sub TxtItemDes_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
txtItemModelNo.SetFocus
End If

End Sub

Private Sub txtItemDiam_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
txtItemQty.SetFocus
End If

End Sub

Private Sub txtItemInvInn_KeyPress(KeyAscii As Integer)
'This will add the  Deatails in the Textboxes to  the Listview when i press the enter key

If KeyAscii = 13 Then

     Set MItem = Me.ListView2.ListItems.Add(, , Trim(txtItemCode.Text))
     MItem.SubItems(1) = Trim(TxtItemDes.Text)
     MItem.SubItems(2) = Trim(txtItemModelNo.Text)
     MItem.SubItems(3) = Format(txtItemDiam.Text)
     MItem.SubItems(4) = Format(txtItemQty.Text)
     MItem.SubItems(5) = Format(txtItemPrize.Text)
     MItem.SubItems(6) = Format(txtSurTax.Text)
     MItem.SubItems(7) = Format(txtVAt.Text)
     MItem.SubItems(8) = Format(txtTaxCredit.Text)
     MItem.SubItems(9) = Format(txtItemInvInn.Text)




'TotList3 = Val(txtTotList3) + Val(Trim(txtAmo.Text)) 'This is for the Total of the List
     txtItemCode.Text = ""
      TxtItemDes.Text = ""
      txtItemModelNo.Text = ""
      txtItemDiam.Text = ""
      txtItemQty.Text = ""
      txtItemPrize.Text = ""
      txtSurTax.Text = ""
      txtVAt.Text = ""
      txtTaxCredit.Text = ""
      txtItemInvInn.Text = ""
      
      

End If

End Sub

Private Sub txtItemModelNo_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
txtItemDiam.SetFocus
End If

End Sub

Private Sub txtItemPrize_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
txtVAt.SetFocus
End If
End Sub


Private Sub txtItemQty_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
txtItemPrize.SetFocus
End If

End Sub

Private Sub txtNoOfAtt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbProfCenter.SetFocus
End If

End Sub

Private Sub txtOutBal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtPercentage.SetFocus
End If

End Sub

Private Sub txtPercentage_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtRefNo.SetFocus
End If

End Sub

Private Sub txtPoNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
mskStoreEnDate.SetFocus
End If

End Sub

Private Sub txtProTax_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtOutBal.SetFocus
End If

End Sub

Private Sub txtRefNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtAmountDue.SetFocus
End If

End Sub

Private Sub txtSearch1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmdSearch_Click
End If
End Sub

Private Sub txtserialNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
MskDate.SetFocus
End If

End Sub

Private Sub txtStorEntryNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtAmtReq.SetFocus
End If

End Sub

Private Sub txtSurTax_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
txtTaxCredit.SetFocus
End If

End Sub

Private Sub txtTaxCredit_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtItemInvInn.SetFocus
End If

End Sub

Private Sub txtVAt_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
txtSurTax.SetFocus
End If

End Sub

Public Sub UpdateList2Item()

   RstItem.MoveFirst
      While RstItem.EOF = False
       Dim xvar As String
      ' xvar = Trim(rstItem!accno)
         If Trim(RstItem!SerialNo) = Trim(frmPurchaseSetup.txtserialNo.Text) And Trim(RstItem!itemcode) = Trim(frmPurchaseSetup.txtItemTemp1List1.Text) Then
        RstItem!itemcode = frmPurchaseSetup.txtItemCode.Text
       
RstItem!itemdesc = frmPurchaseSetup.TxtItemDes.Text
RstItem!itemmodelno = frmPurchaseSetup.txtItemModelNo.Text
RstItem!itemdiamention = frmPurchaseSetup.txtItemDiam.Text
RstItem!itemqty = frmPurchaseSetup.txtItemQty.Text
RstItem!itemprice = frmPurchaseSetup.txtItemPrize.Text
RstItem!SurTax = frmPurchaseSetup.txtSurTax.Text
RstItem!vat = frmPurchaseSetup.txtVAt.Text
RstItem!taxcredit = frmPurchaseSetup.txtTaxCredit.Text
RstItem!inventoryinn = frmPurchaseSetup.txtItemInvInn.Text
        
        Myvalitem = "Gotit"
      End If
   RstItem.MoveNext
   Wend
If Myvalitem = "Gotit" Then
MsgBox "Records for Payment Details Updated Seccussfully"
End If



'Refresh ListView 2
''..This is to REFRESH the ListView1, after updating the Details
   frmPurchaseSetup.ListView2.ListItems.clear
'this is to add Listview2
 RstItem.MoveFirst
  While RstItem.EOF = False
  If Trim(RstItem!SerialNo) = Me.txtserialNo.Text Then
     Set MItem = Me.ListView2.ListItems.Add(, , Format(RstItem!itemcode))
     MItem.SubItems(1) = Format(RstItem!itemdesc)
     MItem.SubItems(2) = Format(RstItem!itemmodelno)
     MItem.SubItems(3) = Format(RstItem!itemdiamention)
     MItem.SubItems(4) = Format(RstItem!itemqty)
     MItem.SubItems(5) = Format(RstItem!itemprice)
     MItem.SubItems(6) = Format(RstItem!SurTax)
     MItem.SubItems(7) = Format(RstItem!vat)
     MItem.SubItems(8) = Format(RstItem!taxcredit)
     MItem.SubItems(9) = Format(RstItem!inventoryinn)


       TotList2 = Val(TotList2) + Val(Trim(RstItem!totalInventoryInn)) 'This is for the Total of the List
 End If
     RstItem.MoveNext
     Wend
txtTotalInvInn.Text = Trim(TotList2)
End Sub

Private Sub txtvenCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtBranch.SetFocus
End If

End Sub

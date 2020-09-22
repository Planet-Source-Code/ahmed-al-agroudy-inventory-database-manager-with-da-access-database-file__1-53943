VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form AssetSetup 
   Caption         =   "Asset Setup"
   ClientHeight    =   6135
   ClientLeft      =   720
   ClientTop       =   1500
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check1 
      Caption         =   "Set to Arabic"
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
      Left            =   10500
      TabIndex        =   76
      Top             =   600
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1200
      ScaleHeight     =   225
      ScaleWidth      =   3195
      TabIndex        =   64
      Top             =   5280
      Visible         =   0   'False
      Width           =   3255
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   0
         TabIndex        =   65
         Top             =   0
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Left            =   2760
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.Timer Timer2 
      Left            =   3960
      Top             =   5520
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4080
      Top             =   5400
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Ne&xt ÇáÊÇáí >>"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   4680
      TabIndex        =   59
      Top             =   5280
      Width           =   1335
   End
   Begin VB.ComboBox Combo19 
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
      Left            =   6480
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   120
      Width           =   3550
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Apply ãæÇÝÞ"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   9120
      TabIndex        =   62
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel ÇáÛÇÁ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   7800
      TabIndex        =   61
      Top             =   5280
      Width           =   1215
   End
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
      Left            =   1800
      Style           =   1  'Simple Combo
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4575
      Left            =   75
      TabIndex        =   5
      Top             =   600
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   8070
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   617
      TabMaxWidth     =   4410
      BackColor       =   -2147483637
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Asset Info ãÚáæãÇÊ ÇáÃÕá"
      TabPicture(0)   =   "FixedAssetSetup.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label15"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label16"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label17"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label18"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label20"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label24"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label25"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label26"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label27"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label28"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label41"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Combo14"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Combo15"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Combo16"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Combo17"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Combo18"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Combo20"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "Item User ÇÎÊíÇÑ ÇáãÓÊÎÏã "
      TabPicture(1)   =   "FixedAssetSetup.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame8"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Acctg Entry ÊÓÌíá ÇáÇÕá"
      TabPicture(2)   =   "FixedAssetSetup.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame3"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame5"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Frame2"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Frame1"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Frame6"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Frame7"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).ControlCount=   7
      TabCaption(3)   =   "List ÞÇÆãÉ"
      TabPicture(3)   =   "FixedAssetSetup.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "ListView1"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VB.ComboBox Combo20 
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
         Height          =   315
         Left            =   1560
         TabIndex        =   88
         Top             =   1560
         Width           =   4935
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4110
         Left            =   -74940
         TabIndex        =   58
         Top             =   405
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   7250
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList2"
         SmallIcons      =   "ImageList2"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   18
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "a"
            Text            =   "AssetID"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "b"
            Text            =   "AssetCode"
            Object.Width           =   3351
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Key             =   "c"
            Text            =   "AssetName"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Key             =   "d"
            Text            =   "ModelNo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Key             =   "e"
            Text            =   "SerialNo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Key             =   "f"
            Text            =   "AssetType"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Key             =   "g"
            Text            =   "AcquiredDate"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Key             =   "h"
            Text            =   "AcquiredValue"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Key             =   "i"
            Text            =   "SalvageValue"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Key             =   "j"
            Text            =   "ComputationMethod"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Key             =   "k"
            Text            =   "UseFullLife"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   11
            Key             =   "l"
            Text            =   "YearlyFixedRate"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Key             =   "m"
            Text            =   "ProductionUnit"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   13
            Key             =   "n"
            Text            =   "AccumulatedDep"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Key             =   "o"
            Text            =   "DebitAccount"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Key             =   "p"
            Text            =   "CreditAcct"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Key             =   "q"
            Text            =   "Assign To"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.ComboBox Combo18 
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
         Left            =   1560
         Style           =   1  'Simple Combo
         TabIndex        =   15
         Top             =   2640
         Width           =   7815
      End
      Begin VB.ComboBox Combo17 
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
         Left            =   1560
         Locked          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   13
         Top             =   2280
         Width           =   2655
      End
      Begin VB.ComboBox Combo16 
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
         Left            =   1560
         Locked          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   11
         Top             =   1920
         Width           =   2655
      End
      Begin VB.ComboBox Combo15 
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
         Height          =   315
         Left            =   1560
         TabIndex        =   9
         Top             =   1200
         Width           =   4935
      End
      Begin VB.ComboBox Combo14 
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
         Height          =   315
         Left            =   1560
         TabIndex        =   7
         Top             =   840
         Width           =   2655
      End
      Begin VB.Frame Frame8 
         Caption         =   "Assign to"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3735
         Left            =   -74760
         TabIndex        =   16
         Top             =   480
         Width           =   8800
         Begin VB.ComboBox Combo13 
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
            Left            =   1560
            Style           =   1  'Simple Combo
            TabIndex        =   22
            Top             =   1200
            Width           =   4215
         End
         Begin VB.ComboBox Combo12 
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
            Left            =   1560
            Style           =   1  'Simple Combo
            TabIndex        =   20
            Top             =   840
            Width           =   4215
         End
         Begin VB.ComboBox Combo11 
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
            Left            =   1560
            Style           =   1  'Simple Combo
            TabIndex        =   18
            Top             =   480
            Width           =   2535
         End
         Begin VB.Label Label31 
            Caption         =   "ÇáæÍÏÉ"
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
            Left            =   5880
            TabIndex        =   75
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label30 
            Caption         =   "ÇáÇÏÇÑÉ "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5880
            TabIndex        =   74
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label29 
            Caption         =   "ÇáÞÓã "
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
            Left            =   4200
            TabIndex        =   73
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label Label14 
            Caption         =   "&Unit/Section"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   21
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label13 
            Caption         =   "&Department"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   19
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label12 
            Caption         =   "Di&vision"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   17
            Top             =   480
            Width           =   1095
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Accumulated Depreciation ÇáÇÓÊåáÇß ÇáãÊÑÇßã "
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   -68280
         TabIndex        =   53
         Top             =   3120
         Width           =   4815
         Begin VB.ComboBox Combo10 
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
            Left            =   960
            Locked          =   -1  'True
            Style           =   1  'Simple Combo
            TabIndex        =   57
            Top             =   600
            Width           =   2655
         End
         Begin VB.ComboBox Combo9 
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
            Left            =   960
            Locked          =   -1  'True
            Style           =   1  'Simple Combo
            TabIndex        =   55
            Top             =   240
            Width           =   2655
         End
         Begin VB.Label Label40 
            Alignment       =   1  'Right Justify
            Caption         =   " ÇÓã ÇáÍÓÇÈ"
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
            Left            =   3480
            TabIndex        =   86
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label39 
            Alignment       =   1  'Right Justify
            Caption         =   " ÑÞã ÇáÍÓÇÈ "
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
            Left            =   3360
            TabIndex        =   85
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label11 
            Caption         =   "Acct Name"
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
            TabIndex        =   56
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label10 
            Caption         =   "Acct. No."
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
            TabIndex        =   54
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Depreciation Expense ãÕÑæÝÇÊ ÇáÇÓÊåáÇß"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   -73080
         TabIndex        =   48
         Top             =   3120
         Width           =   4695
         Begin VB.ComboBox Combo8 
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
            TabIndex        =   52
            Top             =   600
            Width           =   2535
         End
         Begin VB.ComboBox Combo2 
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
            TabIndex        =   50
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label Label38 
            Caption         =   "ÇÓã ÇáÍÓÇÈ "
            Height          =   255
            Left            =   3600
            TabIndex        =   84
            Top             =   600
            Width           =   1000
         End
         Begin VB.Label Label37 
            Caption         =   " ÑÞã ÇáÍÓÇÈ "
            Height          =   255
            Left            =   3600
            TabIndex        =   83
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label9 
            Caption         =   "Acct Name"
            Height          =   450
            Left            =   120
            TabIndex        =   51
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label8 
            Caption         =   "Acct. No."
            Height          =   495
            Left            =   120
            TabIndex        =   49
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label6 
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
            Left            =   240
            TabIndex        =   63
            Top             =   480
            Width           =   2055
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Types of Asset"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   -74760
         TabIndex        =   23
         Top             =   480
         Width           =   2295
         Begin VB.OptionButton Option1 
            Caption         =   "&Depreciable  ÇÓÊåáÇß"
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
            TabIndex        =   24
            Top             =   280
            Value           =   -1  'True
            Width           =   2055
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Amorti&zable ÇØÝÇÁ"
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
            TabIndex        =   25
            Top             =   600
            Width           =   2055
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Computation Methods"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   -74760
         TabIndex        =   26
         Top             =   1680
         Width           =   2295
         Begin VB.OptionButton Option3 
            Caption         =   "&Straight Line ÇáÞÓØ ÇáËÇÈÊ"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   280
            Value           =   -1  'True
            Width           =   2175
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Diminis&hing Rate ÇáãÚÏá ÇáãÊäÇÞÕ  "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   28
            Top             =   550
            Width           =   2055
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Activity &Base  Úáì ÇÓÇÓ ÇáäÔÇØ"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   29
            Top             =   960
            Width           =   2055
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Activity Base Unit æÍÏÉ ÞíÇÓ ÇáäÔÇØ "
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   -67800
         TabIndex        =   42
         Top             =   480
         Width           =   4335
         Begin VB.CommandButton Command6 
            Caption         =   "Data &Entry...ÏÎæá ÇáãÚáæãÇÊ"
            Enabled         =   0   'False
            Height          =   375
            Left            =   1800
            TabIndex        =   47
            Top             =   1920
            Width           =   2175
         End
         Begin VB.ComboBox Combo5 
            Enabled         =   0   'False
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
            Left            =   720
            TabIndex        =   46
            Top             =   1200
            Width           =   2055
         End
         Begin VB.OptionButton Option6 
            Caption         =   "by &Productivity Unit ÈæÇÓØÉ æÍÏÉ ÇáãäÊÌ  "
            Enabled         =   0   'False
            Height          =   495
            Left            =   2040
            TabIndex        =   44
            Top             =   360
            Width           =   1815
         End
         Begin VB.OptionButton Option7 
            Caption         =   "by Machine &Hours ÈæÇÓØÉ ÓÇÚÇÊ ÇáÚãá"
            Enabled         =   0   'False
            Height          =   615
            Left            =   240
            TabIndex        =   43
            Top             =   360
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.Label Label36 
            Caption         =   "æÍÏÉ"
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
            Left            =   2880
            TabIndex        =   82
            Top             =   1200
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "&Units"
            Enabled         =   0   'False
            Height          =   375
            Left            =   240
            TabIndex        =   45
            Top             =   1200
            Width           =   615
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Items"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   -72360
         TabIndex        =   32
         Top             =   480
         Width           =   4455
         Begin VB.ComboBox Combo7 
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
            Left            =   1680
            Style           =   1  'Simple Combo
            TabIndex        =   40
            Top             =   1200
            Width           =   1335
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Additional Cost...ÊßÇáíÝ ÇÖÇÝíÉ"
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
            Left            =   1920
            TabIndex        =   41
            Top             =   1920
            Width           =   2295
         End
         Begin VB.ComboBox Combo3 
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
            Left            =   1680
            Style           =   1  'Simple Combo
            TabIndex        =   36
            Top             =   480
            Width           =   1335
         End
         Begin VB.ComboBox Combo4 
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
            Left            =   1680
            Style           =   1  'Simple Combo
            TabIndex        =   38
            Top             =   840
            Width           =   1335
         End
         Begin MSMask.MaskEdBox MaskEdBox1 
            Height          =   330
            Left            =   1680
            TabIndex        =   34
            Top             =   120
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label35 
            Alignment       =   1  'Right Justify
            Caption         =   "ÇáÚãÑ ÇáÇäÊÇÌí "
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
            Left            =   2760
            TabIndex        =   81
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Label Label34 
            Alignment       =   1  'Right Justify
            Caption         =   "ÇáÞíãÉ ÇáÊÎÑíÏíÉ"
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
            Left            =   2760
            TabIndex        =   80
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            Caption         =   "ÇáßáÝÉ ÇáÊÇÑíÎíÉ "
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
            Left            =   2760
            TabIndex        =   79
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            Caption         =   " ÊÇÑíÎ ÇáÍíÇÒÉ"
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
            Left            =   2760
            TabIndex        =   78
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label7 
            Caption         =   "Usefull Life in Years"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   39
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label Label2 
            Caption         =   "Acquisiti&on Date"
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
            TabIndex        =   33
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label3 
            Caption         =   "Historical Value"
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
            TabIndex        =   35
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label4 
            Caption         =   "Salvage V&alue"
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
            TabIndex        =   37
            Top             =   960
            Width           =   1095
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Yrly Fixed Rate%"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   -74760
         TabIndex        =   30
         Top             =   3120
         Width           =   1575
         Begin VB.ComboBox Combo6 
            Enabled         =   0   'False
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
            Style           =   1  'Simple Combo
            TabIndex        =   31
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label22 
            Caption         =   "ÇáãÚÏá ÇáÓäæí ÇáËÇÈÊ"
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
            TabIndex        =   77
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
         Caption         =   "Item Name Arab"
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
         TabIndex        =   87
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label28 
         Caption         =   "ãáÇÍÙÇÊ"
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
         Left            =   9480
         TabIndex        =   72
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label27 
         Caption         =   "ÑÞã ÇáÊÓáÓá"
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
         Left            =   4440
         TabIndex        =   71
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label26 
         Caption         =   "ÑÞã ÇáãæÏíá "
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
         Left            =   4320
         TabIndex        =   70
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label25 
         Caption         =   "ÇÓã ÇáÕäÝ"
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
         Left            =   6600
         TabIndex        =   69
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label24 
         Caption         =   "ÑÞã ÇáÕäÝ"
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
         Left            =   4320
         TabIndex        =   68
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "Remark/Usage"
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
         TabIndex        =   14
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "&Serial No."
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
         TabIndex        =   12
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "&Model No."
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
         TabIndex        =   10
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "Item &Name Eng"
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
         Left            =   240
         TabIndex        =   8
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "&Item Code"
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
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   6480
      TabIndex        =   60
      Top             =   5280
      Width           =   1215
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   0
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FixedAssetSetup.frx":0070
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FixedAssetSetup.frx":04C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FixedAssetSetup.frx":0914
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FixedAssetSetup.frx":0D66
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FixedAssetSetup.frx":1080
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label23 
      Caption         =   "ÇÎÊÇÑ ãÌãæÚÉ ÇáãÎÒæä "
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
      Left            =   10200
      TabIndex        =   67
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label21 
      Caption         =   "Generating..."
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
      TabIndex        =   66
      Top             =   5320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label19 
      Caption         =   "Select Inventory Ca&tegory"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   3
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Asset &No ÑÞã ÇáÃÕá"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Menu Main 
      Caption         =   "Main"
      Visible         =   0   'False
      Begin VB.Menu SEtupNew 
         Caption         =   "Setup New Item"
      End
      Begin VB.Menu xedit 
         Caption         =   "Edit"
      End
      Begin VB.Menu xDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu bar1 
         Caption         =   "-"
      End
      Begin VB.Menu xGenerate 
         Caption         =   "&Generate Assets Journal"
      End
      Begin VB.Menu xMachineUsed 
         Caption         =   "Set Machine Hour Used..."
         Enabled         =   0   'False
      End
      Begin VB.Menu xFind 
         Caption         =   "&Find"
      End
      Begin VB.Menu AssetValuation 
         Caption         =   "Asset Valuation Statement"
      End
   End
End
Attribute VB_Name = "AssetSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstDEpAsset As ADODB.Recordset
Dim CON1 As New ADODB.Connection
Dim rsAssetNAmeArab As New ADODB.Recordset
Dim find As Boolean
Dim GotFocus As Boolean
Dim MItem As ListItem
Dim xPrevTAb As Integer
Dim xSeek As Boolean
Dim LoadSeriesNo As Boolean
Dim xdecimal As Integer
Dim rstEditRec As New ADODB.Recordset
Dim rstAsset As New ADODB.Recordset
Sub ClickTab()
    If Left(Trim(Me.SSTab1.caption), 10) = "Asset Info" Then
        Me.SSTab1.SetFocus
        SendKeys "{Left}"
        Me.Command7.Enabled = True
     ElseIf Left(Trim(Me.SSTab1.caption), 9) = "Item User" Then
        Me.SSTab1.SetFocus
        SendKeys "{Left}"
        SendKeys "{Right}"
        Me.Timer1.Enabled = False
        Me.Command7.Enabled = False
    ElseIf Left(Trim(Me.SSTab1.caption), 11) = "Acctg Entry" Then
        Me.SSTab1.SetFocus
        SendKeys "{Right}"
        Me.Timer1.Enabled = False
        Me.Command7.Enabled = False
    End If

End Sub
Sub ListIt()
Me.ListView1.ListItems.clear
On Error Resume Next
rstAsset.close
On Error GoTo 0
rstAsset.Open "Select * from AssetSetup order by Assetno", constring, adOpenKeyset, adLockPessimistic, adCmdText
Do Until rstAsset.EOF = True
    Set MItem = Me.ListView1.ListItems.Add(, , rstAsset!ASsetNo, , IIf(rstAsset!AssetType = "Depreciable", 4, 1))
    MItem.SubItems(1) = rstAsset!AssetCode
    MItem.SubItems(2) = rstAsset!AssetName
    MItem.SubItems(3) = rstAsset!Modelno
    MItem.SubItems(4) = rstAsset!SerialNo
    MItem.SubItems(5) = rstAsset!AssetType
    MItem.SubItems(6) = Format(rstAsset!AcquisitionDate, "dd/mm/yyyy")
    MItem.SubItems(7) = FormatNumber(rstAsset!AcquisitionValue, 2, vbTrue, vbTrue, vbTrue)
    MItem.SubItems(8) = FormatNumber(rstAsset!SalvageValue, 2, vbTrue, vbTrue, vbTrue)
    MItem.SubItems(9) = rstAsset!COmputationMethod
    MItem.SubItems(10) = rstAsset!UsefullLife
    MItem.SubItems(11) = rstAsset!YearlyFixedRate
    'mItem.SubItems(12) = 'rstAsset!MachineHour
    'mItem.SubItems(13) = 'rstAsset!UnitPEriod
    MItem.SubItems(12) = rstAsset!Unitprod
    MItem.SubItems(13) = FormatNumber(rstAsset!AccumulatedDep, 2, vbTrue, vbTrue, vbTrue)
    MItem.SubItems(14) = rstAsset!DEbitAcct
    MItem.SubItems(15) = rstAsset!CreditAcct
    MItem.SubItems(16) = rstAsset!AssignTo
    rstAsset.MoveNext
Loop
rstAsset.close

End Sub

Private Sub AssetValuation_Click()
On Error Resume Next
FinanceDE.rsSetupAsset.close
ASsetValuationRep.Show
End Sub

Private Sub Check1_Click()
On Error Resume Next
If Me.Check1.Value = 1 Then
  For Each Control In Me
     If TypeOf Control Is ComboBox Or TypeOf Control Is MaskEdBox Then
        Control.RightToLeft = True
     End If
  Next
 Else
 For Each Control In Me
  If TypeOf Control Is ComboBox Or TypeOf Control Is MaskEdBox Then
     Control.RightToLeft = False
  End If
  Next
End If

End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
   LoadSeriesNo = False
   Me.Combo1 = Val(Me.Combo1) + 1
   'If Trim(Me.SSTab1.Caption) = "List of Assets Setup" Then
        i = 0
        For i = 1 To Me.ListView1.ListItems.Count
          Me.ListView1.ListItems.Item(i).ForeColor = vbBlack
          Me.ListView1.ListItems.Item(i).Ghosted = False
          Me.ListView1.ListItems.Item(i).Bold = False
        Next
        Dim strFindMe As String
        Dim itmFound As ListItem   ' FoundItem variable.
        intSelectedOption = lvwText
        strFindMe = Trim(Me.Combo1)
        Set itmFound = Me.ListView1.Finditem(strFindMe, intSelectedOption, , lvwPartial)
        If itmFound Is Nothing Then  ' If no match, inform user and exit.
           
         Else
           itmFound.EnsureVisible
           itmFound.Selected = True   ' Select the ListItem.
           Me.ListView1.SetFocus
           Me.ListView1.SelectedItem.ForeColor = vbBlue
           Me.ListView1.SelectedItem.Ghosted = True
        End If
        Me.Combo1.SetFocus
    '    Exit Sub
    ' End If
    
        'Dim rstEditRec As New ADODB.Recordset
        cItem = Trim(Me.Combo1)
        On Error Resume Next
        rstEditRec.close
        On Error GoTo 0
        rstEditRec.Open "Select * from ASsetSetup where AssetNo= " & "'" & cItem & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
        'If rstEditRec.EOF = False Then
         'rstEditRec.MoveFirst
         Me.Combo14.clear
         Me.Combo15.clear
         If rstEditRec.EOF = False Then
          If Trim(rstEditRec!ASsetNo) = cItem Then
            Me.Combo11 = rstEditRec!AssignTo
            Me.Combo14 = rstEditRec!AssetCode
            Me.Combo15 = rstEditRec!AssetName
            Me.Combo16 = rstEditRec!Modelno
            Me.Combo17 = rstEditRec!SerialNo
            Me.Combo18 = rstEditRec!remarks
            If rstEditRec!AssetType = "Depreciable" Then
                Me.Option1.Value = True
               Else
                Me.Option2.Value = True
            End If
            If Trim(rstEditRec!COmputationMethod) = "Straight Line" Then
                Me.Option3.Value = True
             ElseIf (rstEditRec!COmputationMethod) = "Diminishing Rate" Then
                Me.Option4.Value = True
                Me.Combo6 = IIf(IsNull(rstEditRec!YearlyFixedRate) = False, rstEditRec!YearlyFixedRate, " ")
             ElseIf (rstEditRec!COmputationMethod) = "Activity Base" Then
                Me.Option5.Value = True
            End If
            Me.MaskEdBox1.Text = Format(rstEditRec!AcquisitionDate, "dd/mm/yyyy")
            Me.Combo3 = FormatNumber(rstEditRec!AcquisitionValue, 2, vbTrue, vbTrue, vbTrue)
            Me.Combo4 = FormatNumber(rstEditRec!SalvageValue, 2, vbTrue, vbTrue, vbTrue)
            Me.Combo7 = Val(rstEditRec!UsefullLife)
            Me.Combo2 = Left(rstEditRec!DEbitAcct, 12)
            Me.Combo8 = Mid(rstEditRec!DEbitAcct, 14, 30)
            Me.Combo9 = Left(rstEditRec!CreditAcct, 12)
            Me.Combo10 = Mid(rstEditRec!CreditAcct, 14, 30)
            Me.Combo14.Enabled = True
            Me.Combo15.Enabled = True
            'rstEditRec.Close
            'Me.Combo1.SetFocus
            If rstEditRec.EOF = False Then
                rstEditRec.MoveLast
            End If
           End If
           rstEditRec.MoveNext
          End If
       '   Else
       '   For Each Control In Me
       '     On Error Resume Next
       '     If TypeOf Control Is ComboBox Then
       '        If Control.TabIndex > 1 Then
       '         Control.Text = ""
       '        End If
       '     End If
       '    Next
       ' End If
'        If Trim(Me.SSTab1.Caption) = "List of Assets Setup" Then
'                Me.SSTab1.SetFocus
'                SendKeys "{Right}"
'             ElseIf Trim(Me.SSTab1.Caption) = "Item User" Then
'                Me.SSTab1.SetFocus
'                SendKeys "{Left}"
'                Me.Timer1.Enabled = False
'                Me.Command7.Enabled = False
'            ElseIf Trim(Me.SSTab1.Caption) = "Accounting Entry" Then
'                Me.SSTab1.SetFocus
'                SendKeys "{Right}"
'                SendKeys "{Right}"
'                Me.Timer1.Enabled = False
'                Me.Command7.Enabled = False
'            End If
'
        End If

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
   Me.Command5.SetFocus
End If
End Sub

Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyUp Then
    Me.Combo1 = Val(Me.Combo1) - 1
   ' If Trim(Me.SSTab1.Caption) = "List of Assets Setup" Then
        i = 0
        For i = 1 To Me.ListView1.ListItems.Count
          Me.ListView1.ListItems.Item(i).ForeColor = vbBlack
          Me.ListView1.ListItems.Item(i).Ghosted = False
        Next
        Dim strFindMe As String
        Dim itmFound As ListItem   ' FoundItem variable.
        intSelectedOption = lvwText
        strFindMe = Trim(Me.Combo1)
        Set itmFound = Me.ListView1.Finditem(strFindMe, intSelectedOption, , lvwPartial)
        If itmFound Is Nothing Then  ' If no match, inform user and exit.
         Else
           itmFound.EnsureVisible
           itmFound.Selected = True   ' Select the ListItem.
           Me.ListView1.SetFocus
           Me.ListView1.SelectedItem.ForeColor = vbBlue
           Me.ListView1.SelectedItem.Ghosted = True
        End If
        Me.Combo1.SetFocus
    '    Exit Sub
    ' End If
    
    cItem = Trim(Me.Combo1)
        On Error Resume Next
        rstEditRec.close
        On Error GoTo 0
        rstEditRec.Open "Select * from ASsetSetup where AssetNo= " & "'" & cItem & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText

         'rstEditRec.MoveFirst
         Me.Combo14.clear
         Me.Combo15.clear
         If rstEditRec.EOF = False Then
            Me.Combo11 = rstEditRec!AssignTo
            Me.Combo14 = rstEditRec!AssetCode
            Me.Combo15 = rstEditRec!AssetName
            Me.Combo16 = rstEditRec!Modelno
            Me.Combo17 = rstEditRec!SerialNo
            Me.Combo18 = rstEditRec!remarks
            If rstEditRec!AssetType = "Depreciable" Then
                Me.Option1.Value = True
               Else
                Me.Option2.Value = True
            End If
            If Trim(rstEditRec!COmputationMethod) = "Straight Line" Then
                Me.Option3.Value = True
             ElseIf (rstEditRec!COmputationMethod) = "Diminishing Rate" Then
                Me.Option4.Value = True
                Me.Combo6 = IIf(IsNull(rstEditRec!YearlyFixedRate) = False, rstEditRec!YearlyFixedRate, " ")
             ElseIf (rstEditRec!COmputationMethod) = "Activity Base" Then
                Me.Option5.Value = True
            End If
            Me.MaskEdBox1.Text = Format(rstEditRec!AcquisitionDate, "dd/mm/yyyy")
            Me.Combo3 = FormatNumber(rstEditRec!AcquisitionValue, 2, vbTrue, vbTrue, vbTrue)
            Me.Combo4 = FormatNumber(rstEditRec!SalvageValue, 2, vbTrue, vbTrue, vbTrue)
            Me.Combo7 = Val(rstEditRec!UsefullLife)
            Me.Combo2 = Left(rstEditRec!DEbitAcct, 12)
            Me.Combo8 = Mid(rstEditRec!DEbitAcct, 14, 30)
            Me.Combo9 = Left(rstEditRec!CreditAcct, 12)
            Me.Combo10 = Mid(rstEditRec!CreditAcct, 14, 30)
            Me.Combo14.Enabled = True
            Me.Combo15.Enabled = True
            'rstEditRec.Close
            'Me.Combo1.SetFocus
            If rstEditRec.EOF = False Then
                rstEditRec.MoveLast
            End If
          End If
       '   Else
       '   For Each Control In Me
       '     On Error Resume Next
       '     If TypeOf Control Is ComboBox Then
       '        If Control.TabIndex > 1 Then
       '         Control.Text = ""
       '        End If
       '     End If
       '    Next
       ' End If
'    If Trim(Me.SSTab1.Caption) = "List of Assets Setup" Then
'            Me.SSTab1.SetFocus
'            SendKeys "{Right}"
'         ElseIf Trim(Me.SSTab1.Caption) = "Item User" Then
'            Me.SSTab1.SetFocus
'            SendKeys "{Left}"
'            Me.Timer1.Enabled = False
'            Me.Command7.Enabled = False
'        ElseIf Trim(Me.SSTab1.Caption) = "Accounting Entry" Then
'            Me.SSTab1.SetFocus
'            SendKeys "{Right}"
'            SendKeys "{Right}"
'            Me.Timer1.Enabled = False
'            Me.Command7.Enabled = False
'        End If

End If
End Sub

Private Sub Combo10_Click()
Dim rstDX As New ADODB.Recordset
rstDX.Open "Select * from FinanceMAster where AccountNameEng =" & "'" & Trim(Me.Combo10) & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
Me.Combo9 = rstDX!AccountCode
rstDX.close
End Sub

Private Sub Combo10_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.Command3.SetFocus
End If
End Sub

Private Sub Combo11_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.Combo12.SetFocus
End If
End Sub

Private Sub Combo12_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.Combo13.SetFocus
End If
End Sub

Private Sub Combo13_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.Command7.SetFocus
End If
End Sub

Private Sub Combo14_Click()

find = False
If Me.Combo14 <> "" Then
 itemcode = Trim(Me.Combo14)
 On Error Resume Next
 rstAsset.close
 On Error GoTo 0
 rstAsset.Open "select * from AssetRegistered where code=" & "'" & itemcode & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
 If rstAsset.EOF = False Then
     Me.Combo1 = rstAsset!Idno
     Me.Combo15 = rstAsset!NameEng
     Me.Combo20 = rstAsset!NameArab
     Me.Combo16 = IIf(IsNull(rstAsset!Modelno) <> True, rstAsset!Modelno, "None")
     Me.Combo17 = IIf(IsNull(rstAsset!SerialNo) <> True, rstAsset!SerialNo, "None")
     Me.Combo9 = rstAsset!AccumulatedCode
     Me.Combo10 = rstAsset!AccumulatedName
     find = True
     If find = True Then Exit Sub
   Else
    If find = False Then
     Me.Combo1 = ""
     Me.Combo15 = ""
     Me.Combo20 = ""
     Me.Combo16 = ""
     Me.Combo17 = ""
     Me.Combo9 = ""
     Me.Combo10 = ""
    End If
 End If
End If
End Sub

Private Sub Combo14_GotFocus()
Const CB_SHOWDROPDOWN = &H14F
Dim Tmp
If Me.Combo14.ListCount > 1 Then
    Tmp = SendMessage(Combo14.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
    SendKeys "{Down}"
End If
End Sub

Private Sub Combo14_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.Combo15.SetFocus
End If
End Sub

Private Sub Combo14_LostFocus()
Call Combo15_Click
End Sub

Private Sub Combo15_Click()
find = False
If Me.Combo15 <> "" Then
    ItemName = Trim(Me.Combo15)
    On Error Resume Next
    rstAsset.close
    On Error GoTo 0
    rstAsset.Open "select * from Assetregistered where NameEng=" & "'" & ItemName & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
    If rstAsset.EOF = False Then
      Me.Combo1 = rstAsset!Idno
      Me.Combo14 = rstAsset!Code
      Me.Combo20 = rstAsset!NameArab
      Me.Combo16 = IIf(IsNull(rstAsset!Modelno) <> True, rstAsset!Modelno, "None")
      Me.Combo17 = IIf(IsNull(rstAsset!SerialNo) <> True, rstAsset!SerialNo, "None")
      Me.Combo9 = rstAsset!AccumulatedCode
      Me.Combo10 = rstAsset!AccumulatedName
      find = True
      If find = True Then Exit Sub
     Else
      If find = False Then
      Me.Combo1 = ""
      Me.Combo14 = ""
      Me.Combo20 = ""
      Me.Combo16 = ""
      Me.Combo17 = ""
      Me.Combo9 = ""
      Me.Combo10 = ""
      End If
     End If
End If
End Sub

Private Sub Combo15_GotFocus()
If Me.Combo15 = "" Then
    Const CB_SHOWDROPDOWN = &H14F
    Dim Tmp
  If Me.Combo15.ListCount > 1 Then
    Tmp = SendMessage(Combo15.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
  End If
End If
End Sub

Private Sub Combo15_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.Combo16.SetFocus
End If
End Sub

Private Sub Combo15_LostFocus()
Call Combo14_Click
End Sub

Private Sub Combo16_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.Combo17.SetFocus
End If
End Sub

Private Sub Combo17_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.Combo18.SetFocus
End If
End Sub

Private Sub Combo18_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.Command7.SetFocus
End If
End Sub

Private Sub Combo19_Click()
LoadSeriesNo = True
Dim rsRegAsset As New ADODB.Recordset
cLen = InStr(1, Me.Combo19, "-")
AssetCat = Trim(Left(Me.Combo19, cLen - 1))
rsRegAsset.Open "Select * from AssetRegistered where category=" & "'" & AssetCat & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
IC = Left(Me.Combo19.Text, 3)
Me.Combo14.Enabled = True
Me.Combo15.Enabled = True
Me.Combo20.Enabled = True
Me.Combo14.clear
Me.Combo15.clear
Me.Combo20.clear
Do Until rsRegAsset.EOF = True
    Me.Combo14.AddItem rsRegAsset!Code
    Me.Combo15.AddItem rsRegAsset!NameEng
    Me.Combo20.AddItem rsRegAsset!NameArab
    rsRegAsset.MoveNext
Loop
rsRegAsset.close
If Left(Trim(Me.SSTab1.caption), 4) = "List" Then
    Me.SSTab1.SetFocus
    SendKeys "{Right}"
 ElseIf Left(Trim(Me.SSTab1.caption), 11) = "Acctg Entry" Then
    Me.SSTab1.SetFocus
    SendKeys "{Right}"
    SendKeys "{Right}"
 ElseIf Left(Trim(Me.SSTab1.caption), 9) = "Item User" Then
    Me.SSTab1.SetFocus
    SendKeys "{Left}"
End If

End Sub

Private Sub Combo19_GotFocus()
Const CB_SHOWDROPDOWN = &H14F
Dim Tmp
Tmp = SendMessage(Combo19.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
'SendKeys "{Down}"
End Sub

Private Sub Combo19_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   
    If Left(Trim(Me.SSTab1.caption), 4) = "List" Then
       Me.SSTab1.SetFocus
       SendKeys "{Right}"
    ElseIf Left(Trim(Me.SSTab1.caption), 11) = "Acctg Entry" Then
       Me.SSTab1.SetFocus
       SendKeys "{Right}"
       SendKeys "{Right}"
    ElseIf Left(Trim(Me.SSTab1.caption), 9) = "Item User" Then
       Me.SSTab1.SetFocus
       SendKeys "{Left}"
    End If
    Me.Combo14.SetFocus
End If
End Sub

Private Sub Combo2_Click()
Dim rstDX As New ADODB.Recordset
rstDX.Open "Select * from FinanceMAster where AccountCode =" & "'" & Trim(Me.Combo2) & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
Me.Combo8 = rstDX!accountnameeng
rstDX.close
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.Combo8.SetFocus
End If
End Sub

Private Sub Combo20_Click()
find = False
rsAssetNAmeArab.MoveFirst
While rsAssetNAmeArab.EOF = False
    If UCase(Trim(Me.Combo20)) = UCase(Trim(rsAssetNAmeArab!NameArab)) Then
      find = True
      Me.Combo1 = rsAssetNAmeArab!Idno
      Me.Combo14 = rsAssetNAmeArab!Code
      Me.Combo15 = rsAssetNAmeArab!NameArab
      Me.Combo16 = IIf(IsNull(rsAssetNAmeArab!Modelno) <> True, rsAssetNAmeArab!Modelno, "None")
      Me.Combo17 = IIf(IsNull(rsAssetNAmeArab!SerialNo) <> True, rsAssetNAmeArab!SerialNo, "None")
      Me.Combo9 = rstAsset!AccumulatedCode
      Me.Combo10 = rstAsset!AccumulatedName
      If find = True Then Exit Sub
     Else
     If find <> True Then
      Me.Combo1 = ""
      Me.Combo14 = ""
      Me.Combo15 = ""
      Me.Combo16 = ""
      Me.Combo17 = ""
      Me.Combo9 = ""
      Me.Combo10 = ""
     End If
    End If
      
   rsAssetNAmeArab.MoveNext
Wend
End Sub

Private Sub Combo20_GotFocus()
GotFocus = True
End Sub

Private Sub Combo20_LostFocus()
GotFocus = False
End Sub

Private Sub Combo3_GotFocus()
Dim havedot As Boolean
  For i = 1 To Len(Trim(Me.Combo3.Text))
        Havedott = Mid(Me.Combo3.Text, i, 1)
        If Havedott = "." Then Exit For
           
  Next
  If Havedott = "." Then
     xdecimal = 1
     Exit Sub
    Else
      xdecimal = 0
  End If

i = 0
For i = 1 To Len(Trim(Me.Combo3.Text))
      X = Mid(Trim(Me.Combo3.Text), i, 1)
      If X = "." Then
        havedot = True
        Exit For
       Else
        havedot = False
      End If
Next i
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.Combo4.SetFocus
End If
If KeyAscii = 46 Then
   xdecimal = xdecimal + 1
 End If
If xdecimal = 2 Then
  Beep
  xdecimal = 1
  SendKeys "{End}"
  SendKeys "{Left}"
  SendKeys "{Delete}"
End If

If KeyAscii = 13 Or KeyAscii = 46 Or KeyAscii >= 48 And KeyAscii <= 57 Then
 c = KeyAscii
 Else
   SendKeys "{Home}"
   SendKeys "{Delete}"
   If KeyAscii = 46 And xdecimal > 1 Then
       xdecimal = 0
   End If
 If Me.Combo3.Text <> " " Then
  xdecimal = 0
  End If
End If

End Sub

Private Sub Combo3_KeyUp(KeyCode As Integer, Shift As Integer)
Dim havedot As Boolean
  For i = 1 To Len(Trim(Me.Combo3.Text))
        Havedott = Mid(Me.Combo3.Text, i, 1)
        If Havedott = "." Then Exit For
           
  Next
  If Havedott = "." Then
     xdecimal = 1
     Exit Sub
    Else
      xdecimal = 0
  End If

i = 0
For i = 1 To Len(Trim(Me.Combo3.Text))
      X = Mid(Trim(Me.Combo3.Text), i, 1)
      If X = "." Then
        havedot = True
        Exit For
       Else
        havedot = False
      End If
Next i
End Sub

Private Sub Combo3_LostFocus()
xdecimal = 0
Me.Combo3.Text = Format(Me.Combo3.Text, "###,###,###,###.#0")

End Sub

Private Sub Combo4_GotFocus()
Dim havedot As Boolean
  For i = 1 To Len(Trim(Me.Combo4.Text))
        Havedott = Mid(Me.Combo4.Text, i, 1)
        If Havedott = "." Then Exit For
           
  Next
  If Havedott = "." Then
     xdecimal = 1
     Exit Sub
    Else
      xdecimal = 0
  End If

i = 0
For i = 1 To Len(Trim(Me.Combo4.Text))
      X = Mid(Trim(Me.Combo4.Text), i, 1)
      If X = "." Then
        havedot = True
        Exit For
       Else
        havedot = False
      End If
Next i

End Sub

Private Sub Combo4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.Combo7.SetFocus
End If
If KeyAscii = 46 Then
   xdecimal = xdecimal + 1
 End If
If xdecimal = 2 Then
  Beep
  xdecimal = 1
  SendKeys "{End}"
  SendKeys "{Left}"
  SendKeys "{Delete}"
End If

If KeyAscii = 13 Or KeyAscii = 46 Or KeyAscii >= 48 And KeyAscii <= 57 Then
 c = KeyAscii
 Else
   'SendKeys "{Home}"
   'SendKeys "{Delete}"
   If KeyAscii = 46 And xdecimal > 1 Then
       xdecimal = 0
   End If
 If Me.Combo4.Text <> " " Then
  xdecimal = 0
  SendKeys "{Home}"
  SendKeys "{Delete}"
 End If
End If

End Sub

Private Sub Combo4_LostFocus()
xdecimal = 0
Me.Combo4.Text = Format(Me.Combo4.Text, "###,###,###,###.#0")

End Sub

Private Sub Combo6_GotFocus()
If Me.Combo6 <> "" Then
    Me.Combo6 = Left(Me.Combo6, Len(Me.Combo6) - 1)
End If
End Sub

Private Sub Combo6_LostFocus()
Me.Combo6 = Format(Me.Combo6, "###.#0") & "%"
End Sub

Private Sub Combo7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.Combo2.SetFocus
End If
If KeyAscii = 13 Or KeyAscii = 46 Or KeyAscii >= 48 And KeyAscii <= 57 Then
 c = KeyAscii
 Else
   SendKeys "{Home}"
   SendKeys "{Delete}"
End If
End Sub

Private Sub Combo8_Click()
Dim rstDX As New ADODB.Recordset
rstDX.Open "Select * from FinanceMAster where AccountNameEng =" & "'" & Trim(Me.Combo8) & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
Me.Combo2 = rstDX!AccountCode
rstDX.close
End Sub

Private Sub Combo8_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.Combo9.SetFocus
End If

End Sub

Private Sub Combo9_Click()
Dim rstDX As New ADODB.Recordset
rstDX.Open "Select * from FinanceMAster where AccountCode =" & "'" & Trim(Me.Combo9) & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
Me.Combo10 = rstDX!accountnameeng
rstDX.close

End Sub

Private Sub Combo9_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.Combo10.SetFocus
End If

End Sub

Private Sub Command1_Click()
If Me.Command3.Enabled = True Then
    Call Command3_Click
   Else
   Unload Me
End If
   

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
mess = MsgBox("This will be now saved your entries, Do want to save now? ", vbOKCancel + vbQuestion, "Please confirm")
If mess = vbOK Then
        
        If Me.Combo3 = "" Or Me.Combo4 = "" Or Me.Combo7 = "" Or Me.Combo2 = "" Or Me.Combo9 = "" Or Me.Combo9 = "" Or Me.Combo10 = "" Then
            mess = MsgBox("Please complete to filled up all the information before saved", vbInformation + vbOKOnly, "Message")
            Exit Sub
        End If
        Dim DepreAmount As Currency
        Dim rstAssetSetup As New ADODB.Recordset
        rstAssetSetup.Open "Select * from AssetSetup where Assetno=" & "'" & Trim(Combo1) & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
        If rstAssetSetup.EOF = False Then
            mess = MsgBox("Items is already on Asset Setup, Choose another.", vbInformation + vbOKOnly, "Message")
            Exit Sub
        End If
        With rstAssetSetup
             .addnew
             !ASsetNo = Me.Combo1
             !AssetCode = Me.Combo14
             !AssetName = Me.Combo15
             !Modelno = Me.Combo16
             !SerialNo = Me.Combo17
             cLen = InStr(1, Me.Combo19, "-")
             !InventoryCat = Trim(Left(Me.Combo19, cLen - 1))
             !AssetType = IIf(Me.Option1.Value = True, "Depreciable", "Amortizable")
             !AcquisitionDate = Me.MaskEdBox1.Text
             !AcquisitionValue = Me.Combo3
             !SalvageValue = Me.Combo4
             If Me.Option3.Value = True Then
                !COmputationMethod = "Straight Line"
               ElseIf Me.Option4.Value = True Then
                !COmputationMethod = "Diminishing Rate"
               ElseIf Me.Option5.Value = True Then
                !COmputationMethod = "Activity Base"
             End If
             If Me.Combo6.Enabled = True Then
                !YearlyFixedRate = Me.Combo6
               Else
                !YearlyFixedRate = " "
             End If
            !AccumulatedDep = 0
            !LastTransdate = " "
            !UsefullLife = Me.Combo7 & " " & Right(Me.Label7, 5)
            !MachineHourUsed = IIf(!UsefullLife = "Hours", Me.Combo7, " ")
            !Unitprod = IIf(Me.Option5.Value = True, Me.Combo5, " ")
            !DEbitAcct = Me.Combo2 & "-" & Me.Combo8
            !CreditAcct = Me.Combo9 & "-" & Me.Combo10
            !remarks = Me.Combo18
            !AssignTo = Me.Combo11 & "/" & Me.Combo12 & "/" & Me.Combo13
            .Update
            .close
        End With
        Call ListIt
        Me.Command3.Enabled = False
        Call ClickTab
        For Each Control In Me
            On Error Resume Next
            If TypeOf Control Is ComboBox Then
                Control.Text = ""
            End If
        Next
        Me.Combo1 = NextAssetCOde
        

End If
    
    

'If Me.Option1.Value = True Then
'    If Me.Option3.Value = True Then
'       DepreAmount = FormatNumber(Me.Combo3, 2, vbTrue, vbTrue, vbTrue) - FormatNumber(Me.Combo4, 2, vbTrue, vbTrue, vbTrue) / FormatNumber(Me.Combo7, 2, vbTrue, vbTrue, vbTrue)
'     ElseIf Me.Option4.Value = True Then
'       DepreAmount = FormatNumber(Me.Combo3, 2, vbTrue, vbTrue, vbTrue) - FormatNumber(Me.Combo4, 2, vbTrue, vbTrue, vbTrue) - AccumulatedDep * (FormatNumber(Me.Combo6, 2, vbTrue, vbTrue, vbTrue) / 100)
'     ElseIf Me.Option5.Value = True Then
'       If Me.Option6.Value = True Then
'          DepreAmount = FormatNumber(Me.Combo3, 2, vbTrue, vbTrue, vbTrue) - FormatNumber(Me.Combo4, 2, vbTrue, vbTrue, vbTrue) / Val(Me.Combo18)
'        Else
'          DepreAmount = FormatNumber(Me.Combo3, 2, vbTrue, vbTrue, vbTrue) - FormatNumber(Me.Combo4, 2, vbTrue, vbTrue, vbTrue) / Val(Me.Combo5)
'       End If
'     End If
'  Else
'    DepreAmount = FormatNumber(Me.Combo3, 2, vbTrue, vbTrue, vbTrue) / Val(Me.Combo7)
' End If
End Sub

Private Sub Command5_Click()
Dim rstEditRec As New ADODB.Recordset
        'to locate also the item in listview
        i = 0
        For i = 1 To Me.ListView1.ListItems.Count
          Me.ListView1.ListItems.Item(i).ForeColor = vbBlack
          Me.ListView1.ListItems.Item(i).Ghosted = False
        Next
        Dim strFindMe As String
        Dim itmFound As ListItem   ' FoundItem variable.
        intSelectedOption = lvwText
        strFindMe = Trim(Me.Combo1)
        Set itmFound = Me.ListView1.Finditem(strFindMe, intSelectedOption, , lvwPartial)
        If itmFound Is Nothing Then  ' If no match, inform user and exit.
         mess = MsgBox("Asset Number not found", vbOKOnly + vbInformation, "Message")
         Else
           itmFound.EnsureVisible
           itmFound.Selected = True   ' Select the ListItem.
           Me.ListView1.SetFocus
           Me.ListView1.SelectedItem.ForeColor = vbBlue
           Me.ListView1.SelectedItem.Ghosted = vbBlue
        End If
CurrentJn = Trim(Me.Combo1)
cItem = Trim(Me.Combo1)
rstEditRec.Open "Select * from ASsetSetup where AssetNo= " & "'" & cItem & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
Me.Combo14.clear
Me.Combo15.clear
If rstEditRec.EOF = False Then
    Me.Combo11 = rstEditRec!AssignTo
    Me.Combo14 = rstEditRec!AssetCode
    Me.Combo15 = rstEditRec!AssetName
    Me.Combo16 = rstEditRec!Modelno
    Me.Combo17 = rstEditRec!SerialNo
    Me.Combo18 = rstEditRec!remarks
    If rstEditRec!AssetType = "Depreciable" Then
        Me.Option1.Value = True
       Else
        Me.Option2.Value = True
    End If
    If Trim(rstEditRec!COmputationMethod) = "Straight Line" Then
        Me.Option3.Value = True
     ElseIf (rstEditRec!COmputationMethod) = "Diminishing Rate" Then
        Me.Option4.Value = True
        Me.Combo6 = IIf(IsNull(rstEditRec!YearlyFixedRate) = False, rstEditRec!YearlyFixedRate, " ")
     ElseIf (rstEditRec!COmputationMethod) = "Activity Base" Then
        Me.Option5.Value = True
    End If
    Me.MaskEdBox1.Text = Format(rstEditRec!AcquisitionDate, "dd/mm/yyyy")
    Me.Combo3 = FormatNumber(rstEditRec!AcquisitionValue, 2, vbTrue, vbTrue, vbTrue)
    Me.Combo4 = FormatNumber(rstEditRec!SalvageValue, 2, vbTrue, vbTrue, vbTrue)
    Me.Combo7 = rstEditRec!UsefullLife
    Me.Combo14.Enabled = True
    Me.Combo15.Enabled = True
    rstEditRec.close
    Me.Combo14.SetFocus
    
  Else
   For Each Control In Me
       On Error Resume Next
       If TypeOf Control Is ComboBox Then
           Control.Text = ""
       End If
   Next
  Dim rstWhatJnNow As New ADODB.Recordset
  rstWhatJnNow.Open "Setup", constring, adOpenKeyset, adLockPessimistic, adCmdTable
  If rstWhatJnNow!NextAssetCOde = Val(CurrentJn) Then
     Me.Combo19.SetFocus
    Else
    Me.Combo1.SetFocus
  End If
End If
Me.Combo1 = CurrentJn
'If Trim(Me.SSTab1.Caption) = "List of Assets Setup" Then
'        Me.SSTab1.SetFocus
'        SendKeys "{Right}"
'     ElseIf Trim(Me.SSTab1.Caption) = "Item User" Then
'        Me.SSTab1.SetFocus
'        SendKeys "{Left}"
'        Me.Timer1.Enabled = False
'        Me.Command7.Enabled = False
'    ElseIf Trim(Me.SSTab1.Caption) = "Accounting Entry" Then
'        Me.SSTab1.SetFocus
'        SendKeys "{Right}"
'        SendKeys "{Right}"
'        Me.Timer1.Enabled = False
'        Me.Command7.Enabled = False
'    End If
End Sub

Private Sub Command7_Click()
'    Dim ACExist As Boolean
'    Dim rstAsset As New ADODB.Recordset
'    rstAsset.Open "SElect * from AssetSetup Where AssetNo=" & "'" & Trim(Me.Combo1) & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
'    If rstAsset.EOF = False Then
'        ACExist = True
'    End If
'    rstAsset.Close
     On Error Resume Next
     rstAsset.close
     On Error GoTo 0
    If Left(Trim(Me.SSTab1.caption), 10) = "Asset Info" Then
        rstAsset.Open "SElect * from AssetSetup Where Assetno=" & "'" & Trim(Me.Combo1) & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
        If rstAsset.EOF = False Then
            mess = MsgBox("Items is already on Asset Setup, Choose another.", vbInformation + vbOKOnly, "Message")
            Me.Combo14.SetFocus
            Exit Sub
        End If
            
        Me.SSTab1.SetFocus
        SendKeys "{Right}"
        Me.Command7.Enabled = True
     ElseIf Left(Trim(Me.SSTab1.caption), 9) = "Item User" Then
        Me.SSTab1.SetFocus
        SendKeys "{Right}"
        Me.Timer1.Enabled = False
        Me.Command7.Enabled = False
    End If
End Sub

Private Sub Command8_Click()
'PrintSalesJOurnal
End Sub

Private Sub Form_Activate()
On Error Resume Next
Me.Combo1.SetFocus

End Sub

Private Sub Form_Load()
LoadSeriesNo = True
With Me.Combo5
    .AddItem "Pcs"
    .AddItem "Metric"
    .AddItem "Linear"
    .AddItem "Tons"
End With
Dim rsAssetCat As New ADODB.Recordset

CodeBelongASset = "11201" 'query in level4
CodeBelongAccumulatedDep = "12201" 'query it at level5 table
rsAssetCat.Open "select * from Level4 where Left(AccountCode,5)=" & "'" & CodeBelongASset & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
Do Until rsAssetCat.EOF = True
   If Trim(rsAssetCat!MainAcct) <> "Main Accts" And UCase(Trim(rsAssetCat!accountnameeng)) <> UCase("Lands") Then
    Me.Combo19.AddItem Trim(rsAssetCat!accountnameeng) & " - " & Trim(rsAssetCat!AccountCode)
   End If
   rsAssetCat.MoveNext
Loop
Me.Combo1 = nextIdno
rsAssetCat.close

 Dim xClass As New HabitatClass
 Dim xtable As String
 Dim sqltable As Boolean
 xtable = "Select * from AssetREgistered order by IDNo"
 sqltable = True
 xClass.GetTables rsAssetNAmeArab, CON1, xtable, constring, sqltable

Call Option1_Click
Call ListIt
End Sub

Private Sub Form_Resize()
On Error Resume Next
Me.SSTab1.Width = Me.Width - 300
Me.SSTab1.Height = Me.Height - 1600
Me.Frame8.Height = Me.SSTab1.Height - 800
Me.Frame8.Width = Me.SSTab1.Width - 500
'Me.Frame5.Width = Me.SSTab1.Width - 6850
'Me.Frame7.Width = Me.SSTab1.Width - 6850
Me.ListView1.Width = Me.SSTab1.Width - 120
Me.ListView1.Height = Me.SSTab1.Height - 470
Me.Command1.Top = Me.SSTab1.Height + 700
Me.Command2.Top = Me.SSTab1.Height + 700
Me.Command3.Top = Me.SSTab1.Height + 700
Me.Command7.Top = Me.SSTab1.Height + 700
Me.Command3.Left = Me.Width - 1650
Me.Command2.Left = Me.Width - 2980
Me.Command1.Left = Me.Width - 4330
Me.Command7.Left = Me.Width - 6000
Me.Picture1.Top = Me.Command7.Top
Me.Label21.Top = Me.Command7.Top
'Me.Combo19.Width = Me.Width - 7050
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
rstEditRec.close
'rstAsset.Close
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Me.ListView1.SortKey = ColumnHeader.Index - 1
Me.ListView1.Sorted = True
End Sub

Private Sub ListView1_DblClick()
Call xedit_Click
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
LoadSeriesNo = False
If UCase(Trim(Me.ListView1.SelectedItem.SubItems(9))) = UCase("Activity Base") Then
    Me.xMachineUsed.Enabled = True
 Else
    Me.xMachineUsed.Enabled = fasle
End If
Me.Combo1 = Me.ListView1.SelectedItem
'Dim rstEditRec As New ADODB.Recordset
cItem = Trim(Me.ListView1.SelectedItem.Text)
cindex = Me.ListView1.SelectedItem.Index
On Error Resume Next
rstEditRec.close
On Error GoTo 0
rstEditRec.Open "Select * from ASsetSetup where AssetNo= " & "'" & cItem & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
rstEditRec.MoveFirst
 'If rstEditRec.EOF = False Then
 While rstEditRec.EOF = False
  If Trim(rstEditRec!ASsetNo) = cItem Then
    Me.Combo14.clear
    Me.Combo15.clear
    Me.Combo1 = rstEditRec!ASsetNo
    Me.Combo11 = rstEditRec!AssignTo
    Me.Combo14 = rstEditRec!AssetCode
    Me.Combo15 = rstEditRec!AssetName
    Me.Combo16 = rstEditRec!Modelno
    Me.Combo17 = rstEditRec!SerialNo
    Me.Combo18 = rstEditRec!remarks
    If rstEditRec!AssetType = "Depreciable" Then
        Me.Option1.Value = True
       Else
        Me.Option2.Value = True
    End If
    If Trim(rstEditRec!COmputationMethod) = "Straight Line" Then
        Me.Option3.Value = True
     ElseIf (rstEditRec!COmputationMethod) = "Diminishing Rate" Then
        Me.Option4.Value = True
        Me.Combo6 = IIf(IsNull(rstEditRec!YearlyFixedRate) = False, rstEditRec!YearlyFixedRate, " ")
     ElseIf (rstEditRec!COmputationMethod) = "Activity Base" Then
        Me.Option5.Value = True
    End If
    Me.MaskEdBox1.Text = Format(rstEditRec!AcquisitionDate, "dd/mm/yyyy")
    Me.Combo3 = FormatNumber(rstEditRec!AcquisitionValue, 2, vbTrue, vbTrue, vbTrue)
    Me.Combo4 = FormatNumber(rstEditRec!SalvageValue, 2, vbTrue, vbTrue, vbTrue)
    Me.Combo7 = Val(rstEditRec!UsefullLife)
    Me.Combo2 = Left(rstEditRec!DEbitAcct, 12)
    Me.Combo8 = Mid(rstEditRec!DEbitAcct, 14, 50)
    Me.Combo9 = Left(rstEditRec!CreditAcct, 12)
    Me.Combo10 = Mid(rstEditRec!CreditAcct, 14, 50)
    Me.Combo14.Enabled = True
    Me.Combo15.Enabled = True
   'rstEditRec.Close
    If rstEditRec.EOF = False Then
        rstEditRec.MoveLast
    End If
 End If
 rstEditRec.MoveNext
Wend
Me.ListView1.SetFocus
End Sub

Private Sub ListView1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
    Call SEtupNew_Click
End If
If KeyCode = 114 Then
    Call xedit_Click
End If

If KeyCode = 46 Then
    Call xdelete_Click
End If

End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = 2 Then
    PopupMenu main
End If
End Sub

Private Sub MaskEdBox1_GotFocus()
Me.MaskEdBox1.Text = Format(Me.MaskEdBox1.Text, "dd/mm/yyyy")
End Sub

Private Sub MaskEdBox1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.Combo3.SetFocus
End If
End Sub

Private Sub MaskEdBox1_LostFocus()
If Me.MaskEdBox1.Text = "__/__/____" Then
    Exit Sub
End If
Me.MaskEdBox1.Text = Format(Me.MaskEdBox1.Text, "dd/mm/yyyy")
cDay = Val(Left(Me.MaskEdBox1.Text, 2))
cMonth = Val(Mid(Me.MaskEdBox1.Text, 4, 2))
cYear = Val(Right(Me.MaskEdBox1.Text, 4))
If cDay > 31 Or cDay < 1 Then
    mess = MsgBox("Invalid Date", vbInformation + vbOKOnly, "Message")
    Me.MaskEdBox1.SetFocus
  ElseIf cMonth > 12 Or cMonth < 1 Then
    mess = MsgBox("Invalid Month", vbInformation + vbOKOnly, "Message")
    Me.MaskEdBox1.SetFocus
ElseIf cYear < 1900 Or cYear > Year(Date) Then
    mess = MsgBox("Invalid Year", vbInformation + vbOKOnly, "Message")
    Me.MaskEdBox1.SetFocus
End If
End Sub

Private Sub Option1_Click()
Me.Option3.Enabled = True
Me.Option4.Enabled = True
Me.Option5.Enabled = True

Dim xtable As String
Dim sqltable As Boolean
Dim xClass As New HabitatClass
Set CON1 = New ADODB.Connection
Set rstDEpAsset = New ADODB.Recordset
'Set acctnames = New ADODB.Recordset
xtable = "Select * from FinanceMaster order by Accountcode"
sqltable = True
xClass.GetTables rstDEpAsset, CON1, xtable, constring, sqltable

Dim DepreciatedAsset As New ADODB.Recordset
Dim AmortizableAsset As New ADODB.Recordset

Dim rstAmortAsset As New ADODB.Recordset
DepreciatedAsset.Open "DepreciatedAssetsSeries", constring, adOpenKeyset, adLockPessimistic, adCmdTable

''rstDEpAsset.Open "FinanceMAster", conString, adOpenKeyset, adLockPessimistic, adCmdTable
'
'i = 0
'xValComb9 = Me.Combo9
'xValCOmb10 = Me.Combo10
'Me.Combo9.clear
'Me.Combo10.clear
'If LoadSeriesNo = False Then
'    Exit Sub
'End If
'Do Until DepreciatedAsset.EOF = True
'    i = i + 1
'    Series = "Series" & Trim(i)
'    Series = DepreciatedAsset!Series1
'    cLen = Len(Series)
'
'    While rstDEpAsset.EOF = False
'      If rstDEpAsset!Active = 1 Then
'       If Left(rstDEpAsset!AccountCode, cLen) = Series Then
'        Me.Combo9.AddItem rstDEpAsset!AccountCode
'        Me.Combo10.AddItem rstDEpAsset!accountnameeng
'       End If
'      End If
'      rstDEpAsset.MoveNext
'   Wend
'   ' rstDEpAsset.Close
'
'    DepreciatedAsset.MoveNext
'Loop
'rstDEpAsset.MoveFirst
'DepreciatedAsset.Close
'Me.Combo9 = xValComb9
'Me.Combo10 = xValCOmb10
'
''open the All Depreciation Expense Account
'DepreciatedAsset.Open "DepreciationExpensesSeries", constring, adOpenKeyset, adLockPessimistic, adCmdTable
'i = 0
'Me.Combo2.clear
'Me.Combo8.clear
'Do Until DepreciatedAsset.EOF = True
'    i = i + 1
'    Series = "Series" & Trim(i)
'    Series = DepreciatedAsset!Series
'    cLen = Len(Series)
'    'rstDEpAsset.Open "FinanceMAster", conString, adOpenKeyset, adLockPessimistic, adCmdTable
'    While rstDEpAsset.EOF = False
'      If rstDEpAsset!Active = 1 Then
'       If Left(rstDEpAsset!AccountCode, cLen) = Series Then
'        Me.Combo2.AddItem rstDEpAsset!AccountCode
'        Me.Combo8.AddItem rstDEpAsset!accountnameeng
'       End If
'      End If
'       rstDEpAsset.MoveNext
'   Wend
'    'rstDEpAsset.Close
'
'    DepreciatedAsset.MoveNext
'Loop
'DepreciatedAsset.Close


End Sub

Private Sub Option2_Click()
If LoadSeriesNo = False Then
    Exit Sub
End If


Me.Command6.Enabled = False
Me.Label7.caption = "Usefull Life in Yrs"
Me.Option3.Enabled = False
Me.Option4.Enabled = False
Me.Option5.Enabled = False

Me.Option6.Enabled = False
Me.Option7.Enabled = False
Me.Label5.Enabled = False
Me.Combo5.Enabled = False
Me.Combo6.Enabled = False
Me.MaskEdBox1.SetFocus
'Dim DepreciatedAsset As New ADODB.Recordset
'Dim Amort
'Dim DepreizableAsset As New ADODB.Recordset
'Dim rstDEpAsset As New ADODB.Recordset
'Dim rstAmortAsset As New ADODB.Recordset
'DepreciatedAsset.Open "AmortizableAssetsSeries", constring, adOpenKeyset, adLockPessimistic, adCmdTable
'i = 0
'xValComb9 = Me.Combo9
'xValCOmb10 = Me.Combo10
'Me.Combo9.clear
'Me.Combo10.clear
'Do Until DepreciatedAsset.EOF = True
'    i = i + 1
'    Series = "Series" & Trim(i)
'    Series = DepreciatedAsset!Series1
'    cLen = Len(Series)
'    rstDEpAsset.Open "FinanceMAster", constring, adOpenKeyset, adLockPessimistic, adCmdTable
'    While rstDEpAsset.EOF = False
'       If Left(rstDEpAsset!AccountCode, cLen) = Series Then
'        Me.Combo9.AddItem rstDEpAsset!AccountCode
'        Me.Combo10.AddItem rstDEpAsset!accountnameeng
'       End If
'       rstDEpAsset.MoveNext
'   Wend
'    rstDEpAsset.Close
'
'    DepreciatedAsset.MoveNext
'Loop
'DepreciatedAsset.Close
'Me.Combo9 = xValComb9
'Me.Combo10 = xValCOmb10


End Sub

Private Sub Option3_Click()
Me.Command6.Enabled = False
Me.Label7.caption = "Usefull Life in Yrs"
Me.Frame4.Enabled = False
Me.Option6.Enabled = False
Me.Option7.Enabled = False
Me.Frame3.Enabled = True
Me.Combo5.Enabled = False
'Me.Combo18.Enabled = False
Me.Label5.Enabled = False
End Sub

Private Sub Option4_Click()
Me.Label7.caption = "Usefull Life in Years"
Me.Option6.Enabled = False
Me.Option7.Enabled = False

Me.Combo6.Enabled = True
Me.Frame4.Enabled = True

Me.Combo5.Enabled = False
Me.Label5.Enabled = False

Me.Command6.Enabled = False
If LoadSeriesNo = True Then
 Me.Combo6.SetFocus
End If
End Sub

Private Sub Option5_Click()
Me.Frame5.Visible = True
Me.Label7.caption = "Usefull Life in Hours"
Me.Option3.Enabled = True
Me.Option5.Enabled = True
Me.Frame5.Enabled = True
Me.Option6.Enabled = True
Me.Option7.Enabled = True
Me.Label5.Enabled = False
Me.Frame4.Enabled = False
Me.Combo6.Enabled = False
Me.Command6.Enabled = True
End Sub

Private Sub Option6_Click()
Me.Label5.caption = "&Unit"
Me.Label5.Enabled = True
Me.Combo5.Enabled = True
Me.Combo5.Visible = True
'Me.Combo18.Visible = False
End Sub

Private Sub Option7_Click()
'Me.Label5.Caption = "Ho&urs"
Me.Label5.Enabled = False
'Me.Combo18.Enabled = True
Me.Combo5.Enabled = False
'Me.Combo18.Visible = True
End Sub

Private Sub SEtupNew_Click()
LoadSeriesNo = True
For Each Control In Me
            On Error Resume Next
            If TypeOf Control Is ComboBox Then
                Control.Text = ""
            End If
 Next
rstAsset.Open "Setup", constring, adOpenKeyset, adLockPessimistic, adCmdTable
Me.Combo1 = rstAsset!NextAssetCOde
rstAsset.close
rstAsset.Open "InventoryCategory", constring, adOpenKeyset, adLockPessimistic, adCmdTable
Do Until rstAsset.EOF = True
    Me.Combo19.AddItem rstAsset!iccode & " - " & rstAsset!icnameeng
    rstAsset.MoveNext
Loop
rstAsset.close
Me.SSTab1.SetFocus
SendKeys "{Right}"
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
X = Me.SSTab1.caption
xPrevTAb = PreviousTab
If Trim(Me.SSTab1.caption) = "List of Assets Setup" Or Trim(Me.SSTab1.caption) = "Accounting Entry" Then
    Me.Timer1.Enabled = False
    Me.Command7.Enabled = False
  Else
  Me.Timer1.Enabled = True
   Me.Command7.Enabled = True
End If
End Sub

Private Sub Timer1_Timer()
If Me.Combo14 = "" Then
    Me.Command7.Enabled = False
    Me.Command3.Enabled = False
 Else
 Me.Command7.Enabled = True
 Me.Command3.Enabled = True
End If
End Sub

Private Sub Timer2_Timer()
If Me.Combo14 = "" Then
     Me.Command3.Enabled = False
 Else
 Me.Command3.Enabled = True
End If

End Sub

Private Sub xdelete_Click()
Dim rstDelRec As New ADODB.Recordset
cItem = Trim(Me.ListView1.SelectedItem.Text)
cindex = Me.ListView1.SelectedItem.Index
mess = MsgBox("Delete " & Me.ListView1.SelectedItem.SubItems(2) & "?", vbQuestion + vbYesNo + vbDefaultButton2, "Please confirm")
If mess = vbYes Then
    rstDelRec.Open "Delete AssetSetup where AssetNo=" & "'" & cItem & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
    Me.ListView1.ListItems.Remove cindex
End If
End Sub

Private Sub xedit_Click()
Dim rstEditRec As New ADODB.Recordset
cItem = Trim(Me.ListView1.SelectedItem.Text)
cindex = Me.ListView1.SelectedItem.Index
rstEditRec.Open "Select * from ASsetSetup where AssetNo= " & "'" & cItem & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
If rstEditRec.EOF = False Then
    Me.Combo14.clear
    Me.Combo15.clear
    Me.Combo1 = rstEditRec!ASsetNo
    Me.Combo11 = rstEditRec!AssignTo
    Me.Combo14 = rstEditRec!AssetCode
    Me.Combo15 = rstEditRec!AssetName
    Me.Combo16 = rstEditRec!Modelno
    Me.Combo17 = rstEditRec!SerialNo
    Me.Combo18 = rstEditRec!remarks
    If rstEditRec!AssetType = "Depreciable" Then
        Me.Option1.Value = True
       Else
        Me.Option2.Value = True
    End If
    If Trim(rstEditRec!COmputationMethod) = "Straight Line" Then
        Me.Option3.Value = True
     ElseIf (rstEditRec!COmputationMethod) = "Diminishing Rate" Then
        Me.Option4.Value = True
        Me.Combo6 = IIf(IsNull(rstEditRec!YearlyFixedRate) = False, rstEditRec!YearlyFixedRate, " ")
     ElseIf (rstEditRec!COmputationMethod) = "Activity Base" Then
        Me.Option5.Value = True
    End If
    Me.MaskEdBox1.Text = Format(rstEditRec!AcquisitionDate, "dd/mm/yyyy")
    Me.Combo3 = FormatNumber(rstEditRec!AcquisitionValue, 2, vbTrue, vbTrue, vbTrue)
    Me.Combo4 = FormatNumber(rstEditRec!SalvageValue, 2, vbTrue, vbTrue, vbTrue)
    Me.Combo7 = Val(rstEditRec!UsefullLife)
    Me.Combo2 = Left(rstEditRec!DEbitAcct, 12)
    Me.Combo8 = Mid(rstEditRec!DEbitAcct, 14, 50)
    Me.Combo9 = Left(rstEditRec!CreditAcct, 12)
    Me.Combo10 = Mid(rstEditRec!CreditAcct, 14, 50)
    Me.Combo14.Enabled = True
    Me.Combo15.Enabled = True
 rstEditRec.close
 Me.SSTab1.SetFocus
 SendKeys "{Right}"
End If
End Sub

Private Sub xFind_Click()
xSeek = True
Me.Combo1.SetFocus

End Sub

Private Sub xGenerate_Click()
Me.ListView1.SetFocus
Dim AssetSetup As New ADODB.Recordset
Dim ASsetJourn As New ADODB.Recordset
Dim JOurnalNo As New ADODB.Recordset
Dim FinanceMAster As New ADODB.Recordset
TRansDate = Format(Date, "dd/mm/yyyy")
ASsetJourn.Open "Delete AssetJournal where Transdate=" & "'" & TRansDate & "'" & "and remarks is null", constring, adOpenKeyset, adLockPessimistic, adCmdText

AssetSetup.Open "Select * from ASsetSetup order by AssetNo", constring, adOpenKeyset, adLockPessimistic, adCmdText
ASsetJourn.Open "ASsetJOurnal", constring, adOpenKeyset, adLockPessimistic, adCmdTable

JOurnalNo.Open "Setup", constring, adOpenKeyset, adLockPessimistic, adCmdTable
If Val(Left(JOurnalNo!CurrentMoYr, 2)) <> Format(Date, "mm") Then
   JOurnalNo!CurrentMoYr = Format(Date, "mm") & Trim(Right(Year(Date), 2))
   JOurnalNo!nextjn = "00001"
   JOurnalNo.Update
Else
   Jn = "AST" & "-" & Right(Year(Date), 2) & Trim(Format(Date, "mm")) & "-" & JOurnalNo!nextjn
   nextjn = Val(JOurnalNo!nextjn)
   If Len(nextjn) = 1 Then
    Zeros = "0000"
    ElseIf Len(nextjn) = 2 Then
    Zeros = "000"
    ElseIf Len(nextjn) = 3 Then
    Zeros = "00"
    ElseIf Len(nextjn) = 4 Then
    Zeros = "0"
    ElseIf Len(nextjn) = 6 Then
    Zeros = ""
   End If
   JOurnalNo!nextjn = Zeros & Trim(Val(nextjn) + 1)
   JOurnalNo.Update
   JOurnalNo.close
End If

Totrec = Int(AssetSetup.RecordCount / 100)
If Totrec = 0 Then
    Totrec = 1
End If
l = 0
cVal = 0
Me.Label21.Visible = True
Me.Picture1.Visible = True
i = 0
TN = 0
Dim strFindMe As String
Dim itmFound As ListItem   ' FoundItem variable.
Do Until AssetSetup.EOF = True

        'Lets navigate the listview items
        intSelectedOption = lvwText
        strFindMe = Trim(AssetSetup!ASsetNo)
        Set itmFound = Me.ListView1.Finditem(strFindMe, intSelectedOption, , lvwPartial)
        If itmFound Is Nothing Then  ' If no match, inform user and exit.
           
         Else
           itmFound.EnsureVisible
           itmFound.Selected = True   ' Select the ListItem.
           Me.ListView1.SetFocus
        End If
        
      TN = TN + 1
      i = i + 1
      l = l + 1
      If l = Totrec Then
        cVal = cVal + 1
        On Error Resume Next
        Me.ProgressBar1.Value = cVal
        l = 0
        On Error GoTo 0
       End If
       
       If UCase(AssetSetup!AssetType) = UCase("Depreciable") Then
              If UCase(AssetSetup!COmputationMethod) = UCase("Straight Line") Then
                 Description = "For the period of " & Format(Date, "dd/mm/yyyy")
                 Depamt = ((AssetSetup!AcquisitionValue - AssetSetup!SalvageValue) / Val(AssetSetup!UsefullLife))
               ElseIf UCase(AssetSetup!COmputationMethod) = UCase("Diminishing Rate") Then
                  Depamt = ((AssetSetup!AcquisitionValue - AssetSetup!SalvageValue) - Val(AssetSetup!AccumulatedDep) * Val(AssetSetup!YearlyFixedRate))
               ElseIf UCase(AssetSetup!COmputationMethod) = UCase("Activity Base") Then
                  If AssetSetup!Unitprod = "" Then
                    Depamt = ((AssetSetup!AcquisitionValue - AssetSetup!SalvageValue) / MachineUsed)
                   Else
                    Depamt = ((AssetSetup!AcquisitionValue - AssetSetup!SalvageValue) / Val(AssetSetup!UsefullLife))
                  End If
               End If
           Else
               Description = "" 'Amotization Amount for the of " & Format(Date, "mmm. yyyy")
               Depamt = (AssetSetup!AcquisitionValue / Val(AssetSetup!UsefullLife))
        End If
       
       
       'this will be appended to ASsetjournal Table
      If AssetSetup!AcquisitionValue - (AssetSetup!AccumulatedDep + Depamt) >= AssetSetup!SalvageValue Then
       With ASsetJourn
            'for Debit side
            .addnew
            !ticket = TN
            TN = TN + 1
            !SerialNo = Jn
            !accountnumber = Left(AssetSetup!DEbitAcct, 12)
            !accountname = Trim(Mid(AssetSetup!DEbitAcct, 14, 30))
            !TRansDate = Format(Date, "dd/mm/yyyy")
            !BeginNingBal = DrBeginBal
            !Description = Description
             CreditDesc = !Description
             !DebitAmount = Depamt
             !creditamount = 0
             !EndingBal = DrBeginBal + Depamt
             .Update
              
             'for Credit Side
             .addnew
            !ticket = TN
            TN = TN
            !SerialNo = Jn
            !accountnumber = Left(AssetSetup!CreditAcct, 12)
            !accountname = Trim(Mid(AssetSetup!CreditAcct, 14, 30))
            !TRansDate = Format(Date, "dd/mm/yyyy")
            !BeginNingBal = 0
            !Description = CreditDesc
            !DebitAmount = 0
            !creditamount = Depamt
            !EndingBal = 0
            .Update
             
           End With
        End If
DoEvents
AssetSetup.MoveNext
Loop
If cVal < 100 And AssetSetup.EOF = True Then
            Me.ProgressBar1.Value = 100
End If
msg = MsgBox("Kalas Sadik,Click me then I will show you the list ", vbInformation + vbOKOnly, "Message")
Me.Picture1.Visible = False
Me.Label21.Visible = False
'Unload Me
AssetJOunralList.Show 1



End Sub

Private Sub xMachineUsed_Click()
MachineHRs.Show 1
End Sub

VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form BankTransaction1 
   Caption         =   "Bank Transaction"
   ClientHeight    =   8490
   ClientLeft      =   75
   ClientTop       =   -76545
   ClientWidth     =   11625
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "BankTransaction1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11625
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   8295
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   14631
      _Version        =   393216
      Tab             =   1
      TabHeight       =   617
      TabMaxWidth     =   3528
      MouseIcon       =   "BankTransaction1.frx":014A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Transaction Entry"
      TabPicture(0)   =   "BankTransaction1.frx":0166
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label24"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label25"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label26"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label27"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label29"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label22"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label30"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label31"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label37"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label38"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label39"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Text9"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Text7"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Text8"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "ListView3"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "ImageList1"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Combo16"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Combo17"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Combo18"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Combo19"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Combo20"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Check1"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "ImageList3"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Timer2"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "CoolBar2"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Combo15"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "MaskEdBox3"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Check2"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Command1"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Combo21"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "List1"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).ControlCount=   31
      TabCaption(1)   =   "Transaction Analysis"
      TabPicture(1)   =   "BankTransaction1.frx":0182
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Combo3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Text2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Text1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Text3"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame5"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Frame2"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "ImageList2"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "CoolBar1"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Frame3"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Frame4"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Frame1"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Text5"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Text4"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Text6"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Timer1"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Frame6"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Frame8"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).ControlCount=   17
      TabCaption(2)   =   "Journal List ÞÇÆãÉ ÇáÚÇã"
      TabPicture(2)   =   "BankTransaction1.frx":019E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "ListView2"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.ListBox List1 
         Height          =   300
         Left            =   -63360
         TabIndex        =   96
         Top             =   720
         Width           =   150
      End
      Begin VB.Frame Frame8 
         Caption         =   "FXCY Amount"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   6720
         TabIndex        =   90
         Top             =   4200
         Width           =   2055
         Begin VB.ComboBox Combo23 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            Style           =   1  'Simple Combo
            TabIndex        =   91
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.ComboBox Combo21 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -65640
         Style           =   1  'Simple Combo
         TabIndex        =   89
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Show Cash Deposit List... ÚÑÖ ÞÇÆãÉ ÈÇáÇíÏÇÚÇÊ ÇáäÞÏíÉ"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   -66120
         TabIndex        =   87
         Top             =   1080
         Width           =   2295
      End
      Begin VB.CheckBox Check2 
         Caption         =   "&Direct Deposit ÇíÏÇÚ ãÈÇÔÑ "
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -66120
         TabIndex        =   86
         Top             =   480
         Width           =   2295
      End
      Begin MSMask.MaskEdBox MaskEdBox3 
         Height          =   375
         Left            =   -68400
         TabIndex        =   85
         Top             =   2280
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox Combo15 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -68400
         Style           =   1  'Simple Combo
         TabIndex        =   83
         Top             =   1920
         Width           =   1575
      End
      Begin ComCtl3.CoolBar CoolBar2 
         Height          =   420
         Left            =   -74910
         TabIndex        =   11
         Top             =   2880
         Width           =   11550
         _ExtentX        =   20373
         _ExtentY        =   741
         _CBWidth        =   11550
         _CBHeight       =   420
         _Version        =   "6.7.8988"
         MinHeight1      =   360
         Width1          =   2880
         NewRow1         =   0   'False
         MinHeight2      =   360
         Width2          =   3495
         NewRow2         =   0   'False
         MinHeight3      =   360
         Width3          =   2505
         NewRow3         =   0   'False
         Begin MSComctlLib.Toolbar Toolbar3 
            Height          =   330
            Left            =   9650
            TabIndex        =   77
            Top             =   45
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   582
            ButtonWidth     =   609
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            Style           =   1
            ImageList       =   "ImageList3"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   2
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "find"
                  Object.ToolTipText     =   "Go"
                  ImageIndex      =   3
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
            EndProperty
         End
         Begin VB.PictureBox Picture2 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   6600
            ScaleHeight     =   255
            ScaleWidth      =   495
            TabIndex        =   76
            Top             =   120
            Width           =   495
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "&Find"
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
               TabIndex        =   15
               Top             =   0
               Width           =   495
            End
         End
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   3120
            ScaleHeight     =   255
            ScaleWidth      =   735
            TabIndex        =   75
            Top             =   120
            Width           =   735
            Begin VB.Label Label8 
               BackStyle       =   0  'Transparent
               Caption         =   "Fil&ter by"
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
               Left            =   0
               TabIndex        =   13
               Top             =   0
               Width           =   735
            End
         End
         Begin VB.ComboBox Combo10 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   7200
            Style           =   1  'Simple Combo
            TabIndex        =   16
            Top             =   45
            Width           =   2055
         End
         Begin VB.ComboBox Combo9 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3960
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   50
            Width           =   2415
         End
         Begin MSComctlLib.Toolbar Toolbar2 
            Height          =   330
            Left            =   120
            TabIndex        =   12
            Top             =   45
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   582
            ButtonWidth     =   609
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            Style           =   1
            ImageList       =   "ImageList1"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   6
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Save"
                  Object.ToolTipText     =   "Save to file"
                  ImageIndex      =   15
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Print"
                  Object.ToolTipText     =   "Print"
                  ImageIndex      =   3
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Delete"
                  Object.ToolTipText     =   "Delete"
                  ImageIndex      =   6
               EndProperty
               BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Return"
                  Object.ToolTipText     =   "Return and close this window"
                  ImageIndex      =   11
               EndProperty
               BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
            EndProperty
         End
      End
      Begin VB.Timer Timer2 
         Interval        =   1000
         Left            =   -74640
         Top             =   4080
      End
      Begin MSComctlLib.ImageList ImageList3 
         Left            =   -63960
         Top             =   2520
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "BankTransaction1.frx":01BA
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "BankTransaction1.frx":04D4
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "BankTransaction1.frx":07EE
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "BankTransaction1.frx":0C40
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Set to Arabic  áÊÍæíá Çáí ÇáÚÑÈí"
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
         Left            =   -66720
         TabIndex        =   18
         Top             =   50
         Width           =   3255
      End
      Begin VB.ComboBox Combo20 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -72720
         TabIndex        =   2
         Top             =   600
         Width           =   2775
      End
      Begin VB.ComboBox Combo19 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -72720
         TabIndex        =   4
         Top             =   960
         Width           =   4335
      End
      Begin VB.ComboBox Combo18 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -72720
         TabIndex        =   6
         Top             =   1320
         Width           =   4335
      End
      Begin VB.ComboBox Combo17 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -72720
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1920
         Width           =   2655
      End
      Begin VB.ComboBox Combo16 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -72720
         Style           =   1  'Simple Combo
         TabIndex        =   10
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Frame Frame6 
         Caption         =   "  Info ãÚáæãÇÊ "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1455
         Left            =   240
         TabIndex        =   19
         Top             =   480
         Width           =   5175
         Begin VB.ComboBox Combo11 
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
            Left            =   1680
            Locked          =   -1  'True
            Style           =   1  'Simple Combo
            TabIndex        =   24
            Top             =   600
            Width           =   2175
         End
         Begin VB.ComboBox Combo2 
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
            Left            =   1680
            Locked          =   -1  'True
            Style           =   1  'Simple Combo
            TabIndex        =   21
            Top             =   240
            Width           =   2175
         End
         Begin VB.ComboBox Combo14 
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
            Left            =   1680
            Locked          =   -1  'True
            Style           =   1  'Simple Combo
            TabIndex        =   27
            Top             =   960
            Width           =   2175
         End
         Begin VB.Label Label33 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "ÑÞã ÇáÞíÏ "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   4080
            TabIndex        =   25
            Top             =   675
            Width           =   735
         End
         Begin VB.Label Label28 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Journal Entry No."
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
            Height          =   375
            Left            =   240
            TabIndex        =   23
            Top             =   675
            Width           =   1335
         End
         Begin VB.Label Label34 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "ÇáÊÇÑíÎ"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   3960
            TabIndex        =   22
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Transaction Date"
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
            Left            =   240
            TabIndex        =   20
            Top             =   320
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Ticket No."
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
            Left            =   240
            TabIndex        =   26
            Top             =   1005
            Width           =   1095
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "ÑÞã ÇáÞíÏ "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   4080
            TabIndex        =   28
            Top             =   1005
            Width           =   735
         End
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   1920
         Top             =   4635
      End
      Begin VB.TextBox Text6 
         Height          =   360
         Left            =   600
         TabIndex        =   70
         Text            =   "Text6"
         Top             =   1275
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   1320
         TabIndex        =   69
         Text            =   "when editing TN placed here"
         Top             =   1275
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.TextBox Text5 
         Height          =   360
         Left            =   3120
         TabIndex        =   68
         Text            =   "A/C# placed here when edit"
         Top             =   1275
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Frame Frame1 
         Caption         =   " Total Today ÇÌãÇáí Çáíæãí "
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
         Height          =   1455
         Left            =   5640
         TabIndex        =   29
         Top             =   480
         Width           =   5775
         Begin VB.ComboBox Combo5 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   345
            Left            =   1440
            Locked          =   -1  'True
            Style           =   1  'Simple Combo
            TabIndex        =   33
            Text            =   "0.00"
            Top             =   600
            Width           =   2175
         End
         Begin VB.ComboBox Combo4 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   345
            ItemData        =   "BankTransaction1.frx":1092
            Left            =   1440
            List            =   "BankTransaction1.frx":1094
            Locked          =   -1  'True
            Style           =   1  'Simple Combo
            TabIndex        =   31
            Text            =   "0.00"
            Top             =   240
            Width           =   2175
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H80000014&
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
            Left            =   1440
            Locked          =   -1  'True
            Style           =   1  'Simple Combo
            TabIndex        =   36
            Text            =   "0"
            Top             =   960
            Width           =   2175
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " ÅÌãÇáí ÇáØÑÝ ÇáÏÇÆä"
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   3600
            TabIndex        =   34
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ÅÌãÇáí ÇáØÑÝ ÇáãÏíä"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   3600
            TabIndex        =   32
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ÚÏÏ ÇáÍÑßÇÊ "
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   3960
            TabIndex        =   37
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Total Credits"
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
            Left            =   240
            TabIndex        =   78
            Top             =   680
            Width           =   975
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Total Debits"
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
            Left            =   240
            TabIndex        =   30
            Top             =   320
            Width           =   1095
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Total Entries"
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
            Left            =   240
            TabIndex        =   35
            Top             =   1010
            Width           =   975
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "  De&scriptions æÕÝ "
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
         Height          =   735
         Left            =   240
         TabIndex        =   60
         Top             =   4200
         Width           =   6255
         Begin VB.ComboBox Combo8 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   240
            Style           =   1  'Simple Combo
            TabIndex        =   61
            Top             =   240
            Width           =   5895
         End
         Begin VB.Label Label23 
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "&Arabic"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   120
            TabIndex        =   67
            Top             =   720
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "  Credit  ÏÇÆä"
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
         Height          =   975
         Left            =   240
         TabIndex        =   49
         Top             =   3075
         Width           =   11175
         Begin MSMask.MaskEdBox MaskEdBox2 
            Height          =   350
            Left            =   8400
            TabIndex        =   59
            Top             =   480
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin VB.ComboBox Combo13 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2640
            TabIndex        =   55
            Top             =   480
            Width           =   5175
         End
         Begin VB.ComboBox Combo7 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   240
            TabIndex        =   52
            Top             =   480
            Width           =   2175
         End
         Begin VB.CommandButton Command5 
            BackColor       =   &H8000000C&
            Height          =   350
            Left            =   7800
            Picture         =   "BankTransaction1.frx":1096
            Style           =   1  'Graphical
            TabIndex        =   56
            Top             =   480
            Width           =   375
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ÅÓã ÇáÍÓÇÈ"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   6720
            TabIndex        =   54
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label14 
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "Account Na&me"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2640
            TabIndex        =   53
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "&Cr Amount(LFXCY)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   8400
            TabIndex        =   57
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "ÑÞã ÇáÍÓÇÈ"
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   1560
            TabIndex        =   51
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label17 
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "Account C&ode"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   240
            TabIndex        =   50
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ÇáãÈáÛ"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   10200
            TabIndex        =   58
            Top             =   240
            Width           =   735
         End
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   -64200
         Top             =   5760
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   15
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "BankTransaction1.frx":1630
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "BankTransaction1.frx":168E
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "BankTransaction1.frx":16EC
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "BankTransaction1.frx":1C2E
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "BankTransaction1.frx":2170
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "BankTransaction1.frx":25C2
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "BankTransaction1.frx":2A14
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "BankTransaction1.frx":2E66
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "BankTransaction1.frx":32B8
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "BankTransaction1.frx":370A
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "BankTransaction1.frx":3B5C
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "BankTransaction1.frx":3CB6
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "BankTransaction1.frx":3F68
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "BankTransaction1.frx":43BA
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "BankTransaction1.frx":44BC
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   7710
         Left            =   -74920
         TabIndex        =   66
         Top             =   400
         Width           =   11580
         _ExtentX        =   20426
         _ExtentY        =   13600
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         PictureAlignment=   4
         TextBackground  =   -1  'True
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
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   12
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "aa"
            Text            =   "!"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Ticket #"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Key             =   "a"
            Text            =   "AccountCode"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Key             =   "b"
            Text            =   "Account Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Key             =   "c"
            Text            =   "TransDate"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Key             =   "d"
            Text            =   "Debit"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Key             =   "e"
            Text            =   "Credit"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "FXCY Amount"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Key             =   "f"
            Text            =   "Journal #"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Key             =   "h"
            Text            =   "Descriptions"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Classification"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Prepby"
            Object.Width           =   2540
         EndProperty
         Picture         =   "BankTransaction1.frx":490E
      End
      Begin ComCtl3.CoolBar CoolBar1 
         Height          =   420
         Left            =   9000
         TabIndex        =   62
         Top             =   4395
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   741
         BandCount       =   2
         _CBWidth        =   2415
         _CBHeight       =   420
         _Version        =   "6.7.8988"
         MinHeight1      =   360
         Width1          =   2880
         NewRow1         =   0   'False
         MinHeight2      =   360
         Width2          =   1440
         NewRow2         =   0   'False
         Begin MSComctlLib.Toolbar Toolbar1 
            Height          =   330
            Left            =   120
            TabIndex        =   63
            Top             =   45
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   582
            ButtonWidth     =   609
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            Style           =   1
            ImageList       =   "ImageList1"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   7
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Add"
                  Object.ToolTipText     =   "Add to listview"
                  ImageIndex      =   13
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Save"
                  Object.ToolTipText     =   "Save to file"
                  ImageIndex      =   10
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Print"
                  Object.ToolTipText     =   "Print"
                  ImageIndex      =   3
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Delete"
                  Object.ToolTipText     =   "Delete"
                  ImageIndex      =   6
               EndProperty
               BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Return"
                  Object.ToolTipText     =   "Return and close this window"
                  ImageIndex      =   11
               EndProperty
               BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
            EndProperty
         End
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   2760
         Top             =   1155
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
               Picture         =   "BankTransaction1.frx":15DB4
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "BankTransaction1.frx":16206
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "BankTransaction1.frx":16658
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "BankTransaction1.frx":16AAA
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "BankTransaction1.frx":16DC4
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame2 
         Caption         =   "  Debit ãÏíä"
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
         Height          =   975
         Left            =   240
         TabIndex        =   38
         Top             =   1995
         Width           =   11175
         Begin MSMask.MaskEdBox MaskEdBox1 
            Height          =   350
            Left            =   8400
            TabIndex        =   47
            Top             =   480
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin VB.ComboBox Combo12 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2640
            TabIndex        =   44
            Top             =   480
            Width           =   5175
         End
         Begin VB.ComboBox Combo6 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   240
            TabIndex        =   41
            Top             =   480
            Width           =   2175
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H8000000A&
            Height          =   340
            Left            =   7800
            Picture         =   "BankTransaction1.frx":17216
            Style           =   1  'Graphical
            TabIndex        =   45
            ToolTipText     =   "Find Accounts"
            Top             =   480
            Width           =   375
         End
         Begin VB.Label Label35 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ÅÓã ÇáÍÓÇÈ"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   6240
            TabIndex        =   43
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ÇáãÈáÛ"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   10080
            TabIndex        =   48
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label32 
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "Account &Name"
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
            Left            =   2640
            TabIndex        =   42
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ÑÞã ÇáÍÓÇÈ"
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   1560
            TabIndex        =   40
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "&Db Amount(LXCY)"
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
            Left            =   8400
            TabIndex        =   46
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label12 
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "Acco&unt Code"
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
            Left            =   240
            TabIndex        =   39
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "  Journal Entries ÇÎÇá ÚÇã "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   2895
         Left            =   240
         TabIndex        =   65
         Top             =   5010
         Width           =   11175
         Begin MSComctlLib.ListView ListView1 
            Height          =   2535
            Left            =   240
            TabIndex        =   64
            Top             =   240
            Width           =   10695
            _ExtentX        =   18865
            _ExtentY        =   4471
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
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   11
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "!"
               Object.Width           =   529
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "TN"
               Object.Width           =   882
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "AccountCode"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "AccountNumber"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "TransDate"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Text            =   "DebitAmt"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Text            =   "CreditAmt"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   7
               Text            =   "FXCY Amount"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "JournalNo"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "Descriptions"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   10
               Text            =   "Classification"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.TextBox Text3 
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
         Left            =   240
         TabIndex        =   73
         Text            =   "this identify if not bal"
         Top             =   6960
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox Text1 
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
         TabIndex        =   72
         Text            =   "Text1"
         Top             =   6600
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox Text2 
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
         TabIndex        =   71
         Text            =   "Text2"
         Top             =   6240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox Combo3 
         Height          =   360
         Left            =   1200
         TabIndex        =   74
         Top             =   6840
         Visible         =   0   'False
         Width           =   615
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   4890
         Left            =   -74880
         TabIndex        =   17
         Top             =   3360
         Width           =   11580
         _ExtentX        =   20426
         _ExtentY        =   8625
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "TransDate"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Ref No"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "GL Acct No"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Bank Account Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Transaction Amt"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Transaction Type"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Cheque No."
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "ClearingDate"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "ReceiptNo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Payee/Payor"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   -65160
         TabIndex        =   80
         Text            =   "For Cr acct No"
         Top             =   2880
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox Text7 
         Height          =   360
         Left            =   -65160
         TabIndex        =   79
         Text            =   "Serialno is place here"
         Top             =   2880
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox Text9 
         Height          =   375
         Left            =   -65160
         TabIndex        =   81
         Text            =   "For AcctNAme"
         Top             =   2880
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label39 
         Caption         =   "ÇÓã ÇáÍÓÇÈ ÈÇáÚÑÈí "
         Height          =   255
         Left            =   -68040
         TabIndex        =   95
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label38 
         Caption         =   "ÇÓã ÇáÍÓÇÈ ÈÇáÇäÌáíÒí"
         Height          =   255
         Left            =   -68280
         TabIndex        =   94
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label37 
         Alignment       =   1  'Right Justify
         Caption         =   "ÇÎÊÇÑ ÑÞã ÇáÍÓÇÈ "
         Height          =   255
         Left            =   -69840
         TabIndex        =   93
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label31 
         Caption         =   "Receipt No.   ÑÞã ÇáÞÈÖ "
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -66600
         TabIndex        =   88
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label30 
         Caption         =   "Transaction Date       ÊÇÑíÎ ÇáÇíÏÇÚ "
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -69840
         TabIndex        =   84
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label22 
         Caption         =   "Deposit Slip No.       ÑÞã ÇáÇíÏÇÚ "
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -69600
         TabIndex        =   82
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label29 
         Caption         =   "Select  GL  Account  No"
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
         Left            =   -74760
         TabIndex        =   1
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label27 
         Caption         =   "Account Name in Eng"
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
         Left            =   -74760
         TabIndex        =   3
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label26 
         Caption         =   "Account Name in Arab"
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
         Left            =   -74760
         TabIndex        =   5
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label25 
         Caption         =   "Transaction Type       äæÚ ÇáÚãáíÉ "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74760
         TabIndex        =   7
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label24 
         Caption         =   "Transaction Amount      ãÈáÛ ÇáÚãáíÉ "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74760
         TabIndex        =   9
         Top             =   2400
         Width           =   1455
      End
   End
   Begin VB.Label Label36 
      Caption         =   "Label36"
      Height          =   495
      Left            =   5160
      TabIndex        =   92
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Menu main 
      Caption         =   "Main"
      Visible         =   0   'False
      Begin VB.Menu xCancel 
         Caption         =   "Cancel"
      End
      Begin VB.Menu eEdit 
         Caption         =   "Modify/Edit  ÇÖÇÝÉ "
      End
      Begin VB.Menu bar3 
         Caption         =   "-"
      End
      Begin VB.Menu xFind 
         Caption         =   "Find... íÌÏ"
         Visible         =   0   'False
      End
      Begin VB.Menu xPreview 
         Caption         =   "Print Preview...ÇáãÚÇíäÉ ÞÈá ÇáØÈÇÚÉ "
      End
      Begin VB.Menu xREfresh 
         Caption         =   "Refresh ÊÍÏíË"
      End
      Begin VB.Menu bar1 
         Caption         =   "-"
      End
      Begin VB.Menu xView 
         Caption         =   "View íÑí"
         Begin VB.Menu Li 
            Caption         =   "Large Icon ÕäÝ ßÈíÑ "
         End
         Begin VB.Menu SI 
            Caption         =   "Small Icon ÕäÝ ÕÛíÑ "
         End
         Begin VB.Menu xList 
            Caption         =   "List ÞÇÆãÉ "
         End
         Begin VB.Menu xDEtails 
            Caption         =   "Details æÕÝ "
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu POstToGL 
         Caption         =   "Post to GL...áÕÞ "
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu Main2 
      Caption         =   "Main2"
      Visible         =   0   'False
      Begin VB.Menu DElSelectitem 
         Caption         =   "Delete Selected Item ÇáÛÇÁ ÊÌãíÚ ÇáÃÕäÇÝ "
      End
      Begin VB.Menu PrintItem 
         Caption         =   "Print items ØÈÇÚÉ ÇáÇÕäÇÝ "
      End
      Begin VB.Menu Fv 
         Caption         =   "Full View ÑÄí ÚÇãÉ "
      End
   End
   Begin VB.Menu main3 
      Caption         =   "main3"
      Visible         =   0   'False
      Begin VB.Menu XdeleteMain3 
         Caption         =   "Delete"
      End
      Begin VB.Menu xSelectall 
         Caption         =   "Select All"
      End
   End
End
Attribute VB_Name = "BankTransaction1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FirstTimeLoad As Boolean   ' control for activating this form
Dim xdecimal As Integer                 ' to control decimal point in maskedit
Dim xCtrlKeyPress As Integer        ' this is used for a search engine key. When it is press searching of account will be shifted to account name or
Dim pressno As Integer
Dim Accts As ADODB.Recordset
Dim sqltable As Boolean
Dim mConnek As Boolean
Dim xtable As String
Dim CON1 As ADODB.Connection
Dim opensearch As Boolean

Dim CheckOwnedby As String
Dim Payee As String
Dim CheckNo As String
Dim BankTRansdate As String
Dim Checktype As String
Dim Trantype As String
Dim ORno As String
Dim Codetype As String

Dim catName As String
Dim DrCat As String
Dim CrCat As String
Dim MItem As ListItem
Dim xcol As ColumnHeader
Dim NextTn As Long
Dim acctnames As New ADODB.Recordset
Dim AcctCode As New ADODB.Recordset
Dim rstTemp As New ADODB.Recordset
Dim TAbClic As Integer
Dim xCtrlKeyPress2 As Integer       ' account name to account Number.
Sub PrintEntries()
           Printer.Orientation = 1
           Printer.Print
           Printer.Print
           Printer.Print
           Printer.Print
           
           Printer.FontSize = 12
           Printer.FontName = "Times new Roman"
           Printer.Print ; Tab(8); "Journal No: "; Me.Combo11
           Printer.Print
           Printer.FontSize = 10
           Printer.FontName = "Arabic Transparent"
           Printer.Print ; Tab(10); Date
           Printer.Print ; Tab(10); Time
           Printer.FontName = "Times new Roman"
           Printer.Print ; Tab(10); "-----------------------------------------------------------------------------------------------------------------------------------------------------------"
           i = 0
           Dim cR As Currency
           Dim dr As Currency
           Dim TotalDb  As Currency
           Dim TotalCr As Currency
           For i = 1 To Me.ListView1.ListItems.Count
                  TN = Me.ListView1.ListItems.Item(i).SubItems(1)
                  Ac = Me.ListView1.ListItems.Item(i).SubItems(2)
                  an = Me.ListView1.ListItems.Item(i).SubItems(3)
                  dr = FormatNumber(Me.ListView1.ListItems.Item(i).SubItems(5), 2, vbTrue, vbTrue, vbTrue)
                  cR = FormatNumber(Me.ListView1.ListItems.Item(i).SubItems(6), 2, vbTrue, vbTrue, vbTrue)
                  Jn = Me.ListView1.ListItems.Item(i).SubItems(7)
                  descr = Me.ListView1.ListItems.Item(i).SubItems(8)
                  'TRantYpe = Me.ListView1.ListItems.Item(i).SubItems(9)
                  Class = Trim(Me.ListView1.ListItems.Item(i).SubItems(9))
                  TotalDb = TotalDb + dr
                  TotalCr = TotalCr + cR
                  
                  Printer.Print ; Tab(10); "Classification :" & Right(Class, 60)
                  Printer.Print ; Tab(10); an; Tab(50); Ac _
                              ; Tab(115 - Len(dr)); IIf(dr <> 0, Format(dr, "###,###,###.#0"), "") _
                              ; Tab(130 - Len(cR)); IIf(cR <> 0, Format(cR, "###,###,###.#0"), "")
                  Printer.FontName = "Arabic Transparent"
                  Printer.Print ; Tab(10); "Desc : " & descr
                  Printer.FontName = "Times new Roman"
                  If i = Me.ListView1.ListItems.Count Then
                    Printer.Print ; Tab(10); "-----------------------------------------------------------------------------------------------------------------------------------------------------------"
                   Else
                   Printer.Print ""
                  End If
           Next i
           
           Printer.Print ; Tab(10); "Totals: " & (i - 1); Tab(115 - Len(TotalDb)); Format(TotalDb, "###,###,###.#0"); Tab(130 - Len(TotalCr)); Format(TotalCr, "###,###,###.#0")
           Printer.Print ; Tab(10); "========================================================================================"
           Printer.Print
           Printer.Print
           Printer.Print ; Tab(10); "Prepared by: "; cLogUser
           Printer.EndDoc
End Sub
Sub DeleteItems()
If Me.ListView1.ListItems.Count <> 0 Then
 xmsg = MsgBox("Delete TN " & Me.ListView1.SelectedItem & "?", vbQuestion + vbYesNo + vbDefaultButton2, "Please confirm")
        If xmsg = vbYes Then
            TN = Me.ListView1.SelectedItem.SubItems(1)
            Dim xtemp As New ADODB.Recordset
            xtemp.Open "delete TempBankJournal  where ticket=" & "'" & TN & "'", constring, adOpenKeyset, adLockOptimistic, adCmdText
            IndeOnSelectedItem = Me.ListView1.SelectedItem.SubItems(1)
            'totalItem = Me.ListView1.ListItems.Count
            cindex = Val(Me.ListView1.SelectedItem.Index)
            If Val(Me.ListView1.ListItems.Item(1).SubItems(1)) <> 1 Then
              Newtn = Val(Me.ListView1.ListItems.Item(1).SubItems(1))
              xItem = Val(cindex)
             Else
              xItem = 1
            End If
            Me.ListView1.ListItems.Remove cindex
            totalItem = Me.ListView1.ListItems.Count
            TN = 0
            Me.Combo6 = ""
            Me.Combo12 = ""
            Me.MaskEdBox1.Text = ""
            Me.Combo8 = ""
            Me.Combo7 = ""
            Me.Combo13 = ""
            Me.MaskEdBox2.Text = ""
            Me.Combo12 = ""
                      
            Dim TN1 As Long
            'Xdate = Format(Date, "dd/mm/yyyy")
            'xTemp.Open "SElect count(ticket) as [tn] from BankJournal where transdate=" & "'" & Xdate & "'", conString, adOpenDynamic, adLockOptimistic, adCmdText
            'If Me.Text4 = "Edit" Then 'dont change the TN if it is editing
                'TN1 = Me.ListView1.ListItems.Count
             ' Else
                'TN1 = xTemp!TN '+ 1
            'End If
            'xTemp.Close
            
           'it is not necessary to to add if it the last item in the list equal to listcount
            If totalItem + 1 <> Val(IndeOnSelectedItem) Then
            
                 i = 0
                 
                 For i = 1 To totalItem 'Me.ListView1.ListItems.Count
                    TN = i 'TN + 1 'i ' Me.ListView1.ListItems.Item(i).SubItems(1)\
                    Dim GetNewTn As Boolean
                    'If GetNewTn = False Then
                      'If Val(Me.ListView1.ListItems.Item(1).SubItems(1)) <> 1 Then
                        'newtn = Val(Me.ListView1.ListItems.Item(1).SubItems(1))
                     '   GetNewTn = True
                      '  If xitem > 1 Then
                      '   TN = Newtn + 1
                      '   Newtn = Newtn + 1
                      '   Else
                      '   TN = IIf(IsEmpty(Newtn) = True, 1, Newtn)
                      '  End If
                     'Else
                     'Newtn = IIf(IsEmpty(Newtn) = True, TN, Newtn + 1)
                     'TN = Newtn
                    'End If
                    'were going to delete all the items in the list one by
                    'one thats why we place the value of each item into memory variables
                    Ac = Me.ListView1.ListItems.Item(i).SubItems(2)
                    an = Me.ListView1.ListItems.Item(i).SubItems(3)
                    dr = Me.ListView1.ListItems.Item(i).SubItems(5)
                    cR = Me.ListView1.ListItems.Item(i).SubItems(6)
                    Jn = Me.ListView1.ListItems.Item(i).SubItems(7)
                    'fcy = Me.ListView1.ListItems.Item(i).SubItems(8)
                    descr = Me.ListView1.ListItems.Item(i).SubItems(8)
                    xTN = Me.ListView1.ListItems.Item(i).SubItems(1)
                    NextTn = TN
                    'cindex = Val(Me.ListView1.ListItems.Item(i).Index)
                    
                    'Me.ListView1.ListItems.Remove cindex
                    xtemp.Open "delete TempBankJournal  where ticket=" & "'" & xTN & "'", constring, adOpenKeyset, adLockOptimistic, adCmdText
                    With rstTemp
                            .addnew
                            !ticket = NextTn
                            !accountnumber = Ac
                            !accountname = an
                            !TRansDate = Me.Combo2
                            !DebitAmount = dr
                            !creditamount = cR
                            !fcdebit = fcy
                            !FcCredit = fcy
                            !SerialNo = Jn
                            !Description = descr
                            !deletemark = 0
                            !Status = "Unposted"
                            .Update
                      End With
                    Next
                        'add again but with the new TN
                       Me.ListView1.ListItems.clear
                       rstTemp.close
                       rstTemp.Open "TempBankJournal", constring, adOpenKeyset, adLockOptimistic, adCmdTable
                       While rstTemp.EOF = False
                         Set MItem = Me.ListView1.ListItems.Add(, , rstTemp!ticket, , 1)
                         MItem.SubItems(1) = rstTemp!ticket
                         MItem.SubItems(2) = rstTemp!accountnumber
                         MItem.SubItems(3) = rstTemp!accountname
                         MItem.SubItems(4) = rstTemp!TRansDate
                         MItem.SubItems(5) = rstTemp!DebitAmount
                         MItem.SubItems(6) = rstTemp!creditamount
                         MItem.SubItems(7) = rstTemp!SerialNo
                         'MItem.SubItems(8) = fcy
                         MItem.SubItems(8) = rstTemp!Description
                         rstTemp.MoveNext
                        Wend
                    
                 Else
                 
                 xtemp.Open "delete TempBankJournal  where ticket=" & "'" & xTN & "'", constring, adOpenKeyset, adLockOptimistic, adCmdText
                End If
                    ListView1.SortKey = 1
                    ListView1.Sorted = True
          End If
          Exit Sub
   End If
 End Sub
 Sub DisplayTransToday()
 Me.ListView2.ListItems.clear
 Dim rstTRansToday As New ADODB.Recordset
 rstTRansToday.Open "Select * from BankJournal where remarks is null order by serialno", CON1, adOpenKeyset, adLockOptimistic, adCmdText
       If rstTRansToday.EOF = False Then
        rstTRansToday.MoveFirst
       End If
        While rstTRansToday.EOF = False
          On Error Resume Next
          Set MItem = Me.ListView2.ListItems.Add(, , rstTRansToday!ticket, , 1)
          MItem.SubItems(1) = rstTRansToday!ticket
          MItem.SubItems(2) = rstTRansToday!accountnumber
          MItem.SubItems(3) = rstTRansToday!accountname
          MItem.SubItems(4) = rstTRansToday!TRansDate
          MItem.SubItems(5) = Format(rstTRansToday!DebitAmount, "###,###,###.#0")
          MItem.SubItems(6) = Format(rstTRansToday!creditamount, "###,###,###.#0")
          MItem.SubItems(7) = Format(rstTRansToday!fxcyAmount, "###,###,###.#0")
          MItem.SubItems(8) = rstTRansToday!SerialNo
          MItem.SubItems(9) = rstTRansToday!Description
          MItem.SubItems(10) = rstTRansToday!Classification
          MItem.SubItems(11) = rstTRansToday!Prepby
          rstTRansToday.MoveNext
       Wend
       
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


Private Sub Check2_Click()
If Me.Check2.Value = 0 Then
    Me.Command1.Enabled = True
    Me.Label31.Enabled = True
    Me.Combo21.Enabled = True
    PaymentList.Show
  Else
    Me.Command1.Enabled = False
    Me.Label31.Enabled = False
    Me.Combo21.Enabled = False
   End If
End Sub

Private Sub Combo10_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Dim strFindMe As String
    Dim intSelectedOption As String
    Dim itmFnd As ListItem   ' FoundItem variable.
    intSelectedOption = lvwSubItem
    strFindMe = Trim(Me.Combo10)
    cLen = Len(strFindMe)
    Set itmFnd = Me.ListView3.Finditem(strFindMe, intSelectedOption, , Left(lvwPartial, cLen))
     If itmFnd Is Nothing Then  ' If no match, inform user and exit.
       msg = MsgBox("Search Text is not found", vbExclamation + vbOKOnly, "Find")
       Me.Combo10.SetFocus
       Exit Sub
       Else
         
        itmFnd.EnsureVisible
        itmFnd.Selected = True
        Me.ListView3.SetFocus
     End If
End If
End Sub

Private Sub Combo12_Click()
X = 0
For X = 1 To Len(Me.Combo12)
    If Mid(Trim(Me.Combo12), X, 1) = "\" Then Exit For
     xname = xname & Mid(Me.Combo12, X, 1)
Next
On Error GoTo RefresTable
c = Err.Number
acctnames.MoveFirst
While acctnames.EOF = False
    If xname = Trim(acctnames!accountnameeng) Then
       Me.Combo6 = acctnames!AccountCode
    End If
    acctnames.MoveNext
Wend
RefresTable:
If c = 3704 Then
  Dim xClass As New HabitatClass
  xtable = xtable = "Select * from FinanceMaster order by AccountName"
  acctnames.Open "FinanceMaster", constring, adOpenKeyset, adLockPessimistic, adCmdText
  While acctnames.EOF = False
    Me.Combo12.AddItem acctnames!accountnameeng
    acctnames.MoveNext
  Wend
 End If
 
Dim catName As String
Dim Prevcap As String
Dim acctNo As String
Acctname = InStr(1, Trim(Me.Combo12), "\", vbTextCompare)
acctNo = Left(Trim(Me.Combo12), (Acctname) - 1)
Prevcap = Trim(Me.caption)
Call DisplayCatsName(Prevcap, acctNo, catName)
Me.caption = "Bank Transactions" & "//" & catName
DrCat = catName
End Sub

Private Sub Combo12_GotFocus()
If Me.Combo6 = "" Then
  Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
  Tmp = SendMessage(Combo12.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End If
End Sub

Private Sub Combo12_KeyPress(KeyAscii As Integer)
Dim rst As ADODB.Recordset
Set rst = New ADODB.Recordset

If KeyAscii = 13 Then
   If Me.MaskEdBox1.Enabled = True Then
       Me.MaskEdBox1.SetFocus
      Else
     Me.Combo7.SetFocus
    End If
End If
End Sub

Private Sub Combo12_LostFocus()
'Call Combo6_LostFocus
End Sub

Private Sub Combo13_Click()
X = 0
For X = 1 To Len(Me.Combo13)
    If Mid(Trim(Me.Combo13), X, 1) = "\" Then Exit For
     xname = xname & Mid(Me.Combo13, X, 1)
Next


On Error GoTo RefresTable
c = Err.Number
acctnames.MoveFirst
While acctnames.EOF = False
    If Trim(xname) = Trim(acctnames!accountnameeng) Then
       Me.Combo7 = acctnames!AccountCode
    End If
    acctnames.MoveNext
Wend
RefresTable:
 If c = 3704 Then
    Dim xClass As New HabitatClass
    xtable = "Select * from FinanceMaster order by AccountName"
    acctnames.Open "FinanceMaster", constring, adOpenKeyset, adLockPessimistic, adCmdText
    While acctnames.EOF = False
       Me.Combo12.AddItem acctnames!accountnameeng
       acctnames.MoveNext
    Wend
 End If
 
Dim catName As String
Dim Prevcap As String
Dim acctNo As String
Acctname = InStr(1, Trim(Me.Combo13), "\", vbTextCompare)
acctNo = Left(Trim(Me.Combo13), (Acctname) - 1)
Prevcap = Trim(Me.caption)
Call DisplayCatsName(Prevcap, acctNo, catName)
Me.caption = "Bank Transactions" & "//" & catName
CrCat = catName
End Sub

Private Sub Combo13_GotFocus()
If Me.Combo7 = "" Then
  Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
  Tmp = SendMessage(Combo13.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End If
End Sub

Private Sub Combo13_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
   If Me.MaskEdBox2.Enabled = True Then
      Me.MaskEdBox2.SetFocus
     Else
    Me.Combo8.SetFocus
   End If
End If
End Sub

Private Sub Combo13_LostFocus()
'Call Combo7_LostFocus
End Sub

Private Sub Combo15_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Me.Combo16.SetFocus
End If
End Sub

Private Sub Combo16_GotFocus()

Dim havedot As Boolean
  For i = 1 To Len(Trim(Me.Combo16.Text))
        Havedott = Mid(Me.Combo16.Text, i, 1)
        If Havedott = "." Then Exit For
           
  Next
  If Havedott = "." Then
     xdecimal = 1
     Exit Sub
    Else
      xdecimal = 0
  End If

i = 0
For i = 1 To Len(Trim(Me.Combo16.Text))
      X = Mid(Trim(Me.Combo16.Text), i, 1)
      If X = "." Then
        havedot = True
        Exit For
       Else
        havedot = False
      End If
Next i

End Sub

Sub SaveTRan()
Dim rsBankTRan As New ADODB.Recordset
rsBankTRan.Open "Select count(*) as cTotals  from BankTRansaction where AccountCode=" & "'" & Trim(Me.Combo20) & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
cTOtals = rsBankTRan!cTOtals + 1

rsBankTRan.close
rsBankTRan.Open "BankTRansaction", constring, adOpenKeyset, adLockPessimistic, adCmdTable
With rsBankTRan
    .addnew
    If Left(Me.Combo17, 2) <> "DC" Then
       !TRansDate = Me.MaskEdBox3.Text 'Format(Date, "mm/dd/yyyy")
      Else
      !TRansDate = Me.MaskEdBox3.Text 'Format(Date, "mm/dd/yyyy")
     End If
    !AccountCode = Trim(Me.Combo20)
    !NameEng = Me.Combo19
    !RefNo = cTOtals
    !NameArab = Me.Combo18
    !TransAmt = Me.Combo16
    !Trantype = Me.Combo17
    If Left(Me.Combo17, 2) <> "DC" Then
     !Payno_Or_ChekNo = Trim(Me.Text7)
     !CrAcctCode = Trim(Me.Text8)
     !CrAcctName = Trim(Me.Text9)
     Else
       If Me.Check2.Value = 0 Then
         !DirectDeposit = 1
         !Payno_Or_ChekNo = Me.Combo15
         !CrAcctCode = "111020105001"
         !CrAcctName = "Deposit-in-Transit Control Account"
         !receiptno = Me.Combo21
        Else
        !Payno_Or_ChekNo = Me.Combo15
       End If
    End If
    
        
    .Update
End With

Trancode = Left(Me.Combo9, 2)
If Trancode = "DQ" Or Trancode = "WQ" Then
 Me.ListView3.ListItems.clear
End If
Set MItem = Me.ListView3.ListItems.Add(, , Format(BankTRansdate, "dd/mm/yyyy"))
MItem.SubItems(1) = cTOtals
MItem.SubItems(2) = Me.Combo20
MItem.SubItems(3) = Me.Combo19
MItem.SubItems(4) = Me.Combo16
MItem.SubItems(5) = Me.Combo17

If Left(Me.Combo17, 2) <> "DC" Then
  MItem.SubItems(6) = Me.Text7
  MItem.SubItems(7) = Me.Text8 & "-" & Me.Text9
 Else
  If Me.Check2.Value = 0 Then
         '!DirectDeposit = 1
         MItem.SubItems(6) = Me.Combo15
         MItem.SubItems(7) = "111020105001-Deposit-in-Transit Control Account"
         MItem.SubItems(8) = Me.Combo21
        Else
        MItem.SubItems(6) = Me.Combo15
   End If
End If
Me.Combo16 = ""
Me.Combo18 = ""
Me.Combo19 = ""
Me.Combo20 = ""
Me.Combo20.SetFocus
End Sub
Private Sub Combo16_KeyPress(KeyAscii As Integer)



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
 If Me.Combo16.Text <> " " Then
  xdecimal = 0
  End If
End If

If KeyAscii = 13 Then
     If Me.MaskEdBox3.Enabled = True Then
        Me.MaskEdBox3.SetFocus
     Else
        Call SaveTRan
     End If
End If
End Sub

Private Sub Combo16_KeyUp(KeyCode As Integer, Shift As Integer)
Dim havedot As Boolean
  For i = 1 To Len(Trim(Me.Combo16.Text))
        Havedott = Mid(Me.Combo16.Text, i, 1)
        If Havedott = "." Then Exit For
           
  Next
  If Havedott = "." Then
     xdecimal = 1
     Exit Sub
    Else
      xdecimal = 0
  End If

i = 0
For i = 1 To Len(Trim(Me.Combo16.Text))
      X = Mid(Trim(Me.Combo16.Text), i, 1)
      If X = "." Then
        havedot = True
        Exit For
       Else
        havedot = False
      End If
Next i

End Sub

Private Sub Combo16_LostFocus()
xdecimal = 0
Me.Combo16.Text = Format(Me.Combo16.Text, "###,###,###,###.#0")

End Sub

Private Sub Combo17_Click()
If Left(Me.Combo17, 2) = "TO" Then
  BankTransaction1.Combo16.Locked = True
   BankTransaction1.ListView3.ColumnHeaders(7).Text = "Payment No."
   BankTransaction1.ListView3.ColumnHeaders(8).Text = "Cr AcctCode"
   Me.Command1.caption = "Show Payment List..."
   Me.Check2.Enabled = False
   Me.Check2.Value = 0
   PaymentList.Show
 Else
 Me.ListView3.ColumnHeaders(7).Text = "Checque No."
 BankTransaction1.ListView3.ColumnHeaders(8).Text = "ClearingDate"
 BankTransaction1.Combo16.Locked = False
End If
If Left(Me.Combo17, 2) = "DC" Then
   Me.Command1.caption = "Show Cash Dep List..."
   BankTransaction1.ListView3.ColumnHeaders(7).Text = "Dep Slip No."
   BankTransaction1.ListView3.ColumnHeaders(8).Text = "Cr AcctCode"
   
End If


If Left(Me.Combo17, 2) = "DC" Then
    Me.Label22.Enabled = True
    Me.Label30.Enabled = True
    Me.MaskEdBox3.Enabled = True
    Me.Combo15.Enabled = True
    Me.Check2.Enabled = True
    If Me.Check2.Value = 0 Then
        Me.Command1.Enabled = True
    End If
   Else
     Me.Label22.Enabled = False
    Me.Label30.Enabled = False
    Me.MaskEdBox3.Enabled = False
    Me.Combo15.Enabled = False
    Me.Check2.Enabled = False
    Me.Command1.Enabled = False
End If
End Sub

Private Sub Combo17_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 If Me.Combo15.Enabled = True Then
    Me.Combo15.SetFocus
   Else
   Me.Combo16.SetFocus
  End If
End If
End Sub

Private Sub Combo18_Click()
AcctCode.MoveFirst
While AcctCode.EOF = False
    If UCase(Trim(Me.Combo18)) = UCase(Trim(AcctCode!accountnamearab)) Then
       Me.Combo20 = AcctCode!AccountCode
       Me.Combo19 = AcctCode!accountnameeng
       Me.Combo18 = AcctCode!accountnamearab
    End If
    AcctCode.MoveNext
Wend
Call Combo20_Click
End Sub

Private Sub Combo19_Click()
xAccount = Trim(Me.Combo19)
Dim RstBA As New ADODB.Recordset
RstBA.Open "Select * from FinanceMaster where accountNameEng=" & "'" & xAccount & "'", constring, adOpenDynamic, adLockOptimistic, adCmdText
If RstBA.EOF = False Then
  Me.Combo20 = RstBA!AccountCode
  Me.Combo18 = RTrim(RstBA!accountnamearab)
 Else
 Me.Combo20 = "Not found select again"
 Me.Combo18 = "Not found select again"
End If
RstBA.close
Call Combo20_Click
End Sub

Private Sub Combo20_Click()
xAccount = Trim(Me.Combo20)
Dim RstBA As New ADODB.Recordset
RstBA.Open "Select * from FinanceMaster where accountCode=" & "'" & xAccount & "'", constring, adOpenDynamic, adLockOptimistic, adCmdText
If RstBA.EOF = False Then
  Me.Combo19 = RstBA!accountnameeng
  Me.Combo18 = RTrim(RstBA!accountnamearab)
 Else
 Me.Combo18 = "Not found select again"
 Me.Combo19 = "Not found select again"
End If
RstBA.close
Dim rsFilter1 As New ADODB.Recordset
If Trim(Me.Combo9) = "[All]" Then
  rsFilter1.Open "seLect * from BankTransaction where status is null and AccountCode =" & "'" & Trim(Combo20) & "'" & " order by transdate", constring, adOpenKeyset, adLockPessimistic, adCmdText
 Else
  rsFilter1.Open "seLect * from BankTransaction where status is null and AccountCode =" & "'" & Trim(Combo20) & "'" & " and LefT(tranType,2)=" & "'" & Left(Me.Combo9, 2) & "'" & "order by transdate", constring, adOpenKeyset, adLockPessimistic, adCmdText
End If
Me.ListView3.ListItems.clear
Do Until rsFilter1.EOF = True
  On Error Resume Next
   Set MItem = Me.ListView3.ListItems.Add(, , Format(rsFilter1!TRansDate, "dd/mm/yyyy"))
    MItem.SubItems(1) = rsFilter1!RefNo
    MItem.SubItems(2) = rsFilter1!AccountCode
    MItem.SubItems(3) = rsFilter1!NameEng
    MItem.SubItems(4) = Format(rsFilter1!TransAmt, "###,###,###.#0")
    MItem.SubItems(5) = rsFilter1!Trantype
    MItem.SubItems(6) = rsFilter1!Payno_Or_ChekNo
    If IsNull(rsFilter1!CrAcctCode) = False Then
     MItem.SubItems(7) = rsFilter1!CrAcctCode & "-" & rsFilter1!CrAcctName
    End If
    MItem.SubItems(8) = rsFilter1!receiptno
    rsFilter1.MoveNext
Loop
rsFilter1.close

End Sub

Private Sub Combo20_GotFocus()

Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
  Tmp = SendMessage(Combo20.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)

End Sub

Private Sub Combo20_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.Combo17.SetFocus
End If
End Sub

Private Sub Combo23_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If Me.Combo6 <> "" Or Me.Combo7 <> "" Then
    xmsg = MsgBox("Are Entries Okay?", vbQuestion + vbOKCancel, "Please confirm")
    If xmsg = vbOK Then
        Dim strFindMe As String
        Dim itmFound As ListItem   ' FoundItem variable.
        intSelectedOption = lvwText ' lvwSubItem'lvwSubItem
        strFindMe = Trim(Me.Text4.Text)
        Set itmFound = Me.ListView1.Finditem(strFindMe, intSelectedOption, , lvwPartial)
        If itmFound Is Nothing Then  ' If no match, inform user and exit.
            AddEntries
        Else
            cindex = Val(Me.ListView1.SelectedItem.Index)
            Me.ListView1.ListItems.Remove cindex
            'remove also in temp table
            Dim xtemp As ADODB.Recordset
            Set xtemp = New ADODB.Recordset
           
            xtemp.Open "delete TempBankJournal  where ticket=" & "'" & strFindMe & "'", constring, adOpenKeyset, adLockOptimistic, adCmdText
            AddEntries
        End If
        Me.Combo6.SetFocus
    
      End If
    End If
 End If

End Sub

Private Sub Combo3_GotFocus()
Me.ListView1.SetFocus
End Sub

Private Sub Combo6_Click()
'we don't allow the user to input same account number
If Me.Combo6 <> "" And Me.Combo7 = "" Then
    If Me.Combo6 = Me.Combo7 Then
     xmsg = MsgBox("You have entered same Account number in Debit Side", vbInformation + vbOKOnly, "Message")
     Me.Combo6.SetFocus
     Exit Sub
    End If
End If


If Me.Combo6 = "" Then
 If Me.Combo6.Enabled <> False Then
   Me.Combo6.SetFocus
  Else
   Me.Combo7.SetFocus
 End If
End If


xAccount = Me.Combo6
Dim RstBA As ADODB.Recordset
Set RstBA = New ADODB.Recordset
xtable = "FinanceMaster"
xKey = "select * from " & xtable & " where " & _
       " AccountCode = " & "'" & xAccount & "'"
         
RstBA.Open xKey, constring, adOpenDynamic, adLockOptimistic, adCmdText
If RstBA.EOF = False Then
  Me.Combo12 = RstBA!accountnameeng & "\" & RTrim(RstBA!accountnamearab)
End If
RstBA.close
Dim catName As String
Dim Prevcap As String
Dim acctNo As String
acctNo = Trim(Me.Combo6)
Prevcap = Trim(Me.caption)
Call DisplayCats(Prevcap, acctNo, catName)
Me.caption = "Bank Transactions" & "//" & catName
DrCat = catName
End Sub

Private Sub Combo6_GotFocus()

Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
  Tmp = SendMessage(Combo6.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub Combo6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Me.Combo12.SetFocus
End If
End Sub

Private Sub Combo6_LostFocus()

'we don't allow the user to input same account number
If Me.Combo6 <> "" And Me.Combo7 <> "" Then
    If Me.Combo6 = Me.Combo7 Then
     xmsg = MsgBox("You have entered same Account number in Credit Side", vbInformation + vbOKOnly, "Message")
     Me.Combo6.SetFocus
     Exit Sub
    End If
End If

xAccount = Me.Combo6
If xAccount = "" Then
   Exit Sub
End If

If Me.Combo6 = "" Then
    Me.Combo6.SetFocus
     Exit Sub
End If

Dim RstBA As New ADODB.Recordset
xtable = "FinanceMAster"
xKey = "select * from " & xtable & " where " & _
       " AccountCode = " & "'" & xAccount & "'"

'Validate the entered account NUmber if it is in the right format.
RstBA.Open xKey, constring, adOpenDynamic, adLockOptimistic, adCmdText
If RstBA.EOF = True Then
    RstBA.close
    xmsg = MsgBox("Account Number does not exist", vbInformation + vbOKOnly, "Message")
    Me.Combo6.SetFocus
    Exit Sub
Else
  Me.Combo12 = RstBA!accountnameeng & "\" & RstBA!accountnamearab
End If

If Me.Combo6 <> Me.Text5.Text Then
    Me.Text4 = "Text"
    Exit Sub
End If

End Sub


Private Sub Combo9_Click()
Dim rsFilter1 As New ADODB.Recordset
Dim rsFilter2 As New ADODB.Recordset
Me.MaskEdBox1.Text = ""
Me.MaskEdBox2.Text = ""
Trancode = Left(Me.Combo9, 2)
On Error Resume Next
rsFilter1.close
rsFilter2.close
On Error GoTo 0
If Trim(Me.Combo9.Text) = "[All]" And Me.Combo20 = "" Then
   rsFilter1.Open "seLect * from BankTransaction where status is null order by transdate", constring, adOpenKeyset, adLockPessimistic, adCmdText
   rsFilter2.Open "seLect * from vouchers  where Left(Paymode,2)='03' or Left(Paymode,2)='10' and receiptdate > '01/21/2003'", constring, adOpenKeyset, adLockPessimistic, adCmdText
  ElseIf Trim(Me.Combo9.Text) = "[All]" And Me.Combo20 <> "" Then
   rsFilter1.Open "seLect * from BankTransaction where status is null and AccountCode=" & "'" & Trim(Me.Combo20) & "'" & " order by transdate", constring, adOpenKeyset, adLockPessimistic, adCmdText
  
  Else
    rsFilter1.Open "seLect * from BankTransaction where status is null and Left(tranType,2)=" & "'" & Trancode & "'" & "order by transdate", constring, adOpenKeyset, adLockPessimistic, adCmdText
    If Trancode = "DQ" Then
     'rsFilter2.Open "seLect * from vouchers where Left(Paymode,2)=03" & "and Left(Payopt,3)='006' and BankName is null and receiptdate > '01/01/2003' and  moderef<> 'Checks for Deposit'", constring, adOpenKeyset, adLockPessimistic, adCmdText
     rsFilter2.Open "SELECT * from vouchers WHERE     (svoucher = 'Collections') AND (PAYMODE = '03     Check' OR PAYMODE = '10     Returned Check') AND (deposit = '1') AND (bankname IS NULL)ORDER BY ReceiptNo", constring, adOpenKeyset, adLockPessimistic, adCmdText
    ElseIf Trancode = "WQ" Then
     rsFilter2.Open "seLect * from vouchers where Left(Paymode,2)=03" & "and BankName is not null and receiptdate > '01/01/2003'", constring, adOpenKeyset, adLockPessimistic, adCmdText
    ElseIf Trancode = "RQ" Then
     rsFilter2.Open "seLect * from vouchers where Left(Paymode,2)=10" & "and BankName is not null and receiptdate > '01/01/2003'", constring, adOpenKeyset, adLockPessimistic, adCmdText
     
   End If
End If

Me.ListView3.ListItems.clear
Do Until rsFilter1.EOF = True
  On Error Resume Next
   Set MItem = Me.ListView3.ListItems.Add(, , Format(rsFilter1!TRansDate, "dd/mm/yyyy"))
    MItem.SubItems(1) = rsFilter1!RefNo
    MItem.SubItems(2) = rsFilter1!AccountCode
    MItem.SubItems(3) = rsFilter1!NameEng
    MItem.SubItems(4) = Format(rsFilter1!TransAmt, "###,###,###.#0")
    MItem.SubItems(5) = rsFilter1!Trantype
    MItem.SubItems(6) = rsFilter1!Payno_Or_ChekNo
    If IsNull(rsFilter1!CrAcctCode) = False Then
     MItem.SubItems(7) = rsFilter1!CrAcctCode & "-" & rsFilter1!CrAcctName
    End If
    MItem.SubItems(8) = rsFilter1!receiptno
    rsFilter1.MoveNext
Loop
rsFilter1.close
 
 
If Trancode = "DQ" Or Trancode = "RQ" Or Trancode = "WQ" Or Trancode = "[A" And Me.Combo20 = "" Then
 Do Until rsFilter2.EOF = True
        On Error Resume Next
        Set MItem = Me.ListView3.ListItems.Add(, , Format(rsFilter2!receiptdate, "dd/mm/yyyy"))
        MItem.SubItems(1) = rsFilter2!receiptno
        If Trancode = "DQ" Then
           MItem.SubItems(2) = Trim(Right(rsFilter2!banknumber, 12)) 'Trim(Right(rsFilter2!custname, 12))
           MItem.SubItems(3) = Trim(Left(rsFilter2!banknumber, Trim(Len(rsFilter2!banknumber)) - 14))
           Me.ListView3.ColumnHeaders(4).Text = "BankName"
          Else
           Me.ListView3.ColumnHeaders(4).Text = "Bank AccountName"
           MItem.SubItems(2) = rsFilter2!banknumber
           MItem.SubItems(3) = rsFilter2!bankname
        End If
        MItem.SubItems(4) = Format(rsFilter2!checkreceipt, "###,###,###.#0")
        If Trancode = "DQ" Then
          MItem.SubItems(5) = "DQ-Cheque Deposit" 'rsBankTran!CQ
        ElseIf Trancode = "WQ" Then
          MItem.SubItems(5) = "WQ-Cheque Payment" 'rsBankTran!CQ
        ElseIf Trancode = "RQ" Then
          MItem.SubItems(5) = "RQ-Cheque Return" 'rsBankTran!CQ
        End If
        MItem.SubItems(6) = rsFilter2!moderef
        MItem.SubItems(7) = Format(rsFilter2!chkdue, "dd/mm/yyyy")
        MItem.SubItems(8) = rsFilter2!receiptno
        MItem.SubItems(9) = rsFilter2!custname
  rsFilter2.MoveNext
 Loop
 rsFilter2.close
End If
 
If Left(Me.Combo9, 2) = "TO" Then
   BankTransaction1.ListView3.ColumnHeaders(7).Text = "Payment No."
   BankTransaction1.ListView3.ColumnHeaders(8).Text = "Cr AcctCode"
  Else
 Me.ListView3.ColumnHeaders(7).Text = "Checque No."
 BankTransaction1.ListView3.ColumnHeaders(8).Text = "ClearingDate"
End If
If Left(Me.Combo9, 2) = "DC" Then
   BankTransaction1.ListView3.ColumnHeaders(7).Text = "Dep Slip No."
   BankTransaction1.ListView3.ColumnHeaders(8).Text = "Cr AcctCode"
End If
End Sub



Private Sub Command1_Click()
 PaymentList.Show
End Sub

Private Sub Command3_Click()
FindAcctNAme = True
FindAcctNames.Show 1
End Sub

Private Sub Combo7_Click()
xAccount = Me.Combo7
'open all accounts
Dim RstBA As ADODB.Recordset
Set RstBA = New ADODB.Recordset
xtable = "FinanceMaster"
xKey = "select * from " & xtable & " where " & _
         " AccountCode = " & "'" & xAccount & "'"
     
         
RstBA.Open xKey, constring, adOpenDynamic, adLockOptimistic, adCmdText
If RstBA.EOF = False Then
     Me.Combo13 = RstBA!accountnameeng & "\" & RTrim(RstBA!accountnamearab)
End If
RstBA.close
Dim catName As String
Dim Prevcap As String
Dim acctNo As String
acctNo = Trim(Me.Combo7)
Prevcap = Trim(Me.caption)
Call DisplayCats(Prevcap, acctNo, catName)
Me.caption = "Bank Transactions" & "//" & catName
CrCat = catName
End Sub

Private Sub Combo7_GotFocus()
Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(Combo7.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)

End Sub

Private Sub Combo7_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 34 Then
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(Combo7.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End If

End Sub

Private Sub Combo7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.Combo13.SetFocus
End If

End Sub

Private Sub Combo7_LostFocus()
'we don't allow the user to input same account number
If Me.Combo6 <> "" And Me.Combo7 <> "" Then
    If Me.Combo6 = Me.Combo7 Then
     xmsg = MsgBox("You have entered same Account number in Debit Side", vbInformation + vbOKOnly, "Message")
     Me.Combo7.SetFocus
     Exit Sub
    End If
End If

xdecimal = 0
If Me.Combo7 = "" Then
    Exit Sub
End If
xAccount = Me.Combo7
Dim RstBA As ADODB.Recordset
Set RstBA = New ADODB.Recordset
xtable = "FinanceMAster"
  xKey = "select * from " & xtable & " where " & _
         " AccountCode = " & "'" & xAccount & "'"
RstBA.Open xKey, constring, adOpenDynamic, adLockOptimistic, adCmdText
If RstBA.EOF = True Then
  xmsg = MsgBox("Account Code does not exist", vbInformation + vbOKOnly, "Message")
    Me.Combo7.SetFocus
    Exit Sub
Else
  Me.Combo13 = RstBA!accountnameeng & "\" & RstBA!accountnamearab
End If
If Me.Combo7 <> Me.Text5.Text Then
    Me.Text4 = "Text"
End If


End Sub

Private Sub Combo8_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.Combo23.SetFocus
End If
End Sub









Private Sub Command5_Click()
FindAcctNAme = True
FindAcctNames.Show 1
End Sub

Private Sub DElSelectitem_Click()
DeleteItems
End Sub

Private Sub eEdit_Click()
  If Me.ListView2.ListItems.Count Then
        Me.Text2 = "Edit"
        Dim rsttran As New ADODB.Recordset
        Dim Sn As String
        trandate = Me.ListView2.SelectedItem.SubItems(4)
        Me.Combo2 = trandate 'put the transdate selected
        Sn = Me.ListView2.SelectedItem.SubItems(8)
        Me.Combo11 = Sn
        'Check TEmptable if i is empty, if not empty we can't edit the a data
        rstTemp.Requery
        If rstTemp.RecordCount <> 0 Then
            mess = MsgBox("There was an unsaved entries that you must first save before editing.", vbInformation + vbOKOnly, "Message")
            Exit Sub
        End If
        rsttran.Open "BankJournal", constring, adOpenDynamic, adLockOptimistic, adCmdTable
        Me.ListView1.ColumnHeaders(2).Alignment = lvwColumnRight
    
        While rsttran.EOF = False
         If rsttran!TRansDate = trandate Then
          If rsttran!SerialNo = Sn Then
            On Error Resume Next
            Set MItem = Me.ListView1.ListItems.Add(, , rsttran!ticket, , 1)
            MItem.SubItems(1) = IIf(Len(rsttran!ticket) = 1, "  " _
                                & (rsttran!ticket), IIf(Len(rsttran!ticket) = 2, _
                                " " & rsttran!ticket, rsttran!ticket))
            MItem.SubItems(2) = rsttran!accountnumber
            MItem.SubItems(3) = rsttran!accountname
            MItem.SubItems(4) = Format(rsttran!TRansDate, "dd/mm/yyyy")
            MItem.SubItems(5) = rsttran!DebitAmount
            MItem.SubItems(6) = rsttran!creditamount
            MItem.SubItems(7) = rsttran!fxcyAmount
            MItem.SubItems(8) = rsttran!SerialNo
            MItem.SubItems(9) = rsttran!Description
            MItem.SubItems(10) = rsttran!Classification
            'transfer transaction to temptable for modification for deletion
            On Error GoTo 0
             With rstTemp
                .addnew
                !ticket = rsttran!ticket
                !accountnumber = rsttran!accountnumber
                !accountname = rsttran!accountname
                !TRansDate = rsttran!TRansDate
                !DebitAmount = rsttran!DebitAmount
                !creditamount = rsttran!creditamount
                !SerialNo = rsttran!SerialNo
                !Description = rsttran!Description
                !Classification = rsttran!Classification
                !Prepby = rsttran!Prepby
                !deletemark = 0
                !Status = "Unposted"
                .Update
              End With
            rsttran.Delete
           End If
          End If
    
         rsttran.MoveNext
        Wend
        
        NextTn = Me.ListView2.ListItems.Count + 1 'take the total Trans for TN reference
        ListView1.SortKey = 1
        ListView1.Sorted = True
       'get the selected item in the listview2
        Me.Text4.Text = Me.ListView2.SelectedItem
        'put selected item in corresponding comboboxes
        If Val(Me.ListView2.SelectedItem.SubItems(5)) <> 0 Then
            Me.Combo6 = Me.ListView2.SelectedItem.SubItems(2)
            Me.Combo12 = Me.ListView2.SelectedItem.SubItems(3)
            Me.MaskEdBox1.Text = Me.ListView2.SelectedItem.SubItems(5)
            Me.Combo11 = Me.ListView2.SelectedItem.SubItems(8)
            Me.Combo14 = Me.ListView2.SelectedItem.SubItems(1)
            Me.Text4.Text = Trim(Me.ListView2.SelectedItem.SubItems(1))
            Me.Combo23 = Me.ListView2.SelectedItem.SubItems(7)
          Else
            Me.Combo7 = Me.ListView2.SelectedItem.SubItems(2)
            Me.Combo13 = Me.ListView2.SelectedItem.SubItems(3)
            Me.MaskEdBox2.Text = Me.ListView2.SelectedItem.SubItems(6)
            Me.Combo23 = Me.ListView2.SelectedItem.SubItems(7)
            Me.Combo11 = Me.ListView2.SelectedItem.SubItems(8)
            Me.Combo14 = Me.ListView2.SelectedItem.SubItems(1)
            Me.Text4.Text = Trim(Me.ListView2.SelectedItem.SubItems(1))
        End If
         DisplayTransToday ' refesh the of listview2
    End If
 Me.SSTab1.SetFocus
 SendKeys "{Left}"

End Sub

Private Sub FI_Click()

End Sub

Private Sub Form_Activate()
'Open all transactions today
 DisplayTransToday
If Me.ListView2.ListItems.Count = 0 And Me.ListView3.ListItems.Count = 0 Then
    Me.Combo6.Enabled = False
    Me.Combo7.Enabled = False
    Me.Combo12.Enabled = False
    Me.Combo13.Enabled = False
    Me.MaskEdBox1.Enabled = False
    Me.MaskEdBox2.Enabled = False
    Me.Combo8.Enabled = False
   Else
    Me.Combo6.Enabled = True
    Me.Combo7.Enabled = True
    Me.Combo12.Enabled = True
    Me.Combo13.Enabled = True
    Me.MaskEdBox1.Enabled = True
    Me.MaskEdBox2.Enabled = True
    Me.Combo8.Enabled = True
End If
If Trim(Me.Combo9.Text) = "" Then
    Me.Combo9.SetFocus
    SendKeys "{Down}"
End If
End Sub

Private Sub Form_GotFocus()
Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(Combo1.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
   SendKeys "{PgUp}"

End Sub

Private Sub Form_Load()
 Me.Combo2.Text = Format(Date, "dd/mm/yyyy")
 Me.ListView1.ListItems.clear
 Dim xClass As New HabitatClass
 Set CON1 = New ADODB.Connection
 xtable = "Select * from FinanceMaster order by AccountNameEng"
 sqltable = True
 xClass.GetTables acctnames, CON1, xtable, constring, sqltable


'open all accounts
Set CON1 = New ADODB.Connection
xtable = "Select * from FinanceMaster order by AccountCode"
sqltable = True
xClass.GetTables AcctCode, CON1, xtable, constring, sqltable
AcctCode.MoveFirst
While AcctCode.EOF = False
 If AcctCode!Active = 1 Then
        If Left(AcctCode!AccountCode, 9) = "111020101" Then
          Me.Combo20.AddItem AcctCode!AccountCode
          Me.Combo19.AddItem AcctCode!accountnameeng
          Me.Combo18.AddItem AcctCode!accountnamearab
         End If
    'If Left(AcctCode!accountCode, 4) <> "2212" Then
    'If Left(AcctCode!accountCode, 2) <> "11" Then
        
  
       Me.Combo6.AddItem AcctCode!AccountCode
       Me.Combo12.AddItem AcctCode!accountnameeng & "\" & RTrim(AcctCode!accountnamearab)
     'Else
       Me.Combo7.AddItem AcctCode!AccountCode
       Me.Combo13.AddItem AcctCode!accountnameeng & "\" & RTrim(AcctCode!accountnamearab)
     End If
   'End If
  'End If
  AcctCode.MoveNext
  Wend
 
 
 'open bank transaction code
 Me.Combo9.AddItem "[All] "
 Dim rsBankTRanCode As New ADODB.Recordset
 rsBankTRanCode.Open "Select * from BankTrancode order by Code", constring, adOpenKeyset, adLockPessimistic, adCmdText
 While rsBankTRanCode.EOF = False
    Me.Combo17.AddItem rsBankTRanCode!Code & "-" & rsBankTRanCode!NameEng
    Me.Combo9.AddItem rsBankTRanCode!Code & "-" & rsBankTRanCode!NameEng
    rsBankTRanCode.MoveNext
 Wend
 rsBankTRanCode.close
 
 
 'Open Bank TRansaction
 Dim rsBankTRan As New ADODB.Recordset
 rsBankTRan.Open "Select * from BankTransaction where status is null", constring, adOpenKeyset, adLockPessimistic, adCmdText
 While rsBankTRan.EOF = False
    On Error Resume Next
    Set MItem = Me.ListView3.ListItems.Add(, , rsBankTRan!TRansDate)
    MItem.SubItems(1) = rsBankTRan!RefNo
    MItem.SubItems(2) = rsBankTRan!AccountCode
    MItem.SubItems(3) = rsBankTRan!NameEng
    MItem.SubItems(4) = Format(rsBankTRan!TransAmt, "###,###,###.#0")
    MItem.SubItems(5) = rsBankTRan!Trantype
    MItem.SubItems(6) = rsBankTRan!ChequeNo
    MItem.SubItems(7) = rsBankTRan!Description
    rsBankTRan.MoveNext
 Wend
 rsBankTRan.close
 On Error GoTo 0
 rsBankTRan.Open "Select * from vouchers where Left(Paymode,2)=03 and receiptdate > '01/10/2003'", constring, adOpenKeyset, adLockPessimistic, adCmdText
 While rsBankTRan.EOF = False
    Set MItem = Me.ListView3.ListItems.Add(, , rsBankTRan!receiptdate)
    On Error Resume Next
    MItem.SubItems(1) = rsBankTRan!receiptno
    MItem.SubItems(2) = rsBankTRan!banknumber
    MItem.SubItems(3) = rsBankTRan!bankname
    MItem.SubItems(4) = Format(rsBankTRan!checkreceipt, "###,###,###.#0")
    If IsNull(rsBankTRan!bankname) = True Then
      MItem.SubItems(5) = "DQ-Cheque Deposit" 'rsBankTran!CQ
     Else
     MItem.SubItems(5) = "WQ-Cheque Payment" 'rsBankTran!CQ
    End If
    MItem.SubItems(6) = rsBankTRan!moderef
    MItem.SubItems(7) = Format(rsBankTRan!chkdue, "dd/mm/yyyy")
    rsBankTRan.MoveNext
 Wend
 
 
 
 If Me.Text2.Text <> "Edit" Then
        Set RstBA = New ADODB.Recordset

        'Check Temp table if it is empty or not. if not put it on the list view
         On Error Resume Next
         rstTemp.close
         rstTemp.Open "TempBankJournal", CON1, adOpenKeyset, adLockOptimistic, adCmdTable
         On Error GoTo 0
         If rstTemp.EOF = False Then
            rstTemp.MoveFirst
            Me.Combo11 = rstTemp!SerialNo
          Else
               'open setuptable to get journal no.
               Dim JOurnalNo As New ADODB.Recordset
               JOurnalNo.Open "Setup", constring, adOpenKeyset, adLockPessimistic, adCmdTable
               If Val(Left(JOurnalNo!CurrentMoYr, 2)) <> (Format(Date, "mm")) Then
                   JOurnalNo!CurrentMoYr = Format(Date, "mm") & Trim(Right(Year(Date), 2))
                   JOurnalNo!nextjn = "00001"
                   JOurnalNo.Update
                   Me.Combo11 = JOurnalNo!nextjn
                   Jn = "BNK" & "-" & Right(Year(Date), 2) & Trim(Format(Date, "mm")) & "-" & Right(JOurnalNo!nextjn, 5)
                   JOurnalNo.close
                Else
                   Jn = "BNK" & "-" & Right(Year(Date), 2) & Trim(Format(Date, "mm")) & "-" & Right(JOurnalNo!nextjn, 5)
                   nextjn = Val(JOurnalNo!nextjn)
                   If (Len(nextjn)) = 1 Then
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
                   'JOurnalNo.Close
               End If
              Me.Combo11 = Jn

         End If
         On Error Resume Next
         While rstTemp.EOF = False
            Set MItem = Me.ListView1.ListItems.Add(, , rstTemp!ticket, , 1)
            MItem.SubItems(1) = rstTemp!ticket
            MItem.SubItems(2) = rstTemp!accountnumber
            MItem.SubItems(3) = rstTemp!accountname
            MItem.SubItems(4) = Format(rstTemp!TRansDate, "dd/mm/yyyy")
            MItem.SubItems(5) = Format(rstTemp!DebitAmount, "###,###,###.#0")
            MItem.SubItems(6) = Format(rstTemp!creditamount, "###,###,###.#0")
            MItem.SubItems(7) = rstTemp!SerialNo
            MItem.SubItems(8) = rstTemp!Description
            MItem.SubItems(9) = rstTemp!Classification
            rstTemp.MoveNext
         Wend
         NextTn = Me.ListView1.ListItems.Count + 1
         Me.Combo14 = NextTn
         On Error GoTo 0

End If
Nelson:
c = Err.Number
If c = 3705 Then
  rstTemp.close
  rstTemp.Open "TempBankJournal", CON1, adOpenKeyset, adLockOptimistic, adCmdTable
End If
End Sub
Sub Displaytrans()

Dim rst As ADODB.Recordset
Set rst = New ADODB.Recordset
Set CON1 = New ADODB.Connection
'conString = "Provider=MSDASQL;DSN=Ledger;UID=sa;pwd=;"
xKey = "SELECT * From Transactions  Where transdate =" & "'" & Date & "'" & "and Deletemark=" & "'" & 0 & "'" & "order by Ticket"
rst.Open xKey, constring, adOpenDynamic, adLockOptimistic, adCmdText
Me.ListView1.ColumnHeaders(2).Alignment = lvwColumnRight
Me.ListView1.ColumnHeaders(6).Alignment = lvwColumnRight
Me.ListView1.ColumnHeaders(7).Alignment = lvwColumnRight
Do Until rst.EOF
 If rst!TRansDate = Date Then
   cAmount = IIf(rst!DebitAmount > 0, Format(rst!DebitAmount, "###,###,###.#0"), Format(rst!creditamount, "###,###,###.#0"))
   Set MItem = Me.ListView1.ListItems.Add(, , "", IIf(rst!TRansDate = Date, 1, 1), IIf(rst!TRansDate = Date, 1, IIf(rst!TRansDate = Date - 1, 2, IIf(rst!TRansDate = Date - 2, 3, IIf(rst!TRansDate = Date - 3, 4, IIf(rst!TRansDate = Date - 4, 5, 5))))))
   MItem.SubItems(1) = rst!ticket
   MItem.SubItems(2) = rst!accountnumber
   MItem.SubItems(3) = rst!accountname
   MItem.SubItems(4) = Format(rst!TRansDate, "dd/mm/yyyy")
   On Error Resume Next
     If IsNull(rst!DebitAmount) = False Then
       MItem.SubItems(5) = Format(rst!DebitAmount, "###,###,###.#0")  '&  IIf(rst!Creditbalance = 0, Format(rst!Debitbalance, "###,###,###.#0"), Format(rst!Creditbalance, "###,###,###.#0"))
      Else
       MItem.SubItems(5) = "0.00"
   End If
   If IsNull(rst!creditamount) = False Then
     MItem.SubItems(6) = Format(rst!creditamount, "###,###,###.#0") '&  IIf(rst!Creditbalance = 0, Format(rst!Debitbalance, "###,###,###.#0"), Format(rst!Creditbalance, "###,###,###.#0"))
    Else
       MItem.SubItems(6) = "0.00"
   End If
   MItem.SubItems(7) = rst!SerialNo
   MItem.SubItems(8) = rst!Description
   MItem.SubItems(9) = rst!Classification
   MItem.SubItems(10) = rst!Prepby
   End If
rst.MoveNext
Loop
rst.close

End Sub

Private Sub Form_Unload(Cancel As Integer)
If Me.Text2.Text <> "Edit" Then

  If Me.ListView1.ListItems.Count <> 0 Then
    If Trim(Me.SSTab1.caption) = "List" Then
      Me.SSTab1.SetFocus
      SendKeys "{Right}"
     End If
        msg = MsgBox("Do you want to keep unsaved entries?", vbQuestion + vbYesNoCancel + vbDefaultButton1, "Please confirm")
        If msg = vbYes Then
            pressno = 0
            rstTemp.close
            Unload Me
         ElseIf msg = vbNo Then
          xmsg = MsgBox("All your entries below will be discarded if you Click Yes button  " & vbCrLf & _
                "Do you want to discard it?    ", vbQuestion + vbYesNo, "Please confirm")
          If xmsg = vbYes Then
            Dim DelTran As New ADODB.Recordset
            DelTran.Open "Delete TempBankJournal", constring, adOpenDynamic, adLockPessimistic, adcmdttext
            rstTemp.close
            Unload Me
           Else
             Cancel = -1
          End If
         Else
           Cancel = -1
        End If
   End If
   
 Else
    'if this transaction is already saved, this will be
    'automatically remove from temptable then placed it
    'to transaction table.
    On Error GoTo 0
    Dim rsttran As New ADODB.Recordset
    rsttran.Open "BankJournal", constring, adOpenDynamic, adLockOptimistic, adCmdTable
    rstTemp.close
    rstTemp.Open "TempBankJournal", constring, adOpenDynamic, adLockOptimistic, adCmdTable
    
    While rstTemp.EOF = False
         With rsttran
            .addnew
            !ticket = rstTemp!ticket
            !accountnumber = rstTemp!accountnumber
            !accountname = rstTemp!accountname
            !TRansDate = rstTemp!TRansDate
            !DebitAmount = rstTemp!DebitAmount
            !creditamount = rstTemp!creditamount
            !fxcyAmount = rstTemp!fxcyAmount
            !SerialNo = rstTemp!SerialNo
            !Description = rstTemp!Description
            '!deletemark = 0
            !Status = "Unposted"
            .Update
          End With
        rstTemp.MoveNext
    Wend
    rstTemp.close
    c = Err.Description
    rstTemp.Open "Delete TempBankJournal", constring, adOpenDynamic, adLockPessimistic, adcmdttext
   
End If
End Sub

Private Sub Fv_Click()
If Me.FV.caption = "&Normal View" Then
  Me.FV.caption = "&Full View"
   BankTransaction1.Frame5.Top = 4920
   BankTransaction1.ListView1.Top = 5160
   BankTransaction1.ListView1.Height = 1935
   BankTransaction1.Frame5.Height = 2295
  Else
   Me.FV.caption = "&Normal View"
   BankTransaction1.Frame5.Top = 350
   BankTransaction1.ListView1.Top = 630
   BankTransaction1.Frame5.Height = 7600
   BankTransaction1.ListView1.Height = 7200
End If

End Sub

Private Sub Li_Click()
Me.ListView2.View = lvwIcon
Me.ListView3.View = lvwIcon
Me.Li.Checked = True
Me.SI.Checked = False
Me.xlist.Checked = False
Me.xDetails.Checked = False
End Sub

Private Sub ListView1_DblClick()
'If Me.ListView1.Top = 240 Then
'   'Scmenu.Fv.caption = "&Full View"
'   BankTransaction1.Frame5.Top = 4940
'   BankTransaction1.ListView1.Top = 5160
'   BankTransaction1.ListView1.Height = 1935
'   BankTransaction1.Frame5.Height = 2295
'  Else
'   'Scmenu.Fv.caption = "&Normal View"
'   BankTransaction1.Frame5.Top = 370
'   BankTransaction1.ListView1.Top = 630
'  BankTransaction1.Frame5.Height = 7600
'   BankTransaction1.ListView1.Height = 7200
' End If

End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error Resume Next
If Val(Me.ListView1.SelectedItem.SubItems(5)) <> 0 Then
    Me.Text4.Text = Trim(Me.ListView1.SelectedItem.SubItems(1))
    Me.Combo14.Text = Me.ListView1.SelectedItem.SubItems(1)
    Me.Combo6 = Me.ListView1.SelectedItem.SubItems(2)
    Me.Combo12 = Me.ListView1.SelectedItem.SubItems(3)
    Me.MaskEdBox1 = Me.ListView1.SelectedItem.SubItems(5)
    Me.Combo8 = Me.ListView1.SelectedItem.SubItems(8)
    DrCat = Me.ListView1.SelectedItem.SubItems(9)
    Me.Combo7 = ""
    Me.Combo13 = ""
    Me.MaskEdBox2 = ""
End If

If Val(Me.ListView1.SelectedItem.SubItems(6)) <> 0 Then
    Me.Text4.Text = Trim(Me.ListView1.SelectedItem.SubItems(1))
    Me.Combo14.Text = Me.ListView1.SelectedItem.SubItems(1)
    Me.Combo7 = Me.ListView1.SelectedItem.SubItems(2)
    Me.Combo13 = Me.ListView1.SelectedItem.SubItems(3)
    Me.MaskEdBox2 = Me.ListView1.SelectedItem.SubItems(6)
    Me.Combo8 = Me.ListView1.SelectedItem.SubItems(8)
    CrCat = Me.ListView1.SelectedItem.SubItems(9)
    Me.Combo6 = ""
    Me.Combo12 = ""
    Me.MaskEdBox1 = ""
    
End If
Me.Text5.Text = Me.ListView1.SelectedItem.SubItems(2)
End Sub

Private Sub ListView1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
    DeleteItems
End If
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Me.ListView1.ListItems.Count <> 0 Then
If Button = 2 Then
   PopupMenu Main2
End If
End If
End Sub

Private Sub ListView2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Me.ListView2.SortKey = ColumnHeader.Index - 1
Me.ListView2.Sorted = True
End Sub

Private Sub ListView2_DblClick()
'reset ghosted items in listview3
i = 0
For i = 1 To Me.ListView3.ListItems.Count
    Me.ListView3.ListItems.Item(i).ForeColor = vbBlack
Next


'Check the temptable if it is empty. if not abort.
If Me.ListView2.ListItems.Count <> 0 Then
 Dim rsttran As ADODB.Recordset
 Set rsttran = New ADODB.Recordset
 rsttran.Open "TempBankJournal ", constring, adOpenDynamic, adLockOptimistic, adCmdTable
 If rsttran.EOF = False Then
    xmsg = MsgBox("There was unsaved transactions that you must save first before Deleting or Modifying any saved transactions.", vbInformation + vbOKOnly, "Message")
     Exit Sub
 End If
 rsttran.close
 Me.Text2.Text = "Edit"
 Call eEdit_Click
 'Me.SSTab1.SetFocus
 'SendKeys "{Left}"

End If

End Sub

Private Sub ListView2_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = 2 Then
  If Trim(Left(Me.SSTab1.caption, 12)) = "Journal List" Then
    If Me.ListView2.ListItems.Count = 0 Then
        Me.xFind.Enabled = False
        Me.eEdit.Enabled = False
        'Me.xrefresh.Enabled = False
        Me.xPreview.Enabled = False
        Me.POstToGL.Enabled = False
        Me.xCancel.Enabled = False
       Else
        Me.xFind.Enabled = True
        Me.eEdit.Enabled = True
        Me.xrefresh.Enabled = True
        Me.xPreview.Enabled = True
        Me.xCancel.Enabled = True
        
    End If
    PopupMenu Me.main
   End If
End If
End Sub

Private Sub ListView3_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Me.ListView3.SortKey = ColumnHeader.Index - 1
Me.ListView3.Sorted = True

End Sub

Private Sub ListView3_DblClick()
'If Me.ListView1.ListItems.Count <> 0 Then
'    mess = MsgBox("You have an unsaved entries Please save it first.", vbInformation + vbOKOnly, "Message")
'    Me.SSTab1.SetFocus
'    SendKeys "{Right}"
'    Exit Sub
'End If
'Dim cAmt As Currency
'If Me.ListView3.ListItems.Count = 0 Then
'    Exit Sub
'End If
'InvCat = Trim(Me.ListView3.SelectedItem.SubItems(2))
'DOno = Trim(Me.ListView3.SelectedItem.SubItems(4))
'TRanType = Trim(Me.ListView3.SelectedItem.SubItems(10))
'amount = Trim(Me.ListView3.SelectedItem.SubItems(5))
''item inventory
'GN = Trim(Me.ListView3.SelectedItem.SubItems(1))
'IC = Trim(Me.ListView3.SelectedItem.SubItems(2))
'VN = Trim(Me.ListView3.SelectedItem.SubItems(3))
'DOno = Trim(Me.ListView3.SelectedItem.SubItems(4))
'FrCC = Trim(Left(Me.ListView3.SelectedItem.SubItems(6), 3))
'FrDept = Trim(Right(Me.ListView3.SelectedItem.SubItems(6), 3))
'ToCC = Trim(Left(Me.ListView3.SelectedItem.SubItems(7), 3))
'ToDept = Trim(Right(Me.ListView3.SelectedItem.SubItems(7), 3))
'Purpose = Trim(Me.ListView3.SelectedItem.SubItems(8))
'WOno = Trim(Me.ListView3.SelectedItem.SubItems(9))
'TrnType = Trim(Me.ListView3.SelectedItem.SubItems(10))
'
'If TRanType = "A" Then
'    cTranType = "A"
'ElseIf TRanType = "B" Then
'cTranType = "B"
'ElseIf TRanType = "C" Then
'    cTranType = "C"
'End If
'If TRanType = "D" Then
'    cTranType = "D"
'End If
'
'i = 0
'For i = 1 To Me.ListView3.ListItems.Count
'    If Trim(Me.ListView3.ListItems.Item(i).SubItems(4)) = DOno And Trim(Me.ListView3.ListItems.Item(i).SubItems(2)) = InvCat Then 'for DO
'        Me.ListView3.ListItems.Item(i).ForeColor = &H80000004
'        amount = Trim(Me.ListView3.ListItems.Item(i).SubItems(5))
'        cAmt = cAmt + Format(amount, "###,###,###.#0")
'      Else
'        Me.ListView3.ListItems.Item(i).ForeColor = vbBlack
'    End If
'Next
'If cTranType <> "A" Then
'   Me.MaskEdBox2.Text = FormatNumber(cAmt, 2, vbTrue, vbTrue, vbTrue)
'   Me.MaskEdBox1.Text = FormatNumber(cAmt, 2, vbTrue, vbTrue, vbTrue)
'   Me.MaskEdBox2.Enabled = False
'   Me.MaskEdBox1.Enabled = True
'  Else
'  Me.MaskEdBox1.Text = FormatNumber(cAmt, 2, vbTrue, vbTrue, vbTrue)
'  Me.MaskEdBox2.Text = FormatNumber(cAmt, 2, vbTrue, vbTrue, vbTrue)
'  Me.MaskEdBox1.Enabled = False
'  Me.MaskEdBox2.Enabled = True
'End If
'Me.SSTab1.SetFocus
'SendKeys "{Right}"

End Sub

Private Sub ListView3_ItemCheck(ByVal Item As MSComctlLib.ListItem)
If Trim(Me.Combo9) = "[All]" Then
    mess = MsgBox("Please filter the items before selecting it", vbExclamation + vbOKOnly, "Message")
    i = 0
    For i = 1 To Me.ListView3.ListItems.Count
         Me.ListView3.ListItems.Item(i).Checked = False
    Next
    Me.Combo9.SetFocus
    Exit Sub
End If
    
Dim cAmt As Currency
Dim MaskedAmt As Currency
On Error Resume Next
MaskedAmt = Me.MaskEdBox1.Text
'CheckOwnedby = ""
'CheckNo = ""
'Payee = ""
'Checktype = ""
ORno = Trim(Item.SubItems(1))
Trantype = "Posted Balance"
Codetype = Left(Item.SubItems(5), 2)
On Error GoTo 0
cAmt = Item.SubItems(4)
If Item.Checked = True Then
    
    Totamt = MaskedAmt + cAmt
    Me.MaskEdBox1.Text = Format(Totamt, "###,###,###.#0")
    Me.MaskEdBox2.Text = Format(Totamt, "###,###,###.#0")
    If Left(Item.SubItems(5), 2) = "DC" Then
        Dim rsCrAcct As New ADODB.Recordset
        rsCrAcct.Open "Select * from financemaster where accountCode='111020105001'", constring, adOpenKeyset, adLockPessimistic, adCmdText
        Me.Combo6 = Item.SubItems(2)
        Me.Combo12 = (Item.SubItems(3))
        If Left(Me.Combo9, 2) = "DC" Then
         If Trim(Item.SubItems(7)) <> "" Then
          Me.Combo7 = rsCrAcct!AccountCode
          Me.Combo13 = rsCrAcct!accountnameeng
         Else
          Me.Combo7 = ""
          Me.Combo13 = ""
         End If
        End If
        Me.Combo8 = "Dep Slip# " & Trim(Item.SubItems(6)) & " and  O.R.# " & Trim(Item.SubItems(1))
        rsCrAcct.close
        'Trantype = Trim(Mid(Item.SubItems(5), 4, 20))
        ORno = "Slip# " & Trim(Item.SubItems(6))
        CheckOwnedby = Item.SubItems(2)
       'for check pmt
      ElseIf Left(Item.SubItems(5), 2) = "WQ" Then
        rsCrAcct.Open "Select * from financemaster where accountCode='131020501001'", constring, adOpenKeyset, adLockPessimistic, adCmdText
        Me.Combo6 = rsCrAcct!AccountCode
        Me.Combo12 = rsCrAcct!accountnameeng
        Me.Combo7 = Item.SubItems(2)
        Me.Combo13 = (Item.SubItems(3))
        Me.Combo8 = "Check# " & Trim(Item.SubItems(6)) & " and  O.R.# " & Trim(Item.SubItems(1))
        CheckOwnedby = Item.SubItems(2)
        CheckNo = Trim(Item.SubItems(6))
        Payee = Trim(Item.SubItems(9))
        Checktype = "WQ"
        Trantype = "Reconciliation-Outstanding Checks" 'Trim(Mid(Item.SubItems(5), 4, 20))
        ORno = "O.R.# " & Trim(Item.SubItems(1))
      
      'for check deposit
      ElseIf Left(Item.SubItems(5), 2) = "DQ" Then
        rsCrAcct.Open "Select * from financemaster where accountCode='111020105001'", constring, adOpenKeyset, adLockPessimistic, adCmdText
        Me.Combo6 = Item.SubItems(2)
        Me.Combo12 = (Item.SubItems(3))
        Me.Combo7 = rsCrAcct!AccountCode
        Me.Combo13 = rsCrAcct!accountnameeng
        Me.Combo8 = "Check# " & Trim(Item.SubItems(6)) & " and  O.R.# " & Trim(Item.SubItems(1))
        CheckOwnedby = Item.SubItems(2)
        CheckNo = Trim(Item.SubItems(6))
        Payee = Trim(Item.SubItems(9))
        Checktype = "DQ"
        Trantype = "Reconciliation-Deposit-In-Transit" 'Trim(Mid(Item.SubItems(5), 4, 20))
        ORno = "O.R.# " & Trim(Item.SubItems(1))
        
      ElseIf Left(Item.SubItems(5), 1) = "T" Then
        Me.Combo6 = Left(Trim(Me.ListView3.SelectedItem.SubItems(7)), 12)
        Me.Combo12 = Trim(Right(Me.ListView3.SelectedItem.SubItems(7), Len(Trim(Me.ListView3.SelectedItem.SubItems(7))) - 13))
        Me.Combo7 = Item.SubItems(2)
        Me.Combo13 = (Item.SubItems(3))
        Trantype = Trim(Mid(Item.SubItems(5), 4, 20))
        ORno = "Pmt# " & Trim(Item.SubItems(1))
        CheckOwnedby = Item.SubItems(2)
    End If
    Item.ForeColor = &H8000000F
  Else
    Me.Combo6 = ""
    Me.Combo7 = ""
    Me.Combo12 = ""
    Me.Combo13 = ""
    Totamt = MaskedAmt - cAmt
    Me.MaskEdBox1.Text = Format(Totamt, "###,###,###.#0")
    Me.MaskEdBox2.Text = Format(Totamt, "###,###,###.#0")
    Item.ForeColor = vbBlack
End If


Dim catName As String
Dim Prevcap As String
Dim acctNo As String
acctNo = Trim(Me.Combo6)
Prevcap = Trim(Me.caption)
Call DisplayCats(Prevcap, acctNo, catName)
DrCat = catName
acctNo = Trim(Me.Combo7)
Prevcap = Trim(Me.caption)
Call DisplayCats(Prevcap, acctNo, catName)
CrCat = catName

End Sub

Private Sub ListView3_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Me.ListView3.ListItems.Count <> 0 Then
   If Mid(Me.ListView3.SelectedItem.SubItems(5), 2, 1) = "Q" Then
     Me.XdeleteMain3.Enabled = False
    Else
    Me.XdeleteMain3.Enabled = True
   End If
    If Trim(Me.Combo9) = "[All]" Then
        Me.xSelectall.Enabled = False
      Else
       Me.xSelectall.Enabled = True
     End If
    If Button = 2 Then
        PopupMenu main3
    End If
End If
End Sub

Private Sub ListView3_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Me.ListView3.ToolTipText = Me.ListView3.ListItems.Count & " total item(s)"
End Sub

Private Sub MaskEdBox1_GotFocus()
Call Combo6_LostFocus
Dim havedot As Boolean
  For i = 1 To Len(Trim(Me.MaskEdBox1.Text))
        Havedott = Mid(Me.MaskEdBox1.Text, i, 1)
        If Havedott = "." Then Exit For
           
  Next
  If Havedott = "." Then
     xdecimal = 1
     Exit Sub
    Else
      xdecimal = 0
  End If

i = 0
For i = 1 To Len(Trim(Me.MaskEdBox1.Text))
      X = Mid(Trim(Me.MaskEdBox1.Text), i, 1)
      If X = "." Then
        havedot = True
        Exit For
       Else
        havedot = False
      End If
Next i


xDot = Right(Trim(Me.MaskEdBox1.Text), 3)

If havedot = True And Left(xDot, 1) = "." Then
   xdecimal = 1
  Else
   xdecimal = 0
End If

End Sub

Private Sub MaskEdBox1_KeyDown(KeyCode As Integer, Shift As Integer)
 i = 0
If KeyCode = vbKeyDelete Then
  For i = 1 To Len(Trim(Me.MaskEdBox1.Text))
        Havedott = Mid(Me.MaskEdBox1.Text, i, 1)
        If Havedott = "." Then Exit For
  Next
  If Havedott = "." Then
     xdecimal = 1
     Exit Sub
    Else
      xdecimal = 0
  End If
  
  
 End If
 
  
If KeyCode = vbKeyDown Then
  If Me.Combo7.Enabled = True Then
    Me.Combo7.SetFocus
   Else
    Me.Combo8.SetFocus
  End If
End If
End Sub

Private Sub MaskEdBox1_KeyPress(KeyAscii As Integer)
'to avoid entering more than one decimal
If KeyAscii = 46 Then
   xdecimal = xdecimal + 1
 End If
If xdecimal = 2 Then
  Beep
  xdecimal = 1
 Me.MaskEdBox1.SetFocus
 SendKeys "{Left}+{End}"
 SendKeys "{Delete}"
End If

If KeyAscii = 13 Or KeyAscii = 46 Or KeyAscii >= 48 And KeyAscii <= 57 Then
 c = KeyAscii
 Else
   If KeyAscii = 46 And xdecimal > 1 Then
       xdecimal = 0
   End If
 If Me.MaskEdBox1.Text <> " " Then
  xdecimal = 0
  SendKeys "{End}+{Home}"
  SendKeys "{Delete}"
  Me.MaskEdBox1.SetFocus

 End If
End If
 

If KeyAscii = 13 Then
 If Me.Combo7.Enabled = True Then
  
  
  
  Me.Combo7.SetFocus
  Else
   'Me.Combo10.SetFocus
  End If
End If

End Sub

Private Sub MaskEdBox1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
  Me.Combo6.SetFocus
End If
End Sub

Private Sub MaskEdBox1_LostFocus()
xdecimal = 0
Me.MaskEdBox1.Text = Format(Me.MaskEdBox1.Text, "###,###,###,###.#0")

End Sub

Private Sub MaskEdBox2_GotFocus()
Call Combo7_LostFocus
Dim havedot As Boolean
  For i = 1 To Len(Trim(Me.MaskEdBox2.Text))
        Havedott = Mid(Me.MaskEdBox2.Text, i, 1)
        If Havedott = "." Then Exit For
           
  Next
  If Havedott = "." Then
     xdecimal = 1
     Exit Sub
    Else
      xdecimal = 0
  End If

i = 0
For i = 1 To Len(Trim(Me.MaskEdBox2.Text))
      X = Mid(Trim(Me.MaskEdBox2.Text), i, 1)
      If X = "." Then
        havedot = True
        Exit For
       Else
        havedot = False
      End If
Next i


xDot = Right(Trim(Me.MaskEdBox2.Text), 3)

If havedot = True And Left(xDot, 1) = "." Then
   xdecimal = 1
  Else
   xdecimal = 0
End If
End Sub

Private Sub MaskEdBox2_KeyDown(KeyCode As Integer, Shift As Integer)
i = 0
If KeyCode = vbKeyDelete Then
  For i = 1 To Len(Trim(Me.MaskEdBox2.Text))
        Havedott = Mid(Me.MaskEdBox2.Text, i, 1)
        If Havedott = "." Then Exit For
  Next
  If Havedott = "." Then
     xdecimal = 1
     Exit Sub
    Else
      xdecimal = 0
  End If
  
  
 End If
End Sub

Private Sub MaskEdBox2_KeyPress(KeyAscii As Integer)
'to avoid not to enter more than one decimal
If KeyAscii = 46 Then
   xdecimal = xdecimal + 1
 End If
If xdecimal = 2 Then
  Beep
  xdecimal = 1
 Me.MaskEdBox2.SetFocus
 SendKeys "{Left}+{End}"
 SendKeys "{Delete}"
End If

If KeyAscii = 13 Or KeyAscii = 46 Or KeyAscii >= 48 And KeyAscii <= 57 Then
 c = KeyAscii
 Else
   If KeyAscii = 46 And xdecimal > 1 Then
       xdecimal = 0
   End If
 If Me.MaskEdBox2.Text <> " " Then
  xdecimal = 0
  SendKeys "{End}+{Home}"
  SendKeys "{Delete}"
  Me.MaskEdBox2.SetFocus

 End If
End If
If KeyAscii = 13 Then
    Me.Combo8.SetFocus
End If

End Sub

Private Sub MaskEdBox2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
  Me.Combo7.SetFocus
End If

End Sub

Private Sub MaskEdBox2_LostFocus()
xdecimal = 0
Me.MaskEdBox2.Text = Format(Me.MaskEdBox2.Text, "###,###,###,###.#0")


End Sub

Private Sub MaskEdBox3_GotFocus()
Me.MaskEdBox3.Text = Format(Me.MaskEdBox3.Text, "dd/mm/yyyy")
End Sub

Private Sub MaskEdBox3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
       
        If Me.MaskEdBox3.Text = "__/__/____" Then
            Exit Sub
        End If
        Me.MaskEdBox3.Text = Format(Me.MaskEdBox3.Text, "dd/mm/yyyy")
        cDay = Val(Left(Me.MaskEdBox3.Text, 2))
        cMonth = Val(Mid(Me.MaskEdBox3.Text, 4, 2))
        cYear = Val(Right(Me.MaskEdBox3.Text, 4))
        If cDay > 31 Or cDay < 1 Then
            mess = MsgBox("Invalid Date", vbInformation + vbOKOnly, "Message")
            Me.MaskEdBox3.SetFocus
            
            Exit Sub
          ElseIf cMonth > 12 Or cMonth < 1 Then
            mess = MsgBox("Invalid Month", vbInformation + vbOKOnly, "Message")
            Me.MaskEdBox3.SetFocus
            Exit Sub
        ElseIf cYear < 1900 Or cYear > Year(Date) Then
            mess = MsgBox("Invalid Year", vbInformation + vbOKOnly, "Message")
            Me.MaskEdBox3.SetFocus
            Exit Sub
        End If
      Me.MaskEdBox3.Text = Format(Me.MaskEdBox3.Text, "dd/mm/yyyy")
      BankTRansdate = Format(Me.MaskEdBox3.Text, "dd/mm/yyyy")
      msg = MsgBox("Save entries? ", vbQuestion + vbOKCancel, "Please confirm")
      If Me.Combo15 = "" Or Me.MaskEdBox3.Text = "__/__/____" Then
         mess = MsgBox("Please enter filled up all the fields completely", vbExclamation + vbOKOnly, "Message")
         Exit Sub
       End If
     If msg = vbOK Then
        Call SaveTRan
    End If
End If
End Sub

Private Sub MaskEdBox3_LostFocus()
Me.MaskEdBox3.Text = Format(Me.MaskEdBox3.Text, "dd/mm/yyyy")
End Sub

Private Sub PostToGl_Click()
Dim rsInvj As New ADODB.Recordset
Dim MItem As ListItem
mess = MsgBox("Do you want to continue?", vbQuestion + vbOKCancel + vbDefaultButton2, "Please confirm")
If mess = vbOK Then
     PostingJournal.Text1.Text = "BNK"
     rsInvj.Open "SELECT  TransDate, COUNT(TransDate) AS TotalTRan, SUM(DebitAmount) AS DrAmt, SUM(CreditAmount) AS CrAmt" _
     & " From BankJournal GROUP BY TransDate ORDER BY TransDate", constring, adOpenKeyset, adLockPessimistic, adCmdText
    Do Until rsInvj.EOF = True
        Set MItem = PostingJournal.ListView1.ListItems.Add(, , rsInvj!TRansDate)
            MItem.SubItems(1) = rsInvj!TotalTRan
            MItem.SubItems(2) = FormatNumber(rsInvj!Dramt, 2, vbTrue, vbTrue, vbTrue)
            MItem.SubItems(3) = FormatNumber(rsInvj!CrAmt, 2, vbTrue, vbTrue, vbTrue)
            MItem.SubItems(4) = "Waiting"
            rsInvj.MoveNext
    Loop
    PostingJournal.Show 1
        
End If


End Sub

Private Sub PrintItem_Click()
PrintEntries
End Sub

Private Sub SI_Click()
Me.ListView2.View = lvwSmallIcon
Me.ListView3.View = lvwSmallIcon
Me.Li.Checked = False
Me.SI.Checked = True
Me.xlist.Checked = False
Me.xDetails.Checked = False

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
TAbClic = PreviousTab
End Sub

Private Sub Timer1_Timer()
'Me.Combo14 = Me.ListView1.ListItems.Count + 1
If Me.ListView1.ListItems.Count = 0 Then
   Me.Toolbar1.Buttons(2).Enabled = False
   Me.Toolbar1.Buttons(3).Enabled = False
   Me.Toolbar1.Buttons(5).Enabled = False
  Else
   Me.Toolbar1.Buttons(2).Enabled = True
   Me.Toolbar1.Buttons(3).Enabled = True
   Me.Toolbar1.Buttons(5).Enabled = True
 End If
 
 'rstPn.Open "SElect count(ticket) as [tn] from transactions where DatePart(d, transdate) = DatePart(d, GETDATE()) And DatePart(m, transdate) = DatePart(m, GETDATE()) And DatePart(yy, transdate) = DatePart(yy, GETDATE())" & " and Deletemark=" & "'" & 0 & "'", conString, adOpenDynamic, adLockOptimistic, adCmdText
 Dim rstPn As ADODB.Recordset
 Set rstPn = New ADODB.Recordset
 Dim TN As Long
 Dim TN1 As Long
 Dim tRANDr As Currency
 Dim tranCr As Currency
 Dim TempDr As Currency
 Dim TempCr As Currency
 Xdate = Format(Date, "mm/dd/yyyy")
 rstPn.Open "SElect count(ticket) as [tn1] from BankJournal where transdate=" & "'" & Xdate & "'", constring, adOpenDynamic, adLockOptimistic, adCmdText
 TN = rstPn!TN1  '+ 1
 Me.Combo1.Text = TN ' IIf(tn = 0, 1, tn)
 rstPn.close
  
 'put total credit in combo4
 rstPn.Open "Select SUm(CreditAmount)  as [TotalCr] from BankJournal where TransDate= " & "'" & Xdate & "'", constring, adOpenDynamic, adLockOptimistic, adCmdText
 Me.Combo5 = IIf(rstPn!TotalCr <> 0, Format(rstPn!TotalCr, "###,###,###.#0"), "0.00")
 Me.Combo5 = Format(Me.Combo5, "###,###,###.#0")
 rstPn.close
 
'put totals debit in combo5
 rstPn.Open "Select SUm(debitAmount)  as [TotalDb] from BankJournal where TransDate= " & "'" & Xdate & "'", constring, adOpenDynamic, adLockOptimistic, adCmdText
 Me.Combo4 = IIf(rstPn!TotalDb <> 0, Format(rstPn!TotalDb, "###,###,###.#0"), "0.00")
 Me.Combo4 = Format(Me.Combo4, "###,###,###.#0")
 rstPn.close
 
'tsik the maximum trans then compared it the number of transaction in the listview.
 Dim MaxTranPerJNL As Long
 rstPn.Open "Setup", constring, adOpenKeyset, adLockPessimistic, adCmdTable
 MaxTranPerJNL = rstPn!MaxTranPerJNL
 If Me.ListView1.ListItems.Count >= MaxTranPerJNL Then
    Me.Toolbar1.Buttons(1).Enabled = False
    rstPn.close
   Else
    'Me.Toolbar1.Buttons(1).Enabled = True
 End If

 
 
End Sub

Private Sub Timer2_Timer()
If Me.Combo20 = "" Or Me.Combo19 = "" Or Me.Combo18 = "" Or Me.Combo17 = "" Or Me.Combo16 = "" Then
    Me.Toolbar2.Buttons(1).Enabled = False
   Else
   Me.Toolbar2.Buttons(1).Enabled = True
End If
If Me.ListView3.ListItems.Count = 0 Then
   Me.Toolbar2.Buttons(2).Enabled = False
   If Me.ListView3.ListItems.Count <> 0 Then
   If Mid(Me.ListView3.SelectedItem.SubItems(5), 2, 1) = "Q" Then
     Me.Toolbar2.Buttons(4).Enabled = True
    Else
    Me.Toolbar2.Buttons(4).Enabled = False
   End If
   End If
  Else
   Me.Toolbar2.Buttons(2).Enabled = True
   If Me.ListView3.ListItems.Count <> 0 Then
    If Mid(Me.ListView3.SelectedItem.SubItems(5), 2, 1) = "Q" Then
      Me.Toolbar2.Buttons(4).Enabled = False
     Else
     Me.Toolbar2.Buttons(4).Enabled = True
    End If
   End If
 End If
End Sub

Private Sub Timer3_Timer()
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim xtemp As ADODB.Recordset
Set xtemp = New ADODB.Recordset
Select Case Button.Key
    Case Is = "Add"
        Call Combo6_LostFocus
        Call Combo7_LostFocus
        Dim strFindMe As String
        Dim itmFound As ListItem   ' FoundItem variable.
        intSelectedOption = lvwText
        
        strFindMe = Me.Text4.Text 'Trim(Combo14)
        Set itmFound = Me.ListView1.Finditem(strFindMe, intSelectedOption, , lvwPartial)
        If itmFound Is Nothing Then  ' If no match, inform user and exit.
            AddEntries
        Else
            cindex = Val(Me.ListView1.SelectedItem.Index)
            Me.ListView1.ListItems.Remove cindex
            'remove also in temp table
            xtemp.Open "delete TempBankJournal  where ticket=" & "'" & strFindMe & "'", constring, adOpenKeyset, adLockOptimistic, adCmdText
            AddEntries
        End If
        ListView1.SortKey = 1
        ListView1.Sorted = True
        Me.Text4.Text = "Text"
        Exit Sub
        
    '---------------------------------------------------------------
    Case Is = "Delete"
       DeleteItems
     '------------------------------------------------------------
     Case Is = "Save"
          X = 0
          Dim ccTotCr As Currency
          Dim ccTotDr As Currency
          Dim xCr As Currency
          Dim xDr As Currency
          
          'we don't allowed the user to save their entries
          'if the debit is not equal to the credit side
          For X = 1 To Me.ListView1.ListItems.Count
                  xDr = Me.ListView1.ListItems.Item(X).SubItems(5)
                  xCr = Me.ListView1.ListItems.Item(X).SubItems(6)
                  ccTotDr = ccTotDr + xDr
                  ccTotCr = ccTotCr + xCr
          Next
          If ccTotCr <> ccTotDr Then
            xmsg = MsgBox("Total Debit and Total Credit is not equal, Please equalized first before saving", vbInformation + vbOKOnly, "Message")
            Exit Sub
          End If
          xmsg = MsgBox(" Are you sure you want to save now?  ", vbQuestion + vbOKCancel, "Please Confirm")
          If xmsg = vbOK Then
              Dim RstBA As ADODB.Recordset
              Set RstBA = New ADODB.Recordset
                
             'increment Journal Entry Number
             Dim JOurnalNo As New ADODB.Recordset
             JOurnalNo.Open "setup", constring, adOpenKeyset, adLockOptimistic, adCmdTable
             Jn = "BNK" & "-" & Right(Year(Date), 2) & Trim(Format(Date, "mm")) & "-" & JOurnalNo!nextjn
             Me.Combo11 = Jn
             If Trim(Me.Text2.Text) = "Edit" Then
              Else
               If Val(Left(JOurnalNo!CurrentMoYr, 2)) <> Format(Date, "mm") Then
                   JOurnalNo!CurrentMoYr = Format(Date, "mm") & Trim(Right(Year(Date), 2))
                   JOurnalNo!nextjn = "00001"
                   JOurnalNo.Update
                   Me.Combo11 = JOurnalNo!nextjn
                   JOurnalNo.close
                Else
                   
                   nextjn = Val(JOurnalNo!nextjn) + 1
                   If (Len(nextjn)) = 1 Then
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
                   JOurnalNo!nextjn = Zeros & Trim(Val(nextjn))
                   JOurnalNo.Update
                   Jn = "BNK" & "-" & Right(Year(Date), 2) & Trim(Format(Date, "mm")) & "-" & JOurnalNo!nextjn
                   JOurnalNo.close
                End If
             End If
             Me.Combo11 = Jn
             Me.Combo14 = 1
             
             
             'save entries only to INVJOuranlTrans table
              RstBA.Open "BankJournal", constring, adOpenDynamic, adLockOptimistic, adCmdTable
              i = 0
              Dim cR As Currency
              Dim dr As Currency
              For i = 1 To Me.ListView1.ListItems.Count
                  TN = Me.ListView1.ListItems.Item(i).SubItems(1)
                  Ac = Me.ListView1.ListItems.Item(i).SubItems(2)
                  an = Me.ListView1.ListItems.Item(i).SubItems(3)
                  dr = Me.ListView1.ListItems.Item(i).SubItems(5)
                  cR = Me.ListView1.ListItems.Item(i).SubItems(6)
                  FXCYAmt = Me.ListView1.ListItems.Item(i).SubItems(7)
                  Jn = Me.ListView1.ListItems.Item(i).SubItems(8)
                  descr = Me.ListView1.ListItems.Item(i).SubItems(9)
                  Class = Me.ListView1.ListItems.Item(i).SubItems(10)
                    
                  'If dR <> "" Then
                     With RstBA
                           .addnew
                           !SerialNo = Jn
                           !ticket = TN
                           !accountnumber = Ac
                           !accountname = an
                           !Description = descr
                           !DebitAmount = dr
                           !creditamount = cR
                           !fxcyAmount = FXCYAmt
                           !TRansDate = Format(Date, "mm/dd/yyyy")
                           !Status = "Unposted"
                           !Prepby = cLogUser
                           !Classification = Class
                           !CheckOwnedby = Ac
                           !CheckNo = CheckNo
                           !Payee = Payee
                           !Checktype = Checktype
                           !Codetype = Codetype
                           !Trantype = Trantype
                           !ORno = ORno
                           .Update
                      End With
                   'End If
               Next
               
              'Make Temptable empty
               xtemp.Open "delete TempBankJournal", constring, adOpenKeyset, adLockOptimistic, adCmdText
               Me.ListView1.ListItems.clear
             
             End If
             
             Me.Combo6 = ""
             Me.Combo7 = ""
             Me.Combo12 = ""
             Me.Combo13 = ""
             Me.MaskEdBox1.Text = ""
             Me.MaskEdBox2.Text = ""
             Me.Combo8 = ""
            
             If Me.Text2.Text = "Edit" Then
                Me.Text2.Text = "text"
                'Unload Me
              Else
             On Error Resume Next
              Me.Combo6.SetFocus
             End If
             On Error GoTo 0
            
            i = 0
            cItems = Me.ListView3.ListItems.Count
            For i = 1 To cItems
                If i >= cItems Then
                    Exit For
                End If
                If Me.ListView3.ListItems.Item(i).ForeColor = &H80000004 Then
                       cindex = Me.ListView3.ListItems.Item(i).Index
                       Me.ListView3.ListItems.Remove cindex
                       i = i - 1
                End If
                cItems = Me.ListView3.ListItems.Count
             Next
             
             'Remove the selectd Transactions items in listview3 which is already transact.
             i = 0
             If Left(Me.Combo9, 2) = "WQ" Or Left(Me.Combo9, 2) = "DQ" Then
                 Dim RsVoucher As New ADODB.Recordset
                 For i = 1 To Me.ListView3.ListItems.Count
                   If i > Me.ListView3.ListItems.Count Then Exit Sub
                    If Me.ListView3.ListItems.Item(i).Checked = True Then
                       cindex = Me.ListView3.ListItems.Item(i).Index
                       rcpNo = Trim(Me.ListView3.ListItems.Item(i).SubItems(1))
                       RsVoucher.Open "Update vouchers set Paymode='08-CashOnBank' where receiptno=" & "'" & rcpNo & "'" & "and Left(Paymode,2)='03'", constring, adOpenKeyset, adLockPessimistic, adCmdText
                       Me.ListView3.ListItems.Remove cindex
                       i = i - 1
                     End If
                    Next
               
               ElseIf Trim(Me.Combo9) = "[All]" Then  'Or Me.Combo9 = "DQ" Then
                 Dim rstBT As New ADODB.Recordset
                 For i = 1 To Me.ListView3.ListItems.Count
                   If i > Me.ListView3.ListItems.Count Then Exit Sub
                  If Me.ListView3.ListItems.Item(i).Checked = True Then
                   cindex = Me.ListView3.ListItems.Item(i).Index
                   acctNo = Trim(Me.ListView3.ListItems.Item(i).SubItems(2))
                   cItem = Trim(Me.ListView3.ListItems.Item(i).SubItems(1))
                   rstBT.Open "Update BankTransaction set status='1' where accountcode=" & "'" & acctNo & "'" & "and refno=" & "'" & cItem & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
                   Me.ListView3.ListItems.Remove cindex
                   i = i - 1
                  End If
                  Next
                  
                 ElseIf Left(Me.Combo9, 1) = "T" Then  'Or Me.Combo9 = "DQ" Then
                 Dim rstpay As New ADODB.Recordset
                 For i = 1 To Me.ListView3.ListItems.Count
                   If i > Me.ListView3.ListItems.Count Then Exit Sub
                  If Me.ListView3.ListItems.Item(i).Checked = True Then
                   cindex = Me.ListView3.ListItems.Item(i).Index
                   acctNo = Trim(Me.ListView3.ListItems.Item(i).SubItems(2))
                   cItem = Trim(Me.ListView3.ListItems.Item(i).SubItems(6))
                   RefNo = Trim(Me.ListView3.ListItems.Item(i).SubItems(1))
                   rstBT.Open "Update payablesetup set paidmark='1' where accno=" & "'" & acctNo & "'" & "and Serialno=" & "'" & cItem & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
                   rstBT.Open "Update BankTransaction set status='1' where accountcode=" & "'" & acctNo & "'" & "and refno=" & "'" & RefNo & "'" & " and Left(TranType,2)='TO'", constring, adOpenKeyset, adLockPessimistic, adCmdText
                   Me.ListView3.ListItems.Remove cindex
                   i = i - 1
                  End If
                  Next
                  
                 'for Cash dep from Cashier
                 ElseIf Left(Me.Combo9, 1) = "D" Then  'Or Me.Combo9 = "DQ" Then
                 i = 0
                 For i = 1 To Me.ListView3.ListItems.Count
                   If i > Me.ListView3.ListItems.Count Then Exit Sub
                  If Me.ListView3.ListItems.Item(i).Checked = True Then
                   cindex = Me.ListView3.ListItems.Item(i).Index
                   acctNo = Trim(Me.ListView3.ListItems.Item(i).SubItems(2))
                   cItem = Trim(Me.ListView3.ListItems.Item(i).SubItems(6))
                   RefNo = Trim(Me.ListView3.ListItems.Item(i).SubItems(1))
                   receiptno = Trim(Me.ListView3.ListItems.Item(i).SubItems(8))
                   rstBT.Open "Update Vouchers set ItsDeposit='Yes' where receiptno=" & "'" & receiptno & "'" & " and Left(payopt,3)='006' and left(paymode,2)='01'", constring, adOpenKeyset, adLockPessimistic, adCmdText
                   rstBT.Open "Update BankTransaction set status='1' where accountcode=" & "'" & acctNo & "'" & "and refno=" & "'" & RefNo & "'" & " and Left(TranType,2)='DC'", constring, adOpenKeyset, adLockPessimistic, adCmdText
                   Me.ListView3.ListItems.Remove cindex
                   i = i - 1
                  End If
                  Next
             DisplayTransToday
             End If
    '--------------------------------------------------------------
    Case Is = "Print"
         PrintEntries
    '--------------------------------------------------------------
    Case Is = "Return"
       Unload Me
End Select
End Sub
Sub AddEntries()
If Val(Me.Combo23) = 0 Then
    msg = MsgBox("Please enter Foreign Currency amount", vbExclamation + vbOKOnly, "Message")
    Me.Combo23.SetFocus
    Exit Sub
End If

Dim RstBA As ADODB.Recordset
Set RstBA = New ADODB.Recordset
'tsik the maximum trans then compared it the number of transaction in the listview.
 Dim MaxTranPerJNL As Long
 RstBA.Open "Setup", constring, adOpenKeyset, adLockPessimistic, adCmdTable
 MaxTranPerJNL = RstBA!MaxTranPerJNL
 If Me.ListView1.ListItems.Count >= MaxTranPerJNL Then
    xmsg = MsgBox("Sorry, You have reach the maximum transaction of " & Str(MaxTranPerJNL) & " Per Journal Entry.", vbOKOnly + vbInformation, "Message")
    Me.Toolbar1.Buttons(1).Enabled = False
    RstBA.close
    Exit Sub
   Else
    Me.Toolbar1.Buttons(1).Enabled = True
 End If
 RstBA.close

If Trim(Me.MaskEdBox1.Text) = "" Or Val(Me.MaskEdBox1.Text) = ".00" Then
   If Me.Combo6 <> "" Then
    xmsg = MsgBox("Please enter Transaction Amount", vbExclamation + vbOKOnly, "Message")
    Exit Sub
   End If
End If

If Trim(Me.MaskEdBox2.Text) = "" Or Val(Me.MaskEdBox2.Text) = ".00" Then
    If Me.Combo7 <> "" Then
      xmsg = MsgBox("Please enter Transaction Amount", vbExclamation + vbOKOnly, "Message")
      Exit Sub
    End If
End If

vtn = Val(Me.Text4.Text) '+ Val(Me.Combo1)
'Dim vtn As Long
If vtn = 0 Then
 vtn = Me.ListView1.ListItems.Count + 1 'Val(Me.Text4.Text) 'Val(Me.Combo1)
End If
If vtn <> 0 Then
 NextTn = vtn ' IIf(ValText4 = 0, vtn + 1, Val(Me.Text4.Text))
Else
 Dim TN1 As Long
 RstBA.Open "SElect count(Ticket) as [tn] from BankJournal where Transdate=" & "'" & Date & "'", constring, adOpenDynamic, adLockOptimistic, adCmdText
 If Me.ListView1.ListItems.Count = 0 Then
  TN1 = Me.ListView1.ListItems.Count + 1 'rstBA!TN
  Else
  TN1 = Me.ListView1.ListItems.Count + 1
 End If
 If Me.Text4.Text = "Edit" Then
  If Val(Me.ListView1.ListItems.Item(1).SubItems(1)) <> 1 Then
   TN1 = Val(Me.ListView1.ListItems.Item(1).SubItems(1)) + Val(Me.ListView1.ListItems.Count)
  End If
 End If
 NextTn = TN1
 RstBA.close
End If
On Error GoTo 0
If Me.Combo6 <> "" Then
    Set MItem = Me.ListView1.ListItems.Add(, , NextTn, , 1)
    MItem.SubItems(1) = NextTn
    MItem.SubItems(2) = Me.Combo6
    MItem.SubItems(3) = Me.Combo12
    MItem.SubItems(4) = Me.Combo2
    MItem.SubItems(5) = Me.MaskEdBox1.Text
    MItem.SubItems(6) = "0.00"
    MItem.SubItems(7) = Combo23
    MItem.SubItems(8) = Combo11
    MItem.SubItems(9) = Me.Combo8
    MItem.SubItems(10) = DrCat
    
    'save also to temp table
    With rstTemp
        .addnew
        !ticket = IIf(TN = 0, NextTn, TN)
        !accountnumber = Me.Combo6
        !accountname = Me.Combo12
        !TRansDate = Format(Me.Combo2, "mm/dd/yyyy")
        !DebitAmount = Me.MaskEdBox1.Text
        !creditamount = 0
        !fxcyAmount = Me.Combo23
        On Error Resume Next
        !SerialNo = Me.Combo11
        !Description = Me.Combo8
        !deletemark = 0
        !Status = "Unposted"
        !Classification = DrCat
        !Prepby = cLogUser
        !CheckOwnedby = CheckOwnedby
        !CheckNo = CheckNo
        !Payee = Payee
        !Checktype = Checktype
        !Trantype = Trantype
        !ORno = ORno
        .Update
    End With
    On Error GoTo 0
    NextTn = Me.ListView1.ListItems.Count + 1 'Val(Me.Combo1)
 End If

If Me.Combo7 <> "" Then
       Set MItem = Me.ListView1.ListItems.Add(, , NextTn, , 1)
       MItem.SubItems(1) = NextTn
       MItem.SubItems(2) = Me.Combo7
       MItem.SubItems(3) = Me.Combo13
       MItem.SubItems(4) = Me.Combo2
       MItem.SubItems(5) = "0.00"
       MItem.SubItems(6) = Me.MaskEdBox2.Text
       MItem.SubItems(7) = Me.Combo23
       MItem.SubItems(8) = Me.Combo11
       MItem.SubItems(9) = Me.Combo8
       MItem.SubItems(10) = CrCat
       
       'save also to temp table
       With rstTemp
        .addnew
        !ticket = NextTn
        !accountnumber = Me.Combo7
        !accountname = Me.Combo13
        !TRansDate = Format(Me.Combo2, "mm/dd/yyyy")
        On Error Resume Next
        !DebitAmount = Me.MaskEdBox1.Text
        !creditamount = Me.MaskEdBox2.Text
        !DebitAmount = 0
        !fxcyAmount = Me.Combo23
        !SerialNo = Me.Combo11
        !Description = Me.Combo8
        !deletemark = 0
        !Status = "Unposted"
        !Classification = CrCat
        .Update
       End With
       On Error GoTo 0
       
       Me.Combo7 = ""
       Me.Combo13 = ""
       Me.MaskEdBox2.Text = ""
       Me.Combo12 = ""
       Me.Combo8 = ""
       vtn = vtn + 1 '+ Val(Me.Combo1)
 End If
ListView1.SortKey = 1
ListView1.Sorted = True
Me.Combo14 = Me.ListView1.ListItems.Count
 
 
 'clear the entries
On Error Resume Next
 Me.Combo6 = ""
 Me.Combo12 = ""
 Me.MaskEdBox1.Text = ""
 Me.Combo8 = ""
 Me.Combo6.SetFocus
 Me.Text4.Text = ""
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key
    Case Is = "Save"
          If Me.Combo15.Enabled = True Then
            If Me.Combo15 = "" Or Me.MaskEdBox3.ClipText = "" Then
                mess = MsgBox("Please enter filled up all the fields completely", vbExclamation + vbOKOnly, "Message")
                Exit Sub
            End If
           End If
           Call SaveTRan
    Case Is = "Delete"
        Call XdeleteMain3_Click
End Select
End Sub

Private Sub Toolbar3_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case Is = "find"
            Dim strFindMe As String
            Dim intSelectedOption As String
            Dim itmFnd As ListItem   ' FoundItem variable.
            intSelectedOption = lvwSubItem
            strFindMe = Trim(Me.Combo10)
            Set itmFnd = Me.ListView3.Finditem(strFindMe, intSelectedOption, , lvwPartial)
             If itmFnd Is Nothing Then  ' If no match, inform user and exit.
               msg = MsgBox("Search Text is not found", vbExclamation + vbOKOnly, "Find")
               Me.Combo10.SetFocus
               Exit Sub
               Else
                 
                itmFnd.EnsureVisible
                itmFnd.Selected = True
                Me.ListView3.SetFocus
             End If
End Select
End Sub

Private Sub XdeleteMain3_Click()
If Me.ListView3.ListItems.Count <> 0 Then
    cindex = Me.ListView3.SelectedItem.Index
    RefNo = Trim(Me.ListView3.SelectedItem.SubItems(1))
    Dim rsRef As New ADODB.Recordset
    
    mess = MsgBox("Delete " & Trim(Me.ListView3.SelectedItem.SubItems(3)) & "?", vbQuestion + vbYesNo + vbDefaultButton2, "Please confirm")
    If mess = vbYes Then
        Me.ListView3.ListItems.Remove cindex
        rsRef.Open "delete banktransaction where refno =" & "'" & RefNo & "'" & "and accountcode=" & "'" & Trim(Me.Combo20) & "'", constring, adOpenKeyset, adLockPessimistic
    End If
End If
End Sub

Private Sub xDetails_Click()
Me.ListView2.View = lvwReport
Me.ListView3.View = lvwReport
Me.Li.Checked = False
Me.SI.Checked = False
Me.xlist.Checked = False
Me.xDetails.Checked = True

End Sub

Private Sub xDownLoad_Click()
DownLoadIvy.Show
End Sub

Private Sub xFind_Click()
FindItemLV.Show 1
End Sub

Private Sub xList_Click()
Me.ListView2.View = lvwList
Me.ListView3.View = lvwList
Me.Li.Checked = False
Me.SI.Checked = False
Me.xlist.Checked = True
Me.xDetails.Checked = False

End Sub


Private Sub xPreview_Click()
Dim todayDate As Date
todayDate = Me.ListView2.SelectedItem.SubItems(4)
On Error Resume Next
If UserRole = "Admin" Then
 FinanceDE.rsBankJournalUnpost.close
 FinanceDE.rsBankJournalUnpost.Open "Select * from Bankjournal where transdate =" & "'" & todayDate & "'" & " and remarks is null order by serialno", constring, adOpenKeyset, adLockPessimistic, adCmdText
 FinanceDE.BankJOurnalUnpost todayDate, cLogUser
Else
 FinanceDE.rsBankJournalUnpost.close
 FinanceDE.BankJOurnalUnpost todayDate, cLogUser
End If
 
BankJOurnalUnpost.Show 1

End Sub

Private Sub xREfresh_Click()
DisplayTransToday
End Sub

Private Sub xSelectall_Click()
Dim cAmt As Currency
Dim MaskedAmt As Currency
i = 0
If Me.xSelectall.Checked = False Then
 For i = 1 To Me.ListView3.ListItems.Count
    If IsEmpty(Me.MaskEdBox1.Text) = True Or Trim(Me.MaskEdBox1.Text) = "" Then
      MaskedAmt = 0
     Else
      MaskedAmt = Me.MaskEdBox1.Text
     End If
    cAmt = Me.ListView3.ListItems.Item(i).SubItems(4)
    Me.ListView3.ListItems.Item(i).Checked = True
    Me.ListView3.ListItems.Item(i).ForeColor = &H8000000F
    Me.xSelectall.Checked = True
    Totamt = MaskedAmt + cAmt
    Me.MaskEdBox1.Text = Format(Totamt, "###,###,###.#0")
    Me.MaskEdBox2.Text = Format(Totamt, "###,###,###.#0")
 Next
 
 Else
 
 For i = 1 To Me.ListView3.ListItems.Count
    MaskedAmt = Me.MaskEdBox1.Text
    cAmt = Me.ListView3.ListItems.Item(i).SubItems(4)
    Me.ListView3.ListItems.Item(i).Checked = False
    Me.ListView3.ListItems.Item(i).ForeColor = vbBlack
    Me.xSelectall.Checked = False
    Totamt = MaskedAmt - cAmt
    Me.MaskEdBox1.Text = Format(Totamt, "###,###,###.#0")
    Me.MaskEdBox2.Text = Format(Totamt, "###,###,###.#0")
 Next
End If
End Sub

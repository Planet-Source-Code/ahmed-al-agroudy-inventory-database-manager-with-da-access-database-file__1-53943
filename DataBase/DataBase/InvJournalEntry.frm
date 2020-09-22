VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form InvJournalEntry 
   Caption         =   "Inventory Journal Entry   ÅÏÎÇá ÞíæÏ ÇáíæãíÉ "
   ClientHeight    =   8490
   ClientLeft      =   75
   ClientTop       =   -12960
   ClientWidth     =   8880
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "InvJournalEntry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   8880
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   8175
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   14420
      _Version        =   393216
      Tab             =   2
      TabHeight       =   617
      TabMaxWidth     =   4410
      MouseIcon       =   "InvJournalEntry.frx":014A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Inventory List ÞÇÆãÉ ÇáãÎÒæä"
      TabPicture(0)   =   "InvJournalEntry.frx":0166
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "ImageList1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ListView3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Data Entry ÇÏÎÇá ÇáÈíÇäÇÊ "
      TabPicture(1)   =   "InvJournalEntry.frx":0182
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label22"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label11"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label8"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Combo3"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Text2"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Text1"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Text3"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "ImageList2"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "CoolBar1"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Combo15"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Combo10"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Combo9"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Frame3"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Frame4"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Frame1"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Text5"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Text4"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Text6"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Timer1"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Frame6"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Frame5"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "Check1"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).ControlCount=   23
      TabCaption(2)   =   "Journal List ÞÇÆãÉ ÇáÚÇã"
      TabPicture(2)   =   "InvJournalEntry.frx":019E
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "ListView2"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin MSComctlLib.ListView ListView2 
         Height          =   7695
         Left            =   0
         TabIndex        =   65
         Top             =   480
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   13573
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
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
         Left            =   -66600
         TabIndex        =   64
         Top             =   50
         Width           =   3255
      End
      Begin VB.Frame Frame5 
         Caption         =   "  Journal Entries ÇÎÇá ÚÇã "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   2295
         Left            =   -74760
         TabIndex        =   1
         Top             =   5010
         Width           =   11175
         Begin MSComctlLib.ListView ListView1 
            Height          =   1935
            Left            =   240
            TabIndex        =   63
            Top             =   240
            Width           =   10695
            _ExtentX        =   18865
            _ExtentY        =   3413
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
            NumItems        =   10
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
               SubItemIndex    =   7
               Text            =   "JournalNo"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "Descriptions"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "Classification"
               Object.Width           =   3528
            EndProperty
         End
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   7710
         Left            =   -74925
         TabIndex        =   55
         Top             =   405
         Width           =   11580
         _ExtentX        =   20426
         _ExtentY        =   13600
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "TransDate"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "GR No."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Invty Cat"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Voucher No"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "DO Number"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Amount"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Fr Cost Center  |  Dept."
            Object.Width           =   3704
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "To Cost Center   |   Dept."
            Object.Width           =   3704
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Purpose"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "W.O. No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "TranType"
            Object.Width           =   2540
         EndProperty
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
         Left            =   -74760
         TabIndex        =   41
         Top             =   480
         Width           =   5175
         Begin VB.ComboBox Combo11 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1680
            Locked          =   -1  'True
            Style           =   1  'Simple Combo
            TabIndex        =   44
            Top             =   600
            Width           =   2175
         End
         Begin VB.ComboBox Combo2 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1680
            Locked          =   -1  'True
            Style           =   1  'Simple Combo
            TabIndex        =   43
            Top             =   240
            Width           =   2175
         End
         Begin VB.ComboBox Combo14 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1680
            Locked          =   -1  'True
            Style           =   1  'Simple Combo
            TabIndex        =   42
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
            TabIndex        =   50
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
            TabIndex        =   49
            Top             =   675
            Width           =   1335
         End
         Begin VB.Label Label34 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "ÇáÊÇÑíÎ"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   3960
            TabIndex        =   48
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
            TabIndex        =   47
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
            TabIndex        =   46
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
            TabIndex        =   45
            Top             =   1005
            Width           =   735
         End
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   -73080
         Top             =   4635
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -74400
         TabIndex        =   37
         Text            =   "Text6"
         Top             =   1275
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73680
         TabIndex        =   36
         Text            =   "when editing TN placed here"
         Top             =   1275
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -71880
         TabIndex        =   35
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
         Left            =   -69360
         TabIndex        =   15
         Top             =   480
         Width           =   5775
         Begin VB.ComboBox Combo5 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1440
            Locked          =   -1  'True
            Style           =   1  'Simple Combo
            TabIndex        =   18
            Text            =   "0.00"
            Top             =   600
            Width           =   2175
         End
         Begin VB.ComboBox Combo4 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            ItemData        =   "InvJournalEntry.frx":01BA
            Left            =   1440
            List            =   "InvJournalEntry.frx":01BC
            Locked          =   -1  'True
            Style           =   1  'Simple Combo
            TabIndex        =   17
            Text            =   "0.00"
            Top             =   240
            Width           =   2175
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1440
            Locked          =   -1  'True
            Style           =   1  'Simple Combo
            TabIndex        =   16
            Text            =   "0"
            Top             =   960
            Width           =   2175
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " ÅÌãÇáí ÇáØÑÝ ÇáÏÇÆä"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   3600
            TabIndex        =   24
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ÅÌãÇáí ÇáØÑÝ ÇáãÏíä"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   3600
            TabIndex        =   23
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ÚÏÏ ÇáÍÑßÇÊ "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   3960
            TabIndex        =   22
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
            TabIndex        =   21
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
            TabIndex        =   20
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
            TabIndex        =   19
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
         Left            =   -74760
         TabIndex        =   13
         Top             =   4200
         Width           =   7815
         Begin VB.ComboBox Combo8 
            Height          =   315
            Left            =   240
            Style           =   1  'Simple Combo
            TabIndex        =   60
            Top             =   240
            Width           =   7335
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
            TabIndex        =   14
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
         Left            =   -74760
         TabIndex        =   5
         Top             =   3075
         Width           =   11175
         Begin MSMask.MaskEdBox MaskEdBox2 
            Height          =   350
            Left            =   8400
            TabIndex        =   62
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
            Height          =   315
            Left            =   2640
            TabIndex        =   59
            Top             =   480
            Width           =   5175
         End
         Begin VB.ComboBox Combo7 
            Height          =   315
            Left            =   240
            TabIndex        =   58
            Top             =   480
            Width           =   2175
         End
         Begin VB.CommandButton Command5 
            BackColor       =   &H8000000C&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   7800
            Picture         =   "InvJournalEntry.frx":01BE
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   480
            Width           =   375
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ÅÓã ÇáÍÓÇÈ"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   6720
            TabIndex        =   12
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label14 
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "Account Na&me"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2640
            TabIndex        =   11
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "&Credit Amount"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   8400
            TabIndex        =   10
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "ÑÞã ÇáÍÓÇÈ"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   1560
            TabIndex        =   9
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label17 
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "Account C&ode"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   240
            TabIndex        =   8
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   10200
            TabIndex        =   7
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.ComboBox Combo9 
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
         Left            =   -74520
         TabIndex        =   4
         Top             =   7635
         Width           =   3015
      End
      Begin VB.ComboBox Combo10 
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
         Left            =   -71040
         TabIndex        =   3
         Top             =   7635
         Width           =   3375
      End
      Begin VB.ComboBox Combo15 
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
         Left            =   -67320
         TabIndex        =   2
         Top             =   7635
         Width           =   3495
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
            NumListImages   =   14
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "InvJournalEntry.frx":0758
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "InvJournalEntry.frx":07B6
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "InvJournalEntry.frx":0814
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "InvJournalEntry.frx":0D56
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "InvJournalEntry.frx":1298
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "InvJournalEntry.frx":16EA
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "InvJournalEntry.frx":1B3C
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "InvJournalEntry.frx":1F8E
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "InvJournalEntry.frx":23E0
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "InvJournalEntry.frx":2832
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "InvJournalEntry.frx":2C84
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "InvJournalEntry.frx":2DDE
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "InvJournalEntry.frx":3090
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "InvJournalEntry.frx":34E2
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin ComCtl3.CoolBar CoolBar1 
         Height          =   420
         Left            =   -66840
         TabIndex        =   33
         Top             =   4395
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   741
         BandCount       =   2
         _CBWidth        =   3255
         _CBHeight       =   420
         _Version        =   "6.0.8169"
         MinHeight1      =   360
         Width1          =   2880
         NewRow1         =   0   'False
         MinHeight2      =   360
         Width2          =   1440
         NewRow2         =   0   'False
         Begin MSComctlLib.Toolbar Toolbar1 
            Height          =   330
            Left            =   120
            TabIndex        =   34
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
         Left            =   -72240
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
               Picture         =   "InvJournalEntry.frx":35E4
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "InvJournalEntry.frx":3A36
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "InvJournalEntry.frx":3E88
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "InvJournalEntry.frx":42DA
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "InvJournalEntry.frx":45F4
               Key             =   ""
            EndProperty
         EndProperty
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
         Left            =   -74760
         TabIndex        =   53
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
         Left            =   -74760
         TabIndex        =   52
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
         Left            =   -74760
         TabIndex        =   51
         Text            =   "Text2"
         Top             =   6240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -73800
         TabIndex        =   54
         Top             =   6840
         Visible         =   0   'False
         Width           =   615
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
         Left            =   -74760
         TabIndex        =   25
         Top             =   1995
         Width           =   11175
         Begin MSMask.MaskEdBox MaskEdBox1 
            Height          =   350
            Left            =   8400
            TabIndex        =   61
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
            Height          =   315
            Left            =   2640
            TabIndex        =   57
            Top             =   480
            Width           =   5175
         End
         Begin VB.ComboBox Combo6 
            Height          =   315
            Left            =   240
            TabIndex        =   56
            Top             =   480
            Width           =   2175
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H8000000A&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   340
            Left            =   7800
            Picture         =   "InvJournalEntry.frx":4A46
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Find Accounts"
            Top             =   480
            Width           =   375
         End
         Begin VB.Label Label35 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ÅÓã ÇáÍÓÇÈ"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   6240
            TabIndex        =   32
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   10080
            TabIndex        =   31
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
            TabIndex        =   30
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ÑÞã ÇáÍÓÇÈ"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   1560
            TabIndex        =   29
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "&Debit Amount"
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
            TabIndex        =   28
            Top             =   240
            Width           =   975
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
            TabIndex        =   27
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Label Label8 
         Caption         =   "Prepared by ÇÚÏ ÈæÇÓØÉ "
         Height          =   255
         Left            =   -74520
         TabIndex        =   40
         Top             =   7395
         Width           =   2055
      End
      Begin VB.Label Label11 
         Caption         =   "Noted by ãáÇÍÙÇÊ "
         Height          =   255
         Left            =   -71040
         TabIndex        =   39
         Top             =   7395
         Width           =   1695
      End
      Begin VB.Label Label22 
         Caption         =   "Approved by ÇáÊæÞíÚ ÈæÇÓØÉ "
         Height          =   255
         Left            =   -67320
         TabIndex        =   38
         Top             =   7395
         Width           =   2055
      End
   End
   Begin VB.Menu Main 
      Caption         =   "Main"
      Visible         =   0   'False
      Begin VB.Menu xDownLoad 
         Caption         =   "Download Items Taken...ÊÍãíá ÇáÇÕäÇÝ "
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
End
Attribute VB_Name = "InvJournalEntry"
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

'item inventory
Dim GN As String
Dim IC As String
Dim VN As String
Dim DOno As String
Dim FrCC As String
Dim FrDept As String
Dim ToCC As String
Dim ToDept As String
Dim Purpose As String
Dim WOno As String
Dim TrnType As String
Dim DrCat As String
Dim CrCar As String

Dim MItem As ListItem
Dim xcol As ColumnHeader
Dim NextTn As Long
Dim acctnames As New ADODB.Recordset
Dim AcctCode As New ADODB.Recordset
Dim rstTemp As New ADODB.Recordset
Dim TAbClic As Integer
Dim xCtrlKeyPress2 As Integer       ' account name to account Number.
Sub PrintEntries()
'tsik the setting for right to left printing
           Set RstBA = New ADODB.Recordset
           RstBA.Open "Setup", constring, adOpenKeyset, adLockPessimistic, adCmdTable
           If RstBA!RightToLeftPrn = "Yes" Then
               Printer.RightToLeft = True
              Else
                Printer.RightToLeft = False
           End If
           RstBA.close
           
           Printer.Orientation = 1
           Printer.FontName = "Arabic Transparent"
           Printer.Print
           Printer.Print
           Printer.Print
           Printer.Print
           Printer.FontSize = 12
           Printer.Print ; Tab(50); Me.Combo11
           Printer.Print
           Printer.Print ; Tab(8); Date
           Printer.Print ; Tab(8); Time
           Printer.Print
           Printer.Print
           Printer.FontSize = 10
           
           i = 0
           For i = 1 To Me.ListView1.ListItems.Count
                  TN = Me.ListView1.ListItems.Item(i).SubItems(1)
                  Ac = Me.ListView1.ListItems.Item(i).SubItems(2)
                  an = Me.ListView1.ListItems.Item(i).SubItems(3)
                  dr = FormatNumber(Me.ListView1.ListItems.Item(i).SubItems(5), 2, vbTrue, vbTrue, vbTrue)
                  cR = FormatNumber(Me.ListView1.ListItems.Item(i).SubItems(6), 2, vbTrue, vbTrue, vbTrue)
                  Jn = Me.ListView1.ListItems.Item(i).SubItems(7)
                  'fcy = Me.ListView1.ListItems.Item(i).SubItems(8)
                  descr = Me.ListView1.ListItems.Item(i).SubItems(8)
                  drCol = 110 - Len(cR)
                  CrCol = 90 - Len(dr)
                  Printer.Print ; Tab(10); an; Tab(50); Ac _
                              ; Tab(drCol - Len(dr)); IIf(cR <> 0, cR, "") _
                              ; Tab(CrCol - Len(cR)); IIf(dr <> 0, dr, "")
                  If dr = 0 Then
                    Printer.Print ; Tab(10); "Desc : " & descr
                  End If
           Next i
           Printer.EndDoc
End Sub
 Sub DeleteItems()
 xmsg = MsgBox("Delete TN " & Me.ListView1.SelectedItem & "?", vbQuestion + vbYesNo + vbDefaultButton2, "Please confirm")
        If xmsg = vbYes Then
            TN = Me.ListView1.SelectedItem.SubItems(1)
            Dim xtemp As New ADODB.Recordset
            xtemp.Open "delete TempInvyJournal  where ticket=" & "'" & TN & "'", constring, adOpenKeyset, adLockOptimistic, adCmdText
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
            'xTemp.Open "SElect count(ticket) as [tn] from InventoryJournal where transdate=" & "'" & Xdate & "'", conString, adOpenDynamic, adLockOptimistic, adCmdText
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
                    xtemp.Open "delete TempInvyJournal  where ticket=" & "'" & xTN & "'", constring, adOpenKeyset, adLockOptimistic, adCmdText
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
                            !Status = "Unposted"
                            .Update
                      End With
                    Next
                        'add again but with the new TN
                       Me.ListView1.ListItems.clear
                       rstTemp.close
                       rstTemp.Open "TempInvyJournal", constring, adOpenKeyset, adLockOptimistic, adCmdTable
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
                 
                 xtemp.Open "delete TempInvyJournal  where ticket=" & "'" & xTN & "'", constring, adOpenKeyset, adLockOptimistic, adCmdText
                End If
                    ListView1.SortKey = 1
                    ListView1.Sorted = True
          End If
          Exit Sub
 End Sub
 Sub DisplayTransToday()
 Me.ListView2.ListItems.clear
 Dim rstTRansToday As New ADODB.Recordset
 rstTRansToday.Open "Select * from InventoryJournal where remarks is null order by serialno", CON1, adOpenKeyset, adLockOptimistic, adCmdText
       If rstTRansToday.EOF = False Then
        rstTRansToday.MoveFirst
       End If
        While rstTRansToday.EOF = False
          On Error Resume Next
          Set MItem = Me.ListView2.ListItems.Add(, , rstTRansToday!ticket, , 1)
          MItem.SubItems(1) = rstTRansToday!ticket
          MItem.SubItems(2) = rstTRansToday!accountnumber
          MItem.SubItems(3) = rstTRansToday!accountname
          MItem.SubItems(4) = Format(rstTRansToday!TRansDate, "dd/mm/yyyy")
          MItem.SubItems(5) = Format(rstTRansToday!DebitAmount, "###,###,###.#0")
          MItem.SubItems(6) = Format(rstTRansToday!creditamount, "###,###,###.#0")
          MItem.SubItems(7) = rstTRansToday!SerialNo
          MItem.SubItems(8) = rstTRansToday!Description
          MItem.SubItems(9) = rstTRansToday!Classification
          MItem.SubItems(10) = rstTRansToday!Prepby
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

Private Sub Combo10_LostFocus()

xdecimal = 0
Me.Combo10.Text = Format(Me.Combo10.Text, "###,###,###,###.#0")

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
Me.caption = "Inventory Journal Entry   ÅÏÎÇá ÞíæÏ ÇáíæãíÉ " & "//" & catName
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
Me.caption = "Inventory Journal Entry   ÅÏÎÇá ÞíæÏ ÇáíæãíÉ " & "//" & catName
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
Me.caption = "Inventory Journal Entry   ÅÏÎÇá ÞíæÏ ÇáíæãíÉ " & "//" & catName
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

Private Sub Command1_Click()
If Me.Combo6 <> "" Then
 Set MItem = Me.ListView1.ListItems.Add(, , Val(Me.Combo1) + 1, , 1)
 MItem.SubItems(2) = Me.Combo6
 MItem.SubItems(3) = Me.Combo12
 MItem.SubItems(4) = Date
 MItem.SubItems(5) = Me.MaskEdBox1.Text
 MItem.SubItems(6) = Me.MaskEdBox2.Text
 MItem.SubItems(7) = Me.Combo11
 MItem.SubItems(8) = Me.Combo9
End If
If Me.Combo7 <> "" Then
 Set MItem = Me.ListView1.ListItems.Add(, , Val(Me.Combo1) + 1, , 1)
 MItem.SubItems(2) = Me.Combo7
 MItem.SubItems(3) = Me.Combo13
 MItem.SubItems(4) = Date
 MItem.SubItems(5) = Me.MaskEdBox1.Text
 MItem.SubItems(6) = Me.MaskEdBox2.Text
 MItem.SubItems(7) = Me.Combo11
 MItem.SubItems(8) = Me.Combo9
End If

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
Me.caption = "Inventory Journal Entry   ÅÏÎÇá ÞíæÏ ÇáíæãíÉ " & "//" & catName
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
           
            xtemp.Open "delete tempInvyJournal  where ticket=" & "'" & strFindMe & "'", constring, adOpenKeyset, adLockOptimistic, adCmdText
            AddEntries
        End If
        Me.Combo6.SetFocus
    
      End If
    End If
 End If

End Sub

Private Sub Command1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
 Unload Me
End If
            
End Sub

Private Sub Command1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
   Me.Combo9.SetFocus
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command4_Click()
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
        trDate = Me.ListView2.SelectedItem.SubItems(4)
        trandate = Mid(trDate, 4, 2) & "/" & Left(trDate, 2) & "/" & Mid(trDate, 7, 4)
        Me.Combo2 = trandate 'put the transdate selected
        Sn = Me.ListView2.SelectedItem.SubItems(7)
        Me.Combo11 = Sn
        'Check TEmptable if i is empty, if not empty we can't edit the a data
        rstTemp.Requery
        If rstTemp.RecordCount <> 0 Then
            mess = MsgBox("There was an unsaved entries that you must first save before editing.", vbInformation + vbOKOnly, "Message")
            Exit Sub
        End If
        rsttran.Open "InventoryJournal", constring, adOpenDynamic, adLockOptimistic, adCmdTable
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
            MItem.SubItems(7) = rsttran!SerialNo
            MItem.SubItems(8) = rsttran!Description
            MItem.SubItems(9) = rsttran!Classification
            'transfer transaction to temptable for modification for deletion
           On Error GoTo 0
             With rstTemp
                .addnew
                !ticket = rsttran!ticket
                !accountnumber = rsttran!accountnumber
                !accountname = rsttran!accountname
                !Classification = rsttran!Classification
                !TRansDate = rsttran!TRansDate
                !DebitAmount = rsttran!DebitAmount
                !creditamount = rsttran!creditamount
                !SerialNo = rsttran!SerialNo
                !Description = rsttran!Description
                !Status = "Unposted"
                !grno = rsttran!grno
                !InventCat = rsttran!InventCat
                !Voucher = rsttran!Voucher
                !DOno = rsttran!DOno
                !fr_CostCenter = rsttran!fr_CostCenter
                !fr_dept = rsttran!fr_dept
                !To_Costcenter = rsttran!To_Costcenter
                !to_dept = rsttran!to_dept
                !Purpose = rsttran!Purpose
                !WOno = rsttran!WOno
                !Trantype = rsttran!Trantype
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
            Me.Combo11 = Me.ListView2.SelectedItem.SubItems(7)
            Me.Combo14 = Me.ListView2.SelectedItem.SubItems(1)
            Me.Text4.Text = Trim(Me.ListView2.SelectedItem.SubItems(1))
          Else
            Me.Combo7 = Me.ListView2.SelectedItem.SubItems(2)
            Me.Combo13 = Me.ListView2.SelectedItem.SubItems(3)
            Me.MaskEdBox2.Text = Me.ListView2.SelectedItem.SubItems(6)
            Me.Combo11 = Me.ListView2.SelectedItem.SubItems(7)
            Me.Combo14 = Me.ListView2.SelectedItem.SubItems(1)
            Me.Text4.Text = Trim(Me.ListView2.SelectedItem.SubItems(1))
        End If
         DisplayTransToday ' refesh the of listview2
    End If
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
  If Left(AcctCode!AccountCode, 4) <> "2212" Then
   If Left(AcctCode!AccountCode, 2) <> "11" Then
       Me.Combo6.AddItem AcctCode!AccountCode
       Me.Combo12.AddItem AcctCode!accountnameeng & "\" & RTrim(AcctCode!accountnamearab)
     Else
       Me.Combo7.AddItem AcctCode!AccountCode
       Me.Combo13.AddItem AcctCode!accountnameeng & "\" & RTrim(AcctCode!accountnamearab)
     End If
   End If
  End If
  AcctCode.MoveNext
  Wend
 AcctCode.close
 
 
 
 If Me.Text2.Text <> "Edit" Then
        Set RstBA = New ADODB.Recordset
        

        'Check Temp table if it is empty or not. if not put it on the list view
         On Error GoTo Nelson
         rstTemp.Open "TempInvyJournal", CON1, adOpenKeyset, adLockOptimistic, adCmdTable
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
                   Jn = "IVY" & "-" & Right(Year(Date), 2) & Trim(Format(Date, "mm")) & "-" & JOurnalNo!nextjn
                   JOurnalNo.close
                Else
                   Jn = "IVY" & "-" & Right(Year(Date), 2) & Trim(Format(Date, "mm")) & "-" & JOurnalNo!nextjn
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
                        ElseIf Len(nextjn) = 6 Then
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
            MItem.SubItems(4) = rstTemp!TRansDate
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
  rstTemp.Open "TempInvyJournal", CON1, adOpenKeyset, adLockOptimistic, adCmdTable
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
            DelTran.Open "Delete TempInvyJournal", constring, adOpenDynamic, adLockPessimistic, adcmdttext
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
    rsttran.Open "InventoryJournal", constring, adOpenDynamic, adLockOptimistic, adCmdTable
    rstTemp.close
    rstTemp.Open "TempInvyJournal", constring, adOpenDynamic, adLockOptimistic, adCmdTable
    
    While rstTemp.EOF = False
         With rsttran
            .addnew
            !ticket = rstTemp!ticket
            !accountnumber = rstTemp!accountnumber
            !accountname = rstTemp!accountname
            !TRansDate = rstTemp!TRansDate
            !DebitAmount = rstTemp!DebitAmount
            !creditamount = rstTemp!creditamount
            !SerialNo = rstTemp!SerialNo
            !Description = rstTemp!Description
            '!deletemark = 0
            !Status = "Unposted"
            !grno = GN
            !InventCat = IC
            !Voucher = VN
            !DOno = DOno
            !fr_CostCenter = FrCC
            !fr_dept = FrDept
            !To_Costcenter = ToCC
            !to_dept = ToDept
            !Purpose = Purpose
            !WOno = WOno
            !Trantype = TrnType
            .Update
          End With
        rstTemp.MoveNext
    Wend
    rstTemp.close
    c = Err.Description
    rstTemp.Open "Delete TempInvyJournal", constring, adOpenDynamic, adLockPessimistic, adcmdttext
   
End If
End Sub

Private Sub Fv_Click()
If Me.FV.caption = "&Normal View" Then
  Me.FV.caption = "&Full View"
   InvJournalEntry.Frame5.Top = 4920
   InvJournalEntry.ListView1.Top = 5160
   InvJournalEntry.ListView1.Height = 1935
   InvJournalEntry.Frame5.Height = 2295
  Else
   Me.FV.caption = "&Normal View"
   InvJournalEntry.Frame5.Top = 350
   InvJournalEntry.ListView1.Top = 630
   InvJournalEntry.Frame5.Height = 7600
   InvJournalEntry.ListView1.Height = 7200
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
If Me.ListView1.Top = 630 Then
   Scmenu.FV.caption = "&Full View"
   InvJournalEntry.Frame5.Top = 4940
   InvJournalEntry.ListView1.Top = 5160
   InvJournalEntry.ListView1.Height = 1935
   InvJournalEntry.Frame5.Height = 2295
  Else
   Scmenu.FV.caption = "&Normal View"
   InvJournalEntry.Frame5.Top = 370
   InvJournalEntry.ListView1.Top = 630
   InvJournalEntry.Frame5.Height = 7600
   InvJournalEntry.ListView1.Height = 7200
 End If

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
If Button = 2 Then
   PopupMenu Main2
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
 rsttran.Open "TempInvyJOurnal ", constring, adOpenDynamic, adLockOptimistic, adCmdTable
 If rsttran.EOF = False Then
    xmsg = MsgBox("There was unsaved transactions that you must save first before Deleting or Modifying any saved transactions.", vbInformation + vbOKOnly, "Message")
     Exit Sub
 End If
 rsttran.close
 Me.Text2.Text = "Edit"
 Call eEdit_Click
 Me.SSTab1.SetFocus
 SendKeys "{Left}"

End If

End Sub

Private Sub ListView2_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = 2 Then
  If Trim(Left(Me.SSTab1.caption, 12)) = "Journal List" Then
    Me.xDownLoad.Enabled = False
    
    If Me.ListView2.ListItems.Count = 0 Then
        Me.xFind.Enabled = False
        Me.eEdit.Enabled = False
        Me.xrefresh.Enabled = False
        Me.xPreview.Enabled = False
        Me.POstToGL.Enabled = False
       Else
        Me.xFind.Enabled = True
        Me.eEdit.Enabled = True
        Me.xrefresh.Enabled = True
        Me.xPreview.Enabled = True
        Me.POstToGL.Enabled = True
    End If
    PopupMenu Me.main
   End If
End If
End Sub

Private Sub ListView3_DblClick()
If Me.ListView1.ListItems.Count <> 0 Then
    mess = MsgBox("You have an unsaved entries Please save it first.", vbInformation + vbOKOnly, "Message")
    Me.SSTab1.SetFocus
    SendKeys "{Right}"
    Exit Sub
End If
Dim cAmt As Currency
If Me.ListView3.ListItems.Count = 0 Then
    Exit Sub
End If
InvCat = Trim(Me.ListView3.SelectedItem.SubItems(2))
DOno = Trim(Me.ListView3.SelectedItem.SubItems(4))
Trantype = Trim(Me.ListView3.SelectedItem.SubItems(10))
amount = Trim(Me.ListView3.SelectedItem.SubItems(5))
'item inventory
GN = Trim(Me.ListView3.SelectedItem.SubItems(1))
IC = Trim(Me.ListView3.SelectedItem.SubItems(2))
VN = Trim(Me.ListView3.SelectedItem.SubItems(3))
DOno = Trim(Me.ListView3.SelectedItem.SubItems(4))
FrCC = Trim(Left(Me.ListView3.SelectedItem.SubItems(6), 3))
FrDept = Trim(Right(Me.ListView3.SelectedItem.SubItems(6), 3))
ToCC = Trim(Left(Me.ListView3.SelectedItem.SubItems(7), 3))
ToDept = Trim(Right(Me.ListView3.SelectedItem.SubItems(7), 3))
Purpose = Trim(Me.ListView3.SelectedItem.SubItems(8))
WOno = Trim(Me.ListView3.SelectedItem.SubItems(9))
TrnType = Trim(Me.ListView3.SelectedItem.SubItems(10))

If Trantype = "A" Then
    cTranType = "A"
ElseIf Trantype = "B" Then
cTranType = "B"
ElseIf Trantype = "C" Then
    cTranType = "C"
End If
If Trantype = "D" Then
    cTranType = "D"
End If
    
i = 0
For i = 1 To Me.ListView3.ListItems.Count
    If Trim(Me.ListView3.ListItems.Item(i).SubItems(4)) = DOno And Trim(Me.ListView3.ListItems.Item(i).SubItems(2)) = InvCat Then 'for DO
        Me.ListView3.ListItems.Item(i).ForeColor = &H80000004
        amount = Trim(Me.ListView3.ListItems.Item(i).SubItems(5))
        cAmt = cAmt + Format(amount, "###,###,###.#0")
      Else
        Me.ListView3.ListItems.Item(i).ForeColor = vbBlack
    End If
Next
If cTranType <> "A" Then
   Me.MaskEdBox2.Text = FormatNumber(cAmt, 2, vbTrue, vbTrue, vbTrue)
   Me.MaskEdBox1.Text = FormatNumber(cAmt, 2, vbTrue, vbTrue, vbTrue)
   Me.MaskEdBox2.Enabled = False
   Me.MaskEdBox1.Enabled = True
  Else
  Me.MaskEdBox1.Text = FormatNumber(cAmt, 2, vbTrue, vbTrue, vbTrue)
  Me.MaskEdBox2.Text = FormatNumber(cAmt, 2, vbTrue, vbTrue, vbTrue)
  Me.MaskEdBox1.Enabled = False
  Me.MaskEdBox2.Enabled = True
End If
Me.SSTab1.SetFocus
SendKeys "{Right}"

End Sub

Private Sub ListView3_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = 2 Then
   'Me.SSTab1.SetFocus
   If Trim(Left(Me.SSTab1.caption, 14)) = "Inventory List" Then
    Me.xDownLoad.Enabled = True
    Me.xFind.Enabled = False
    Me.eEdit.Enabled = False
    Me.xrefresh.Enabled = False
    Me.xPreview.Enabled = False
    Me.POstToGL.Enabled = False
    PopupMenu Me.main
   End If
End If
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

Private Sub PostToGl_Click()
Dim rsInvj As New ADODB.Recordset
Dim MItem As ListItem
mess = MsgBox("Do you want to continue?", vbQuestion + vbOKCancel + vbDefaultButton2, "Please confirm")
If mess = vbOK Then
     PostingJournal.Text1.Text = "IVY"
     rsInvj.Open "SELECT  TransDate, COUNT(TransDate) AS TotalTRan, SUM(DebitAmount) AS DrAmt, SUM(CreditAmount) AS CrAmt" _
     & " From InventoryJournal GROUP BY TransDate ORDER BY TransDate", constring, adOpenKeyset, adLockPessimistic, adCmdText
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
 rstPn.Open "SElect count(ticket) as [tn1] from InventoryJournal where transdate=" & "'" & Xdate & "'", constring, adOpenDynamic, adLockOptimistic, adCmdText
 TN = rstPn!TN1  '+ 1
 Me.Combo1.Text = TN ' IIf(tn = 0, 1, tn)
 rstPn.close
  
 'put total credit in combo4
 rstPn.Open "Select SUm(CreditAmount)  as [TotalCr] from InventoryJournal where TransDate= " & "'" & Xdate & "'", constring, adOpenDynamic, adLockOptimistic, adCmdText
 Me.Combo5 = IIf(rstPn!TotalCr <> 0, Format(rstPn!TotalCr, "###,###,###.#0"), "0.00")
 Me.Combo5 = Format(Me.Combo5, "###,###,###.#0")
 rstPn.close
 
'put totals debit in combo5
 rstPn.Open "Select SUm(debitAmount)  as [TotalDb] from InventoryJournal where TransDate= " & "'" & Xdate & "'", constring, adOpenDynamic, adLockOptimistic, adCmdText
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
            xtemp.Open "delete TempInvyJOurnal  where ticket=" & "'" & strFindMe & "'", constring, adOpenKeyset, adLockOptimistic, adCmdText
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
             Jn = "IVY" & "-" & Right(Year(Date), 2) & Trim(Format(Date, "mm")) & "-" & JOurnalNo!nextjn
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
                   Jn = "IVY" & "-" & Right(Year(Date), 2) & Trim(Format(Date, "mm")) & "-" & JOurnalNo!nextjn
                   JOurnalNo.close
                End If
             End If
             Me.Combo11 = Jn
             Me.Combo14 = 1
             
             
             'save entries only to INVJOuranlTrans table
              RstBA.Open "InventoryJournal", constring, adOpenDynamic, adLockOptimistic, adCmdTable
              i = 0
              For i = 1 To Me.ListView1.ListItems.Count
                  TN = Me.ListView1.ListItems.Item(i).SubItems(1)
                  Ac = Me.ListView1.ListItems.Item(i).SubItems(2)
                  an = Me.ListView1.ListItems.Item(i).SubItems(3)
                  dr = Me.ListView1.ListItems.Item(i).SubItems(5)
                  cR = Me.ListView1.ListItems.Item(i).SubItems(6)
                  Jn = Me.ListView1.ListItems.Item(i).SubItems(7)
                  Class = Me.ListView1.ListItems.Item(i).SubItems(9)
                  descr = Me.ListView1.ListItems.Item(i).SubItems(8)
                  If dr <> "" Then
                     With RstBA
                           .addnew
                           !SerialNo = Jn
                           !ticket = TN
                           '!deletemark = 0
                           !accountnumber = Ac
                           !accountname = an
                           !Classification = Class
                           !Description = descr
                           !DebitAmount = dr
                           !creditamount = cR
                           !TRansDate = Format(Date, "mm/dd/yyyy")
                           !Status = "Unposted"
                           '!deletemark = 0
                           !Status = "Unposted"
                           !grno = GN
                           !InventCat = IC
                           !Voucher = VN
                           !DOno = DOno
                           !fr_CostCenter = FrCC
                           !fr_dept = FrDept
                           !To_Costcenter = ToCC
                           !to_dept = ToDept
                           !Purpose = Purpose
                           !WOno = WOno
                           !Trantype = TrnType
                           
                           .Update
                      End With
                   End If
               Next
               
              'Make Temptable empty
               xtemp.Open "delete TEmpInvyJournal", constring, adOpenKeyset, adLockOptimistic, adCmdText
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
            DisplayTransToday
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
'             'Remove the selectd Inventory items in listview3 which is already transact.
'             i = 0
'             For i = 1 To Me.ListView3.ListItems.Count
'                    If Me.ListView3.ListItems.Item(i).Ghosted = True Then
'                       cIndex = Me.ListView3.ListItems.Item(i).Index
'                       Me.ListView3.ListItems.Remove cIndex
'                     End If
'             Next
    '--------------------------------------------------------------
    Case Is = "Print"
         PrintEntries
    '--------------------------------------------------------------
    Case Is = "Return"
       Unload Me
End Select
End Sub
Sub AddEntries()
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

If Me.MaskEdBox1.Text = "" And Me.Combo6 <> "" Then
    xmsg = MsgBox("Please enter Transaction Amount", vbExclamation + vbOKOnly, "Message")
    Exit Sub
End If
If Me.MaskEdBox2.Text = "" And Me.Combo7 <> "" Then
    xmsg = MsgBox("Please enter Transaction Amount", vbExclamation + vbOKOnly, "Message")
    Exit Sub
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
 RstBA.Open "SElect count(Ticket) as [tn] from InventoryJournal where Transdate=" & "'" & Date & "'", constring, adOpenDynamic, adLockOptimistic, adCmdText
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
    MItem.SubItems(7) = Combo11
    MItem.SubItems(8) = Me.Combo8
    MItem.SubItems(9) = DrCat
   
    
    'save also to temp table
    With rstTemp
        .addnew
        !ticket = IIf(TN = 0, NextTn, TN)
        !accountnumber = Me.Combo6
        !accountname = Me.Combo12
        !Classification = DrCat
        !TRansDate = Format(Me.Combo2, "mm/dd/yyyy")
        !DebitAmount = Me.MaskEdBox1.Text
        !creditamount = 0
        On Error Resume Next
        !SerialNo = Me.Combo11
        !Description = Me.Combo8
        !deletemark = 0
        !Status = "Unposted"
        !grno = GN
        !InventCat = IC
        !Voucher = VN
        !DOno = DOno
        !fr_CostCenter = FrCC
        !fr_dept = FrDept
        !To_Costcenter = ToCC
        !to_dept = ToDept
        !Purpose = Purpose
        !WOno = WOno
        !Trantype = TrnType
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
       MItem.SubItems(7) = Me.Combo11
       MItem.SubItems(8) = Me.Combo8
       MItem.SubItems(9) = DrCat
       
       'save also to temp table
       With rstTemp
        .addnew
        !ticket = NextTn
        !accountnumber = Me.Combo7
        !accountname = Me.Combo13
        !Classification = CrCat
        !TRansDate = Format(Me.Combo2, "mm/dd/yyyy")
        On Error Resume Next
        !DebitAmount = Me.MaskEdBox1.Text
        !creditamount = Me.MaskEdBox2.Text
        !DebitAmount = 0
        !SerialNo = Me.Combo11
        !Description = Me.Combo8
        !deletemark = 0
        !Status = "Unposted"
        !grno = GN
        !InventCat = IC
        !Voucher = VN
        !DOno = DOno
        !fr_CostCenter = FrCC
        !fr_dept = FrDept
        !To_Costcenter = ToCC
        !to_dept = ToDept
        !Purpose = Purpose
        !WOno = WOno
        !Trantype = TrnType
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
Dim SelectedDate As Date
TRansDate = Trim(Me.ListView2.SelectedItem.SubItems(4)) ' , "mm/dd/yyyy")
SelectedDate = Mid(TRansDate, 4, 2) & "/" & Left(TRansDate, 2) & "/" & Mid(TRansDate, 7, 4)
On Error Resume Next
FinanceDE.rsInventoryJournal.close
FinanceDE.InventoryJournal SelectedDate
IVYJounralUnpost.Show
End Sub

Private Sub xREfresh_Click()
DisplayTransToday
End Sub

VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form GenJournalEntry 
   Caption         =   "General Journal Entry   ÅÏÎÇá ÞíæÏ ÇáíæãíÉ"
   ClientHeight    =   8490
   ClientLeft      =   75
   ClientTop       =   -1500
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "GenJournalEntry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   8220
      Left            =   120
      TabIndex        =   0
      Top             =   45
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   14499
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabMaxWidth     =   3528
      MouseIcon       =   "GenJournalEntry.frx":014A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Data Entry ÇÏÎÇá ÇáÈíÇäÇÊ"
      TabPicture(0)   =   "GenJournalEntry.frx":0166
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label26"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label8"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label11"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label22"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame7"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "ImageList1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Combo3"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Timer1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Text3"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Text1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Text2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Text6"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Text4"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Text5"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "ImageList2"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "CoolBar1"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Frame2"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Frame1"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Frame6"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Frame4"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Frame3"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Combo9"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Combo10"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Combo15"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Frame5"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "ListView1"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Check1"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).ControlCount=   27
      TabCaption(1)   =   "List  ÞÇÆãÉ"
      TabPicture(1)   =   "GenJournalEntry.frx":0182
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "ListView2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin MSComctlLib.ListView ListView2 
         Height          =   7815
         Left            =   0
         TabIndex        =   68
         Top             =   360
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   13785
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
         Caption         =   "Set to Arabic áÊÍæíá Çáí ÇáÚÑÈí"
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
         Left            =   -66840
         TabIndex        =   64
         Top             =   0
         Width           =   3255
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1935
         Left            =   -74520
         TabIndex        =   45
         Top             =   5160
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
         ForeColor       =   128
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "aa"
            Text            =   "!"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Key             =   "a"
            Text            =   "Ticket #"
            Object.Width           =   1676
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Key             =   "b"
            Text            =   "Account Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Key             =   "c"
            Text            =   "Account Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Key             =   "dd"
            Text            =   "TransDate"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Key             =   "e"
            Text            =   "Debit"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Key             =   "f"
            Text            =   "Credit"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Key             =   "h"
            Text            =   "Journal No"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Key             =   "j"
            Text            =   "Description"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "TransType"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Classification"
            Object.Width           =   10583
         EndProperty
      End
      Begin VB.Frame Frame5 
         Caption         =   "  Journal Entries ÇÏÎÇá ÚÇã "
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
         Height          =   2295
         Left            =   -74760
         TabIndex        =   46
         Top             =   4920
         Width           =   11175
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
         Left            =   -67200
         TabIndex        =   63
         Top             =   7560
         Width           =   3495
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
         Left            =   -70920
         TabIndex        =   61
         Top             =   7560
         Width           =   3375
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
         Left            =   -74640
         TabIndex        =   59
         Top             =   7560
         Width           =   3375
      End
      Begin VB.Frame Frame3 
         Caption         =   "  Credit ÏÇÆä"
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
         TabIndex        =   32
         Top             =   3000
         Width           =   11175
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
            TabIndex        =   34
            Top             =   480
            Width           =   2175
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
            TabIndex        =   37
            Top             =   480
            Width           =   5775
         End
         Begin VB.CommandButton Command5 
            BackColor       =   &H8000000C&
            Height          =   350
            Left            =   8400
            Picture         =   "GenJournalEntry.frx":019E
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   480
            Width           =   375
         End
         Begin MSMask.MaskEdBox MaskEdBox2 
            Height          =   345
            Left            =   8880
            TabIndex        =   41
            Top             =   480
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arabic Transparent"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ÇáãÈáÛ"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   10200
            TabIndex        =   42
            Top             =   240
            Width           =   735
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
            TabIndex        =   33
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "ÑÞã ÇáÍÓÇÈ"
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   1560
            TabIndex        =   35
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "&Credit Amount"
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
            Left            =   8880
            TabIndex        =   40
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
            TabIndex        =   36
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ÅÓã ÇáÍÓÇÈ"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   7320
            TabIndex        =   38
            Top             =   240
            Width           =   1095
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
         TabIndex        =   43
         Top             =   4080
         Width           =   4815
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
            RightToLeft     =   -1  'True
            Style           =   1  'Simple Combo
            TabIndex        =   44
            Top             =   240
            Width           =   4335
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
            TabIndex        =   55
            Top             =   720
            Width           =   855
         End
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
         TabIndex        =   1
         Top             =   360
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
            TabIndex        =   6
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
            Style           =   1  'Simple Combo
            TabIndex        =   3
            Top             =   240
            Width           =   2175
         End
         Begin VB.ComboBox Combo14 
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
            Left            =   1680
            Locked          =   -1  'True
            Style           =   1  'Simple Combo
            TabIndex        =   9
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
            TabIndex        =   7
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
            TabIndex        =   5
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
            TabIndex        =   4
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
            TabIndex        =   2
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
            TabIndex        =   8
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
            TabIndex        =   10
            Top             =   1005
            Width           =   735
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   " Total Today ÇÌãÇáí Çáíæãí"
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
         TabIndex        =   11
         Top             =   360
         Width           =   5775
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
            TabIndex        =   19
            Text            =   "0"
            Top             =   960
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
            ItemData        =   "GenJournalEntry.frx":0738
            Left            =   1440
            List            =   "GenJournalEntry.frx":073A
            Locked          =   -1  'True
            Style           =   1  'Simple Combo
            TabIndex        =   13
            Text            =   "0.00"
            Top             =   240
            Width           =   2175
         End
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
            TabIndex        =   16
            Text            =   "0.00"
            Top             =   600
            Width           =   2175
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
            TabIndex        =   18
            Top             =   1010
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
            TabIndex        =   12
            Top             =   320
            Width           =   1095
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
            TabIndex        =   15
            Top             =   680
            Width           =   975
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ÚÏÏ ÇáÍÑßÇÊ "
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   3960
            TabIndex        =   20
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ÅÌãÇáí ÇáØÑÝ ÇáãÏíä"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   3600
            TabIndex        =   14
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " ÅÌãÇáí ÇáØÑÝ ÇáÏÇÆä"
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   3600
            TabIndex        =   17
            Top             =   600
            Width           =   1695
         End
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
         TabIndex        =   21
         Top             =   1920
         Width           =   11175
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
            TabIndex        =   23
            Top             =   480
            Width           =   2175
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
            Sorted          =   -1  'True
            TabIndex        =   26
            Top             =   480
            Width           =   5775
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H8000000A&
            Height          =   340
            Left            =   8400
            Picture         =   "GenJournalEntry.frx":073C
            Style           =   1  'Graphical
            TabIndex        =   28
            ToolTipText     =   "Find Accounts"
            Top             =   480
            Width           =   375
         End
         Begin MSMask.MaskEdBox MaskEdBox1 
            Height          =   345
            Left            =   8880
            TabIndex        =   30
            Top             =   480
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arabic Transparent"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
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
            TabIndex        =   22
            Top             =   240
            Width           =   1215
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
            Left            =   8880
            TabIndex        =   29
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ÑÞã ÇáÍÓÇÈ"
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   1560
            TabIndex        =   24
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
            TabIndex        =   25
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ÇáãÈáÛ"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   10080
            TabIndex        =   31
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label35 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ÅÓã ÇáÍÓÇÈ"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   6840
            TabIndex        =   27
            Top             =   240
            Width           =   1575
         End
      End
      Begin ComCtl3.CoolBar CoolBar1 
         Height          =   420
         Left            =   -66840
         TabIndex        =   57
         Top             =   4320
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
            TabIndex        =   51
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
         Left            =   -72120
         Top             =   1080
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
               Picture         =   "GenJournalEntry.frx":0CD6
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "GenJournalEntry.frx":1128
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "GenJournalEntry.frx":157A
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "GenJournalEntry.frx":19CC
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "GenJournalEntry.frx":1CE6
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.TextBox Text5 
         Height          =   360
         Left            =   -71760
         TabIndex        =   52
         Text            =   "A/C# placed here when edit"
         Top             =   1200
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   -73560
         TabIndex        =   53
         Text            =   "when editing TN placed here"
         Top             =   1200
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.TextBox Text6 
         Height          =   360
         Left            =   -74280
         TabIndex        =   54
         Text            =   "Text6"
         Top             =   1200
         Visible         =   0   'False
         Width           =   975
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
         TabIndex        =   47
         Text            =   "Text2"
         Top             =   5160
         Visible         =   0   'False
         Width           =   735
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
         TabIndex        =   49
         Text            =   "Text1"
         Top             =   5520
         Visible         =   0   'False
         Width           =   615
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
         TabIndex        =   48
         Text            =   "this identify if not bal"
         Top             =   5760
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   -72960
         Top             =   4560
      End
      Begin VB.ComboBox Combo3 
         Height          =   360
         Left            =   -74640
         TabIndex        =   50
         Top             =   6360
         Visible         =   0   'False
         Width           =   615
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
               Picture         =   "GenJournalEntry.frx":2138
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "GenJournalEntry.frx":2196
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "GenJournalEntry.frx":21F4
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "GenJournalEntry.frx":2736
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "GenJournalEntry.frx":2C78
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "GenJournalEntry.frx":30CA
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "GenJournalEntry.frx":351C
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "GenJournalEntry.frx":396E
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "GenJournalEntry.frx":3DC0
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "GenJournalEntry.frx":4212
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "GenJournalEntry.frx":4664
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "GenJournalEntry.frx":47BE
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "GenJournalEntry.frx":4A70
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "GenJournalEntry.frx":4EC2
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame7 
         Caption         =   "  Transaction Type"
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
         Left            =   -69720
         TabIndex        =   65
         Top             =   4080
         Width           =   2535
         Begin VB.ComboBox Combo16 
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
            Style           =   2  'Dropdown List
            TabIndex        =   66
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label Label24 
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
      Begin VB.Label Label22 
         Caption         =   "Approved by ÇáÊæÞíÚ ÈæÇÓØÉ "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -67200
         TabIndex        =   62
         Top             =   7320
         Width           =   2055
      End
      Begin VB.Label Label11 
         Caption         =   "Noted by ãáÇÍÙÇÊ "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70920
         TabIndex        =   60
         Top             =   7320
         Width           =   1695
      End
      Begin VB.Label Label8 
         Caption         =   "Prepared by ÇÚÏ ÈæÇÓØÉ "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74640
         TabIndex        =   58
         Top             =   7320
         Width           =   2055
      End
      Begin VB.Label Label26 
         Caption         =   "Label26"
         Height          =   255
         Left            =   -74040
         TabIndex        =   56
         Top             =   1560
         Visible         =   0   'False
         Width           =   2175
      End
   End
   Begin VB.Menu Main 
      Caption         =   "Main"
      Visible         =   0   'False
      Begin VB.Menu xNEwjn 
         Caption         =   "Add New  ÇÖÝÉ ÌÏíÏÉ "
      End
      Begin VB.Menu eEdit 
         Caption         =   "Modify/Edit  ÇÖÇÝÉ "
      End
      Begin VB.Menu bar3 
         Caption         =   "-"
      End
      Begin VB.Menu xFind 
         Caption         =   "Find...íÌÏ "
      End
      Begin VB.Menu xPreview 
         Caption         =   "Print Preview... ãÚÇíäÉ ÞÈá ÇáØÈÇÚÉ "
      End
      Begin VB.Menu xREfresh 
         Caption         =   "Refresh ÊÍÏíË "
      End
      Begin VB.Menu bar1 
         Caption         =   "-"
      End
      Begin VB.Menu xView 
         Caption         =   "View íÑí "
         Begin VB.Menu Li 
            Caption         =   "Large Icon ÒÑ ßÈíÑ "
         End
         Begin VB.Menu SI 
            Caption         =   "Small Icon  ÕÛíÑ ÒÑ "
         End
         Begin VB.Menu xList 
            Caption         =   "List ÞÇÆãÉ "
         End
         Begin VB.Menu xDEtails 
            Caption         =   "Details ÊÝÇÕíá "
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu PostToGl 
         Caption         =   "Post to GL...áÕÞ Ýí Ìá "
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu Main2 
      Caption         =   "Main2"
      Visible         =   0   'False
      Begin VB.Menu DElSelectitem 
         Caption         =   "Delete Selected Item ÇáÛÇÁ ÇáÇÕäÇÝ "
      End
      Begin VB.Menu PrintItem 
         Caption         =   "Print items ØÈÇÚÉ ÇáÇÕäÇÝ "
      End
      Begin VB.Menu Fv 
         Caption         =   "Full View ãá ÇáÔÇÔÉ "
      End
   End
End
Attribute VB_Name = "GenJournalEntry"
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
Dim MItem As ListItem
Dim xcol As ColumnHeader
Dim NextTn As Long
Dim acctnames As New ADODB.Recordset
Dim AcctCode As New ADODB.Recordset
Dim rstTemp As New ADODB.Recordset
Dim DrCat As String
Dim CrCat As String
Dim TAbClic As Integer
Dim xCtrlKeyPress2 As Integer       ' account name to account Number.
Sub PrintEntries()
    Dim Jn As String
    Dim ColumnNo As Integer
    Dim TrnType As String
    Dim xcaption As String
    Jn = Me.ListView1.SelectedItem.SubItems(8)
    TrnType = Me.ListView1.SelectedItem.SubItems(9)
    ColumnNo = Me.ListView1.ColumnHeaders(8).Index - 1
    PutCaption PerGEnjounralEntry.Sections(2).Controls("Label3"), "Company Report", ColumnNo
    ColumnNo = Me.ListView1.ColumnHeaders(10).Index - 1
    PutCaption PerGEnjounralEntry.Sections(2).Controls("Label6"), "Company Report", ColumnNo
    
    xcaption = GenDesc
    PutGenDesc PerGEnjounralEntry.Sections(5).Controls("Label24"), "Company Report", xcaption
    xcaption = cLogUser
    PutGenDesc PerGEnjounralEntry.Sections(5).Controls("Label25"), "Company Report", xcaption
    xcaption = Format(Date, "dd/mm/yyyy")
    PutGenDesc PerGEnjounralEntry.Sections(2).Controls("Label8"), "Company Report", xcaption
    On Error Resume Next
    FinanceDE.rsTempGenJournal.close
    If Trim(WhoProcess) = "" Then
    'FinanceDE.rsTempGenJournal.Open "select * from TempGenJournal where prepby=? order by ticket"
     FinanceDE.TempGenJournal (cLogUser)
     Else
     FinanceDE.TempGenJournal (WhoProcess)
    End If
    PerGEnjounralEntry.Show

'           Printer.Orientation = 1
'           Printer.Print
'           Printer.Print
'           Printer.Print
'           Printer.Print
'
'           Printer.FontSize = 12
'           Printer.FontName = "Times new Roman"
'           Printer.Print ; Tab(8); "Journal No: "; Me.Combo11
'           Printer.Print
'           Printer.FontSize = 10
'           Printer.FontName = "Arabic Transparent"
'           Printer.Print ; Tab(10); Date
'           Printer.Print ; Tab(10); Time
'           Printer.FontName = "Times new Roman"
'           Printer.Print ; Tab(10); "-------------------------------------------------------------------------------------------------------------------------------------------------------"
'           i = 0
'           Dim cr As Currency
'           Dim dR As Currency
'           Dim TotalDb  As Currency
'           Dim TotalCr As Currency
'           For i = 1 To Me.ListView1.ListItems.Count
'                  TN = Me.ListView1.ListItems.Item(i).SubItems(1)
'                  Ac = Me.ListView1.ListItems.Item(i).SubItems(2)
'                  an = Me.ListView1.ListItems.Item(i).SubItems(3)
'                  dR = FormatNumber(Me.ListView1.ListItems.Item(i).SubItems(5), 2, vbTrue, vbTrue, vbTrue)
'                  cr = FormatNumber(Me.ListView1.ListItems.Item(i).SubItems(6), 2, vbTrue, vbTrue, vbTrue)
'                  Jn = Me.ListView1.ListItems.Item(i).SubItems(7)
'                  descr = Me.ListView1.ListItems.Item(i).SubItems(8)
'                  TRantYpe = Me.ListView1.ListItems.Item(i).SubItems(9)
'                  Class = Trim(Me.ListView1.ListItems.Item(i).SubItems(10))
'                  TotalDb = TotalDb + dR
'                  TotalCr = TotalCr + cr
'
'                  Printer.Print ; Tab(10); "Classification :" & Right(Class, 60)
'                  Printer.Print ; Tab(10); an; Tab(50); Ac _
'                              ; Tab(115 - Len(dR)); IIf(dR <> 0, Format(dR, "###,###,###.#0"), "") _
'                              ; Tab(130 - Len(cr)); IIf(cr <> 0, Format(cr, "###,###,###.#0"), "")
'                  Printer.FontName = "Arabic Transparent"
'                  Printer.Print ; Tab(10); "Desc : "; TRantYpe & " " & descr
'                  Printer.FontName = "Times new Roman"
'                  If i = Me.ListView1.ListItems.Count Then
'                    Printer.Print ; Tab(10); "-------------------------------------------------------------------------------------------------------------------------------------------------------"
'                   Else
'                   Printer.Print ""
'                  End If
'           Next i
'
'           Printer.Print ; Tab(10); "Totals: " & (i - 1); Tab(115 - Len(TotalDb)); Format(TotalDb, "###,###,###.#0"); Tab(130 - Len(TotalCr)); Format(TotalCr, "###,###,###.#0")
'           Printer.Print ; Tab(10); "======================================================================================="
'           Printer.Print
'           Printer.Print
'           Printer.Print ; Tab(10); "Prepared by: "; cLogUser
'           Printer.EndDoc
End Sub
Private Sub PutCaption(lblX As RptLabel, caption As String, i As Integer)
   With lblX
      .CanGrow = True
      .caption = Me.ListView1.SelectedItem.SubItems(i)
   End With
End Sub
Private Sub PutGenDesc(lblX As RptLabel, caption As String, xcaption As String)
   With lblX
      .CanGrow = True
      .caption = xcaption
   End With
End Sub
 
 
 Sub DeleteItems()
 xmsg = MsgBox("Delete TN ÇáÛÇÁ Êä" & Me.ListView1.SelectedItem & "?", vbQuestion + vbYesNo + vbDefaultButton2, "Please confirm ãä ÝÖáß ÇáÊÇßíÏ")
        If xmsg = vbYes Then
            TN = Me.ListView1.SelectedItem.SubItems(1)
            Dim xtemp As New ADODB.Recordset
            xtemp.Open "delete TempGenJournal  where ticket=" & "'" & TN & "'", constring, adOpenKeyset, adLockOptimistic, adCmdText
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
            'xTemp.Open "SElect count(ticket) as [tn] from GenJOurnalTrans where transdate=" & "'" & Xdate & "'", conString, adOpenDynamic, adLockOptimistic, adCmdText
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
                    
                    'were going to delete all the items in the list one by
                    'one. thats why we place the value of each item into memory variables
                    Ac = Me.ListView1.ListItems.Item(i).SubItems(2)
                    an = Me.ListView1.ListItems.Item(i).SubItems(3)
                    dr = Me.ListView1.ListItems.Item(i).SubItems(5)
                    cR = Me.ListView1.ListItems.Item(i).SubItems(6)
                    Jn = Me.ListView1.ListItems.Item(i).SubItems(7)
                    descr = Me.ListView1.ListItems.Item(i).SubItems(8)
                    transtype = Me.ListView1.ListItems.Item(i).SubItems(9)
                    cat = Me.ListView1.ListItems.Item(i).SubItems(10)
                    xTN = Me.ListView1.ListItems.Item(i).SubItems(1)
                    NextTn = TN
                    'cindex = Val(Me.ListView1.ListItems.Item(i).Index)
                    
                    'Me.ListView1.ListItems.Remove cindex
                    xtemp.Open "delete TempGenJournal  where ticket=" & "'" & xTN & "'", constring, adOpenKeyset, adLockOptimistic, adCmdText
                    With rstTemp
                            .addnew
                            !ticket = NextTn
                            !accountnumber = Ac
                            !accountname = an
                            !TRansDate = Format(Me.Combo2, "mm/dd/yyyy")
                            !DebitAmount = dr
                            !creditamount = cR
                            !fcdebit = fcy
                            !FcCredit = fcy
                            !SerialNo = Jn
                            !Description = descr
                            !deletemark = 0
                            !Status = "Unposted"
                            !transtype = transtype
                            !Classification = cat
                            .Update
                      End With
                    Next
                        'add again but with the new TN
                       Me.ListView1.ListItems.clear
                       rstTemp.close
                       rstTemp.Open "TempGenJournal", constring, adOpenKeyset, adLockOptimistic, adCmdTable
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
                         MItem.SubItems(10) = rstTemp!Classification
                         MItem.SubItems(9) = rstTemp!transtype
                         rstTemp.MoveNext
                        Wend
                    
                 Else
                 
                 xtemp.Open "delete TempGenJournal  where ticket=" & "'" & xTN & "'", constring, adOpenKeyset, adLockOptimistic, adCmdText
                End If
                    ListView1.SortKey = 1
                    ListView1.Sorted = True
          End If
          Exit Sub
 End Sub
 Sub DisplayTransToday()
 Me.ListView2.ListItems.clear
 Dim TRnDate As Date
 tDate = Format(Date, "dd/mm/yyyy")
 Dim rstTRansToday As New ADODB.Recordset
 rstTRansToday.Open "Select * from GenJournalTrans where remarks is null order by serialno", CON1, adOpenKeyset, adLockOptimistic, adCmdText
       If rstTRansToday.EOF = False Then
        rstTRansToday.MoveFirst
       End If
        While rstTRansToday.EOF = False
          TRnDate = rstTRansToday!TRansDate
          Set MItem = Me.ListView2.ListItems.Add(, , rstTRansToday!ticket, , IIf(TRnDate = Date, 1, IIf(TRnDate + 1 = Date, 2, IIf(TRnDate + 3 = Date, 3, 4))))
          MItem.SubItems(1) = rstTRansToday!ticket
          MItem.SubItems(2) = rstTRansToday!accountnumber
          MItem.SubItems(3) = rstTRansToday!accountname
          MItem.SubItems(4) = rstTRansToday!TRansDate
          MItem.SubItems(5) = Format(rstTRansToday!DebitAmount, "###,###,###.#0")
          MItem.SubItems(6) = Format(rstTRansToday!creditamount, "###,###,###.#0")
          MItem.SubItems(7) = rstTRansToday!SerialNo
          MItem.SubItems(8) = rstTRansToday!Description
          MItem.SubItems(10) = rstTRansToday!Classification
          MItem.SubItems(9) = rstTRansToday!transtype
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
    If UCase(Trim(xname)) = UCase(Trim(acctnames!accountnamearab)) Then
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
    Me.Combo7.AddItem acctnames!AccountCode
    acctnames.MoveNext
  Wend
 End If
 
Dim catName As String
Dim Prevcap As String
Dim acctNo As String
acctNo = Trim(xname)
Prevcap = Trim(Me.caption)
Call DisplayCatsName(Prevcap, acctNo, catName)
Me.caption = "General Journal Entry   ÅÏÎÇá ÞíæÏ ÇáíæãíÉ  " & "//" & catName
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
       Me.MaskEdBox1.SetFocus
End If
End Sub

Private Sub Combo12_LostFocus()
'Call Combo12_Click
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
    If UCase(Trim(xname)) = Trim(UCase(acctnames!accountnameeng)) Then
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
       Me.Combo13.AddItem acctnames!accountnameeng
       Me.Combo6.AddItem acctnames!AccountCode
       acctnames.MoveNext
    Wend
 End If
 
Dim catName As String
Dim Prevcap As String
Dim acctNo As String
acctNo = Trim(xname)
Prevcap = Trim(Me.caption)
Call DisplayCatsName(Prevcap, acctNo, catName)
Me.caption = "General Journal Entry   ÅÏÎÇá ÞíæÏ ÇáíæãíÉ  " & "//" & catName
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
    Me.MaskEdBox2.SetFocus
End If
End Sub

Private Sub Combo13_LostFocus()
'Call Combo7_LostFocus
End Sub

Private Sub Combo16_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If Me.Combo6 <> "" Or Me.Combo7 <> "" Then
    xmsg = MsgBox("Are Entries Okay? åá ÇÏÎÇáß ÕÍíÍ ", vbQuestion + vbOKCancel, "Please confirm ãä ÝÖáß ÇáÊÇßíÏ")
    If xmsg = vbOK Then
        If Me.Combo16 = "" Then
            msg = MsgBox("Please select Transaction Type", vbExclamation + vbOKOnly, "Message")
            Me.Combo16.SetFocus
            Exit Sub
        End If
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
           
            xtemp.Open "delete tempGenJournal  where ticket=" & "'" & strFindMe & "'", constring, adOpenKeyset, adLockOptimistic, adCmdText
            AddEntries
        End If
        Me.Combo6.SetFocus
    
      End If
    End If
 End If

End Sub

Private Sub Combo2_LostFocus()
If Mid(Me.Combo2, 3, 1) <> "/" Then
  mess = MsgBox("Incorrect Date format", vbInformation + vbOKOnly, "Message")
  Me.Combo2.SetFocus
  Exit Sub
End If

If Mid(Me.Combo2, 6, 1) <> "/" Then
  mess = MsgBox("Incorrect Date format", vbInformation + vbOKOnly, "Message")
  Me.Combo2.SetFocus
  Exit Sub
End If

End Sub


Private Sub Combo3_GotFocus()
Me.ListView1.SetFocus
End Sub

Private Sub Combo6_Click()
'we don't allow the user to input same account number
If Me.Combo6 <> "" And Me.Combo7 = "" Then
    If Me.Combo6 = Me.Combo7 Then
     xmsg = MsgBox("You have entered same Account number in Debit Side" & vbCrLf & _
                    "ÇäÊ ããßä ÇÏÎÇá ÍÓÇÈÇÊ ãÔÈÉ ÈÌÇäÈ ÇáãÏíä ", vbInformation + vbOKOnly, "Message ÑÓÇáÉ ")
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
Me.caption = "General Journal Entry   ÅÏÎÇá ÞíæÏ ÇáíæãíÉ  " & "//" & catName
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

Private Sub Combo6_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
    SaveEntries
End If
End Sub

Private Sub Combo6_LostFocus()
Dim catName As String
Dim Prevcap As String
Dim acctNo As String
acctNo = Trim(Me.Combo6)
Prevcap = Trim(Me.caption)
Call DisplayCats(Prevcap, acctNo, catName)
Me.caption = "General Journal Entry   ÅÏÎÇá ÞíæÏ ÇáíæãíÉ  " & "//" & catName
DrCat = catName


'we don't allow the user to input same account number
If Me.Combo6 <> "" And Me.Combo7 <> "" Then
    If Me.Combo6 = Me.Combo7 Then
     xmsg = MsgBox("You have entered same Account number in Debit Side" & vbCrLf & _
                     "ÇäÊ ããßä ÇÏÎÇá ÍÓÇÈÇÊ ãÔÈÉ ÈÌÇäÈ ÇáãÏíä ", vbInformation + vbOKOnly, "Message ÑÓÇáÉ ")
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
    xmsg = MsgBox("Account Number does not exist! ÇÑÞÇã ÇáÍÓÇÈÇÊ áÇ íæÌÏ", vbInformation + vbOKOnly, "Message ÑÓÇáÉ")
    Me.Combo6.SetFocus
    Exit Sub
Else
  Me.Combo12 = Trim(RstBA!accountnameeng) & "\" & Trim(RstBA!accountnamearab) ' '& " :ÑÓíÈáÑÞã ÇáÍÓÇÈ"
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


Private Sub Combo7_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
    SaveEntries
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
Me.caption = "General Journal Entry   ÅÏÎÇá ÞíæÏ ÇáíæãíÉ  " & "//" & catName
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
Dim catName As String
Dim Prevcap As String
Dim acctNo As String
acctNo = Trim(Me.Combo7)
Prevcap = Trim(Me.caption)
Call DisplayCats(Prevcap, acctNo, catName)
Me.caption = "General Journal Entry   ÅÏÎÇá ÞíæÏ ÇáíæãíÉ  " & "//" & catName
CrCat = catName


'we don't allow the user to input same account number
If Me.Combo6 <> "" And Me.Combo7 <> "" Then
    If Me.Combo6 = Me.Combo7 Then
      xmsg = MsgBox("You have entered same Account number in Debit Side" & vbCrLf & _
                    "ÇäÊ ããßä ÇÏÎÇá ÍÓÇÈÇÊ ãÔÈÉ ÈÌÇäÈ ÇáãÏíä ", vbInformation + vbOKOnly, "Message ÑÓÇáÉ")
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
  xmsg = MsgBox("Account Code does not exist ÇßæÇÏ ÇáÍÓÇÈÇÊ áÇ ÊæÌÏ", vbInformation + vbOKOnly, "Message ÑÓÇáÉ ")
    Me.Combo7.SetFocus
    Exit Sub
Else
  Me.Combo13 = RstBA!accountnameeng & "\" & RTrim(RstBA!accountnamearab)
End If
If Me.Combo7 <> Me.Text5.Text Then
    Me.Text4 = "Text"
End If


End Sub

Private Sub Combo8_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Me.Combo16.SetFocus
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
        
        trandate = Me.ListView2.SelectedItem.SubItems(4)
        TRansDate = Mid(trandate, 4, 2) & "/" & Left(trandate, 2) & "/" & Mid(trandate, 7, 4)
        Me.Combo2 = TRansDate 'put the transdate selected
        Sn = Me.ListView2.SelectedItem.SubItems(7)
        Me.Combo11 = Sn
        'Check TEmptable if i is empty, if not empty we can't edit the a data
        rstTemp.Requery
        If rstTemp.RecordCount <> 0 Then
            mess = MsgBox("There was an unsaved entries that you must first save before editing.", vbInformation + vbOKOnly, "Message")
            Exit Sub
        End If
        If UserRole = "Admin" Then
           rsttran.Open "Select * from GenJournalTrans where serialno=" & "'" & Trim(Sn) & "'", constring, adOpenDynamic, adLockOptimistic, adCmdText
          Else
          rsttran.Open "Select * from GenJournalTrans where Prepby=" & "'" & cLogUser & "'" & "and serialno=" & "'" & Trim(Sn) & "'", constring, adOpenDynamic, adLockOptimistic, adCmdText
          If rsttran.EOF = True Then Exit Sub
        End If
        
        Me.ListView1.ColumnHeaders(2).Alignment = lvwColumnRight
        While rsttran.EOF = False
         If rsttran!TRansDate = trandate Then
            On Error Resume Next
            Set MItem = Me.ListView1.ListItems.Add(, , rsttran!ticket, , 1)
            MItem.SubItems(1) = IIf(Len(rsttran!ticket) = 1, "  " _
                                & (rsttran!ticket), IIf(Len(rsttran!ticket) = 2, _
                                " " & rsttran!ticket, rsttran!ticket))
            MItem.SubItems(2) = rsttran!accountnumber
            MItem.SubItems(3) = rsttran!accountname
            MItem.SubItems(4) = rsttran!TRansDate
            MItem.SubItems(5) = rsttran!DebitAmount
            MItem.SubItems(6) = rsttran!creditamount
            MItem.SubItems(7) = rsttran!SerialNo
            MItem.SubItems(8) = rsttran!Description
            MItem.SubItems(10) = rsttran!Classification
            MItem.SubItems(9) = rsttran!transtype
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
                !transtype = rsttran!transtype
                !deletemark = 0
                !Status = "Unposted"
                !Prepby = rsttran!Prepby
                WhoProcess = rsttran!Prepby
                GenDesc = IIf(IsNull(rsttran!GenDesc) = True, "", rsttran!GenDesc)
                .Update
              End With
            rsttran.Delete
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
'            Me.Combo6 = Me.ListView2.SelectedItem.SubItems(2)
'            Me.Combo12 = Me.ListView2.SelectedItem.SubItems(3)
'            Me.MaskEdBox1.Text = Me.ListView2.SelectedItem.SubItems(5)
'            Me.Combo11 = Me.ListView2.SelectedItem.SubItems(7)
'            Me.Combo14 = Me.ListView2.SelectedItem.SubItems(1)
             Me.Text4.Text = Trim(Me.ListView2.SelectedItem.SubItems(1))
          Else
'            Me.Combo7 = Me.ListView2.SelectedItem.SubItems(2)
'            Me.Combo13 = Me.ListView2.SelectedItem.SubItems(3)
'            Me.MaskEdBox2.Text = Me.ListView2.SelectedItem.SubItems(6)
'            Me.Combo11 = Me.ListView2.SelectedItem.SubItems(7)
'            Me.Combo14 = Me.ListView2.SelectedItem.SubItems(1)
            Me.Text4.Text = Trim(Me.ListView2.SelectedItem.SubItems(1))
        End If
         DisplayTransToday ' refesh the of listview2
         Me.SSTab1.SetFocus
         SendKeys "{Right}"
    End If
End Sub

Private Sub FI_Click()

End Sub

Private Sub Form_Activate()
'Open all transactions today
 DisplayTransToday
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
  If AcctCode!Active <> 0 Then
   Me.Combo6.AddItem AcctCode!AccountCode
   Me.Combo7.AddItem AcctCode!AccountCode
   Me.Combo12.AddItem AcctCode!accountnamearab & "\" & RTrim(AcctCode!accountnameeng)
   Me.Combo13.AddItem AcctCode!accountnameeng & "\" & RTrim(AcctCode!accountnamearab)
  End If
   AcctCode.MoveNext
  Wend
 AcctCode.close

 
 If Me.Text2.Text <> "Edit" Then
        Set RstBA = New ADODB.Recordset

        'Check Temp table if it is empty or not. if not put it on the list view
         On Error Resume Next
         rstTemp.close
         rstTemp.Open "Select * from TempGenJournal where Prepby=" & "'" & cLogUser & "'", CON1, adOpenKeyset, adLockOptimistic, adCmdText
         On Error GoTo 0
         If rstTemp.EOF = False Then
            rstTemp.MoveFirst
            Me.Combo11 = rstTemp!SerialNo
          Else
               'open setuptable to get journal no.
               Dim JOurnalNo As New ADODB.Recordset
               JOurnalNo.Open "Setup", constring, adOpenKeyset, adLockPessimistic, adCmdTable
               If Val(Left(JOurnalNo!CurrentMoYr, 2)) <> Format(Date, "mm") Then
                   JOurnalNo!CurrentMoYr = Format(Date, "mm") & Trim(Right(Year(Date), 2))
                   JOurnalNo!nextjn = "00001"
                   JOurnalNo.Update
                   Me.Combo11 = JOurnalNo!nextjn
                   JOurnalNo.close
                Else
                   Jn = "GEN" & "-" & Right(Year(Date), 2) & Trim(Format(Date, "mm")) & "-" & Right(JOurnalNo!nextjn, 5)
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
            MItem.SubItems(10) = rstTemp!Classification
            MItem.SubItems(9) = rstTemp!transtype
            rstTemp.MoveNext
         Wend
         NextTn = Me.ListView1.ListItems.Count + 1
         Me.Combo14 = NextTn
         On Error GoTo 0

 End If
 
'OPen TransType table
Dim rstTransType As New ADODB.Recordset
rstTransType.Open "Select * from GenjournalTransType order by code", constring, adOpenKeyset, adLockPessimistic, adCmdText
Do Until rstTransType.EOF
    Me.Combo16.AddItem rstTransType!Code & "-" & rstTransType!NameEng
    rstTransType.MoveNext
Loop
rstTransType.close


Nelson:
c = Err.Number
If c = 3705 Then
  rstTemp.close
  rstTemp.Open "TempGenJournal", CON1, adOpenKeyset, adLockOptimistic, adCmdTable
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
If Me.Text2.Text <> "Edit" Then

  If Me.ListView1.ListItems.Count <> 0 Then
    If Trim(Me.SSTab1.caption) = "List" Then
      On Error Resume Next
      Me.SSTab1.SetFocus
      SendKeys "{Right}"
     End If
        msg = MsgBox("Do you want to keep unsaved entries ? åá ÇäÊ ÊÑíÏ ÇáÈíÇäÇÊ áÇÊÍÝÙ", vbQuestion + vbYesNoCancel + vbDefaultButton1, "Please confirm ãä ÝÖáß ÇáÊÇßíÏ")
        If msg = vbYes Then
            pressno = 0
            rstTemp.close
            Unload Me
         ElseIf msg = vbNo Then
          xmsg = MsgBox("All your entries will be discarded if you Click Yes button åá ÊÑíÏ ÇáÛÇÁ ßá ÇáÈíÇäÇÊ ÇÐÇ ãÊÇßÏ ÇÖÛØ ãæÇÝÞ ÇÓÝá " & vbCrLf & _
                "Do you want to discard? åá ÊÑíÏ ÇáÛÇÁ ", vbQuestion + vbYesNo, "Please confirm ãä ÝÖáß ÇáÊÇßíÏ")
          If xmsg = vbYes Then
            Dim DelTran As New ADODB.Recordset
            If Trim(WhoProcess) = "" Then
              DelTran.Open "Delete TempGenJournal where prepby=" & "'" & cLogUser & "'", constring, adOpenDynamic, adLockPessimistic, adcmdttext
             Else
              DelTran.Open "Delete TempGenJournal where prepby=" & "'" & WhoProcess & "'", constring, adOpenDynamic, adLockPessimistic, adcmdttext
            End If
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
    rsttran.Open "GenJournalTrans", constring, adOpenDynamic, adLockOptimistic, adCmdTable
    rstTemp.close
    If Trim(WhoProcess) = "" Then
       rstTemp.Open "Select * from TempGenJournal where prepby=" & "'" & cLogUser & "'", constring, adOpenDynamic, adLockOptimistic, adCmdText
     Else
      rstTemp.Open "Select * from TempGenJournal where prepby=" & "'" & WhoProcess & "'" & "and serialno =" & "'" & Trim(Me.Combo11) & "'", constring, adOpenDynamic, adLockOptimistic, adCmdText
    End If
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
            !deletemark = 0
            !Status = "Unposted"
            !Classification = rstTemp!Classification
            !transtype = rstTemp!transtype
            !Prepby = rstTemp!Prepby
            !GenDesc = GenDesc
            .Update
          End With
        rstTemp.MoveNext
    Wend
    rstTemp.close
    c = Err.Description
      If Trim(WhoProcess) = "" Then
          DelTran.Open "Delete TempGenJournal where prepby=" & "'" & cLogUser & "'", constring, adOpenDynamic, adLockPessimistic, adcmdttext
         Else
          DelTran.Open "Delete TempGenJournal where prepby=" & "'" & WhoProcess & "'" & "and serialno=" & "'" & Trim(Me.Combo11) & "'", constring, adOpenDynamic, adLockPessimistic, adcmdttext
      End If
    WhoProcess = ""
End If
End Sub

Private Sub Fv_Click()
If Me.FV.caption = "&Normal View" Then
  Me.FV.caption = "&Full View"
   GenJournalEntry.Frame5.Top = 4920
   GenJournalEntry.ListView1.Top = 5160
   GenJournalEntry.ListView1.Height = 1935
   GenJournalEntry.Frame5.Height = 2295
  Else
   Me.FV.caption = "&Normal View"
   GenJournalEntry.Frame5.Top = 350
   GenJournalEntry.ListView1.Top = 630
   GenJournalEntry.Frame5.Height = 7600
   GenJournalEntry.ListView1.Height = 7200
End If

End Sub

Private Sub Li_Click()
Me.ListView2.View = lvwIcon
Me.Li.Checked = True
Me.SI.Checked = False
Me.xlist.Checked = False
Me.xDetails.Checked = False
End Sub

Private Sub ListView1_DblClick()
If Me.ListView1.Top = 630 Then
   'Scmenu.Fv.caption = "&Full View"
   GenJournalEntry.Frame5.Top = 4920
   GenJournalEntry.ListView1.Top = 5160
   GenJournalEntry.ListView1.Height = 1935
   GenJournalEntry.Frame5.Height = 2295
  Else
   'Scmenu.Fv.caption = "&Normal View"
   GenJournalEntry.Frame5.Top = 350
   GenJournalEntry.ListView1.Top = 630
   GenJournalEntry.Frame5.Height = 7600
   GenJournalEntry.ListView1.Height = 7200
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
    Me.Combo16 = Me.ListView1.SelectedItem.SubItems(9)
    DrCat = Me.ListView1.SelectedItem.SubItems(10)
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
    Me.Combo16 = Me.ListView1.SelectedItem.SubItems(9)
    CrCat = Me.ListView1.SelectedItem.SubItems(10)
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
If KeyCode = 113 Then
    SaveEntries
End If

End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = 2 Then
   PopupMenu Main2
End If
End Sub

Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim TOtdR As Currency
Dim TOTcr As Currency
Dim Dramt As Currency
Dim CrAmt As Currency
i = 0
For i = 1 To Me.ListView1.ListItems.Count
    Dramt = Me.ListView1.ListItems.Item(i).SubItems(5)
    CrAmt = Me.ListView1.ListItems.Item(i).SubItems(6)
    TOtdR = TOtdR + Dramt
    TOTcr = TOTcr + CrAmt
Next
Me.ListView1.ToolTipText = "Total Debit=" & Format(TOtdR, "###,###,###.#0") & " and " & "Total Credit=" & Format(TOTcr, "###,###,###.#0")
End Sub

Private Sub ListView2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Me.ListView2.SortKey = ColumnHeader.Index - 1
Me.ListView2.Sorted = True
End Sub

Private Sub ListView2_DblClick()
'Check the temptable if it is empty. if not abort.
If Me.ListView2.ListItems.Count <> 0 Then
 Dim rsttran As ADODB.Recordset
 Set rsttran = New ADODB.Recordset
If Trim(WhoProcess) = "" Then
   rsttran.Open "Select * from TempGenJOurnal where Prepby=" & "'" & cLogUser & "'", constring, adOpenDynamic, adLockOptimistic, adCmdText
  Else
  rsttran.Open "Select * from TempGenJOurnal where Prepby=" & "'" & WhoProcess & "'", constring, adOpenDynamic, adLockOptimistic, adCmdText
 End If
 If rsttran.EOF = False Then
    xmsg = MsgBox("Please save first your unsaved transactions before Deleting or Modifying any transactions that already saved. ", vbInformation + vbOKOnly, "Message")
     Exit Sub
 End If
 rsttran.close
 Me.Text2.Text = "Edit"
 Call eEdit_Click
 Me.SSTab1.SetFocus
 SendKeys "{Right}"
SendKeys "{Right}"
End If

End Sub

Private Sub ListView2_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = 2 Then
  
    If Me.ListView2.ListItems.Count = 0 Then
        Me.xFind.Enabled = False
        Me.eEdit.Enabled = False
        'Me.xRefresh.Enabled = False
        Me.xPreview.Enabled = False
        Me.POstToGL.Enabled = False
       Else
        Me.xFind.Enabled = True
        Me.eEdit.Enabled = True
        'Me.xRefresh.Enabled = True
        Me.xPreview.Enabled = True
        Me.POstToGL.Enabled = True
    End If
    PopupMenu Me.main
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
If Me.ListView2.ListItems.Count <> 0 Then
    Dim rsInvj As New ADODB.Recordset
    Dim MItem As ListItem
    mess = MsgBox("Do you want to continue? åá ÊÑíÏ ÇáÇÓÊãÑÇÑ", vbQuestion + vbOKCancel + vbDefaultButton2, "Please confirm ãä ÝÖáß ÇáÊÇßíÏ")
    If mess = vbOK Then
         PostingJournal.Text1.Text = "GEN"
         rsInvj.Open "SELECT  TransDate, COUNT(TransDate) AS TotalTRan, SUM(DebitAmount) AS DrAmt, SUM(CreditAmount) AS CrAmt" _
         & " From GEnJournalTrans where remarks is null GROUP BY TransDate ORDER BY TransDate", constring, adOpenKeyset, adLockPessimistic, adCmdText
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
  Else
  mess = MsgBox("No transactions to post ", vbInformation + vbOKOnly, "Message")
End If
End Sub

Private Sub PrintItem_Click()
PrintEntries
End Sub

Private Sub SI_Click()
Me.ListView2.View = lvwSmallIcon
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
 rstPn.Open "SElect count(ticket) as [tn1] from GenJournalTRans where transdate=" & "'" & Xdate & "'" & "and remarks is null", constring, adOpenDynamic, adLockOptimistic, adCmdText
 TN = rstPn!TN1  '+ 1
 Me.Combo1.Text = TN ' IIf(tn = 0, 1, tn)
 rstPn.close
  
 'put total credit in combo4
 rstPn.Open "Select SUm(CreditAmount)  as [TotalCr] from GEnJournalTrans where TransDate= " & "'" & Xdate & "'" & "and remarks is null", constring, adOpenDynamic, adLockOptimistic, adCmdText
 Me.Combo5 = IIf(rstPn!TotalCr <> 0, Format(rstPn!TotalCr, "###,###,###.#0"), "0.00")
 Me.Combo5 = Format(Me.Combo5, "###,###,###.#0")
 rstPn.close
 
'put totals debit in combo5
 rstPn.Open "Select SUm(debitAmount)  as [TotalDb] from GenJOurnaltrans where TransDate= " & "'" & Xdate & "'" & "and remarks is null", constring, adOpenDynamic, adLockOptimistic, adCmdText
 Me.Combo4 = IIf(rstPn!TotalDb <> 0, Format(rstPn!TotalDb, "###,###,###.#0"), "0.00")
 Me.Combo4 = Format(Me.Combo4, "###,###,###.#0")
 rstPn.close
 
 'put totals Trans in combo1
 rstPn.Open "Select count(Ticket)  as TotTrans from GEnJournalTrans where TransDate= " & "'" & Xdate & "'" & "and remarks is null", constring, adOpenDynamic, adLockOptimistic, adCmdText
 Me.Combo1 = IIf(rstPn!TotTrans <> 0, rstPn!TotTrans, 0)
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
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim xtemp As ADODB.Recordset
Set xtemp = New ADODB.Recordset
Select Case Button.Key
    Case Is = "Add"
        Call Combo6_LostFocus
        Call Combo7_LostFocus
        If Me.Combo16 = "" Then
            msg = MsgBox("Please select Transaction Type", vbExclamation + vbOKOnly, "Message")
            Me.Combo16.SetFocus
            Exit Sub
        End If

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
            xtemp.Open "delete TempGenJOurnal  where ticket=" & "'" & strFindMe & "'", constring, adOpenKeyset, adLockOptimistic, adCmdText
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
             CancelGenDesc = True
             GenDescription.Show 1
             If CancelGenDesc <> True Then
               SaveEntries
              Else
               Exit Sub
             End If
             DisplayTransToday
    '--------------------------------------------------------------
    Case Is = "Print"
         Me.ListView1.SetFocus
         CancelGenDesc = True
         GenDescription.Show 1
        If CancelGenDesc <> True Then
          PrintEntries
         Else
         Exit Sub
        End If
    '--------------------------------------------------------------
    Case Is = "Return"
       
       Unload Me
End Select
End Sub
Sub SaveEntries()
          X = 0
          Dim ccTotCr As Currency
          Dim ccTotDr As Currency
          Dim xCr As Currency
          Dim xDr As Currency
          'we don't allowed the user to save their entries
          'if the debit is not equal to the credit side
          trDate = Me.Combo2
          TRansDate = Mid(trDate, 4, 2) & "/" & Left(trDate, 2) & "/" & Mid(trDate, 7, 4)
          For X = 1 To Me.ListView1.ListItems.Count
                  xDr = Me.ListView1.ListItems.Item(X).SubItems(5)
                  xCr = Me.ListView1.ListItems.Item(X).SubItems(6)
                  ccTotDr = ccTotDr + xDr
                  ccTotCr = ccTotCr + xCr
          Next
          If ccTotCr <> ccTotDr Then
            xmsg = MsgBox("Total Debit and Total Credit is not equal, Please equalized first before saving" & vbCrLf & _
                        "ÇÌãÇáí Çáãíä æÇÌãÇáí ÇáÏÇÆä áÇ íÓÇæí ,ãä ÝÖáß ÊÇáíÏ ÇæáÇð ÞÈá ÇáÍÝÙ ", vbInformation + vbOKOnly, "Message ÑÓÇáÉ")
            Exit Sub
          End If
          xmsg = MsgBox(" Are you sure you want to save now?åá ÇäÊ ãÊÇßíÏ ãä ÇáÍÝÙ ÇáÇä  ", vbQuestion + vbOKCancel, "Please Confirm ãä ÝÖáß ÇáÊÇßíÏ")
          If xmsg = vbOK Then
              Dim RstBA As ADODB.Recordset
              Set RstBA = New ADODB.Recordset
                
             'increment Journal Entry Number
             Dim JOurnalNo As New ADODB.Recordset
             JOurnalNo.Open "setup", constring, adOpenKeyset, adLockOptimistic, adCmdTable
             Jn = "GEN" & "-" & Right(Year(Date), 2) & Trim(Format(Date, "mm")) & "-" & JOurnalNo!nextjn
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
                   If Len((nextjn)) = 1 Then
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
                   Jn = "GEN" & "-" & Right(Year(Date), 2) & Trim(Format(Date, "mm")) & "-" & JOurnalNo!nextjn
                   Me.Combo11 = Jn
                   JOurnalNo.close
                End If
             End If
             Me.Combo14 = 1
             
             
             'save entries only to GenJOuranlTrans table
              RstBA.Open "GenJOurnalTrans", constring, adOpenDynamic, adLockOptimistic, adCmdTable
              i = 0
              For i = 1 To Me.ListView1.ListItems.Count
                  TN = Me.ListView1.ListItems.Item(i).SubItems(1)
                  Ac = Me.ListView1.ListItems.Item(i).SubItems(2)
                  an = Me.ListView1.ListItems.Item(i).SubItems(3)
                  dr = Me.ListView1.ListItems.Item(i).SubItems(5)
                  cR = Me.ListView1.ListItems.Item(i).SubItems(6)
                  Jn = Me.ListView1.ListItems.Item(i).SubItems(7)
                  descr = Me.ListView1.ListItems.Item(i).SubItems(8)
                  transtype = Me.ListView1.ListItems.Item(i).SubItems(9)
                  cat = Me.ListView1.ListItems.Item(i).SubItems(10)
                  
                  If dr <> "" Then
                     With RstBA
                           .addnew
                           !SerialNo = Jn
                           !ticket = TN
                           !deletemark = 0
                           !accountnumber = Ac
                           !accountname = an
                           !Description = descr & " (Processed Last:" & Format(Date, "dd/mm/yy") & ")"
                           !DebitAmount = dr
                           !creditamount = cR
                           'TRansDate = Mid(Me.Combo2, 4, 2) & "/" & Left(Me.Combo2, 2) & "/" & Mid(Me.Combo2, 7, 4)
                           !TRansDate = TRansDate
                           !Status = "Unposted"
                           !deletemark = 0
                           !Status = "Unposted"
                           !Classification = cat
                           !transtype = transtype
                           If Trim(WhoProcess) = "" Then
                           !Prepby = cLogUser 'Trim(Mid(Mainform.sbStatusBar.Panels(4).Text, 6, 10))
                             Else
                             !Prepby = Trim(WhoProcess)
                           End If
                           !GenDesc = GenDesc
                           .Update
                      End With
                   End If
               Next
               GenDesc = ""
               
              'Make Temptable empty
               Dim xtemp As New ADODB.Recordset
               If Trim(WhoProcess) = "" Then
                 xtemp.Open "delete TEmpGenJournal where prepby = " & " '" & cLogUser & "'", constring, adOpenKeyset, adLockOptimistic, adCmdText
                Else
                xtemp.Open "delete TEmpGenJournal where prepby = " & " '" & WhoProcess & "'" & "and serialno=" & "'" & Trim(Jn) & "'", constring, adOpenKeyset, adLockOptimistic, adCmdText
               End If
               Me.ListView1.ListItems.clear
             
             End If
             WhoProcess = ""
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
             Me.Combo6.SetFocus
             End If
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
    xmsg = MsgBox("Please enter Transaction Amount ãä ÝÖáß ÇÏÎÇá ãÌãæÚ ÇáÕÝÞÇÊ", vbExclamation + vbOKOnly, "Message ÑÓÇáÉ")
    Exit Sub
End If
If Me.MaskEdBox2.Text = "" And Me.Combo7 <> "" Then
    xmsg = MsgBox("Please enter Transaction Amount ãä ÝÖáß ÇÏÎÇá ãÌãæÚ ÇáÕÝÞÇÊ", vbExclamation + vbOKOnly, "Message ÑÓÇáÉ")
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
 RstBA.Open "SElect count(Ticket) as [tn] from GenJournalTrans where Transdate=" & "'" & Date & "'", constring, adOpenDynamic, adLockOptimistic, adCmdText
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
trDate = Me.Combo2
TRansDate = Mid(trDate, 4, 2) & "/" & Left(trDate, 2) & "/" & Mid(trDate, 7, 4)
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
    MItem.SubItems(10) = DrCat
    MItem.SubItems(9) = Me.Combo16
    
    'save also to temp table
    With rstTemp
        .addnew
        !ticket = IIf(TN = 0, NextTn, TN)
        !accountnumber = Me.Combo6
        !accountname = Me.Combo12
        !TRansDate = TRansDate
        !DebitAmount = Me.MaskEdBox1.Text
        !creditamount = 0
        On Error Resume Next
        !SerialNo = Me.Combo11
        !Description = Me.Combo8 & "(Processed Last:" & Format(Date, "dd/mm/yy") & ")" & Me.Combo16
        !deletemark = 0
        !Status = "Unposted"
        !Classification = DrCat
        !transtype = Me.Combo16
        If Trim(WhoProcess) = "" Then
         !Prepby = cLogUser 'Trim(Mid(Mainform.sbStatusBar.Panels(4).Text, 6, 10))
         Else
         !Prepby = Trim(WhoProcess)
        End If
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
       MItem.SubItems(10) = CrCat
       MItem.SubItems(9) = Me.Combo16
       'save also to temp table
       With rstTemp
        .addnew
        !ticket = NextTn
        !accountnumber = Me.Combo7
        !accountname = Me.Combo13
        !TRansDate = TRansDate
        On Error Resume Next
        !DebitAmount = Me.MaskEdBox1.Text
        !creditamount = Me.MaskEdBox2.Text
        !DebitAmount = 0
        !SerialNo = Me.Combo11
        !Description = Me.Combo8 & "(Processed Last:" & Format(Date, "dd/mm/yy") & ")"
        !deletemark = 0
        !Status = "Unposted"
        !Classification = CrCat
        !transtype = Me.Combo16
        If Trim(WhoProcess) = "" Then
         !Prepby = cLogUser 'Trim(Mid(Mainform.sbStatusBar.Panels(4).Text, 6, 10))
         Else
         !Prepby = Trim(WhoProcess)
        End If
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
 Me.Combo6 = ""
 Me.Combo12 = ""
 Me.MaskEdBox1.Text = ""
 Me.Combo8 = ""
 Me.Combo6.SetFocus
 Me.Text4.Text = ""
End Sub

Private Sub xDetails_Click()
Me.ListView2.View = lvwReport
Me.Li.Checked = False
Me.SI.Checked = False
Me.xlist.Checked = False
Me.xDetails.Checked = True

End Sub

Private Sub xFind_Click()
FindItemLV.Show 1
End Sub

Private Sub xList_Click()
Me.ListView2.View = lvwList
Me.Li.Checked = False
Me.SI.Checked = False
Me.xlist.Checked = True
Me.xDetails.Checked = False

End Sub

Private Sub xNEwjn_Click()
Me.Combo6 = ""
Me.Combo7 = ""
Me.Combo12 = ""
Me.Combo13 = ""
Me.MaskEdBox1 = ""
Me.MaskEdBox2 = ""
Me.Combo8 = ""
'Me.ListView1.ListItems.Clear
Me.SSTab1.SetFocus
SendKeys "{Right}"

    
End Sub

Private Sub xPreview_Click()

Dim todayDate As Date
todayDate = Me.ListView2.SelectedItem.SubItems(4)

Dim rsTRans As New ADODB.Recordset
rsTRans.Open "select Count(*) as GTotal,sum(Debitamount) as DRTotal,Sum(CreditaMount) as CrTotal from genjournaltrans where transdate=" & "'" & todayDate & "'" & " and Remarks is null", constring, adOpenKeyset, adLockPessimistic, adCmdText
If rsTRans.EOF = False Then
    gTotal = rsTRans!gTotal
    dRTotal = FormatNumber(rsTRans!dRTotal, 2, vbTrue, vbTrue, vbTrue)
    cRTotal = FormatNumber(rsTRans!cRTotal, 2, vbTrue, vbTrue, vbTrue)
    GenJournalTotalTrn = gTotal
End If
Dim xcaption As String
xcaption = "As of  " & Format(todayDate, "dd mmmm yyyy")
PutGTotal byGenDescrition.Sections(2).Controls("Label13"), "Company Report", xcaption

xcaption = gTotal
PutGTotal byGenDescrition.Sections(7).Controls("Label16"), "Company Report", xcaption
xcaption = dRTotal
PutGTotal byGenDescrition.Sections(7).Controls("Label17"), "Company Report", xcaption
xcaption = cRTotal
PutGTotal byGenDescrition.Sections(7).Controls("Label18"), "Company Report", xcaption
xcaption = cLogUser
PutGTotal byGenDescrition.Sections(7).Controls("Label27"), "Company Report", xcaption
On Error Resume Next
FinanceDE.rsByGenDescription_Grouping.close
FinanceDE.ByGenDescription_Grouping todayDate ', cLogUser
byGenDescrition.Show 1

End Sub
Private Sub PutGTotal(lblX As RptLabel, caption As String, xcaption As String)
   With lblX
      .CanGrow = True
      .caption = xcaption
   End With
End Sub
Private Sub xREfresh_Click()
DisplayTransToday
End Sub

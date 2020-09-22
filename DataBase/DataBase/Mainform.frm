VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Mainform 
   Caption         =   "Finance System  "
   ClientHeight    =   6300
   ClientLeft      =   735
   ClientTop       =   1755
   ClientWidth     =   8880
   BeginProperty Font 
      Name            =   "Arabic Transparent"
      Size            =   9.75
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Mainform.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6300
   ScaleWidth      =   8880
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   3600
      Top             =   5160
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   3000
      Top             =   5160
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9720
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   480
      Top             =   5400
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
            Picture         =   "Mainform.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mainform.frx":0894
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mainform.frx":0CE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mainform.frx":1138
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mainform.frx":158A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   3360
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mainform.frx":19DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mainform.frx":1E2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mainform.frx":2280
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mainform.frx":26D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mainform.frx":282C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mainform.frx":2B46
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mainform.frx":3088
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mainform.frx":34DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mainform.frx":392C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mainform.frx":3D7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mainform.frx":3ED8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   810
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   1429
      ButtonWidth     =   1508
      ButtonHeight    =   1376
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList3"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Log-in"
            Key             =   "login"
            Object.ToolTipText     =   "Log in User"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Trans "
            Key             =   "Transact"
            Object.ToolTipText     =   "Taking up Transactions"
            ImageIndex      =   4
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   14
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "xGJ"
                  Text            =   "General Journal Setup... ÇÚÏÇÏ ÇáíæãíÇÊ ÇáÚÇãÉ "
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "xAJ"
                  Text            =   "Asset Setup...ÇÚÏÇÏ ÇáÇÕæá "
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "xSJ"
                  Text            =   "Sales... ÇáãÈíÚÇÊ "
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "xIJ"
                  Text            =   "Inventory...ÇáãÎÒæä "
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "BT"
                  Text            =   "Bank Transaction..."
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "bar1"
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "xPettyCash"
                  Text            =   "Petty Cash..."
               EndProperty
               BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "xPayables"
                  Text            =   "Payables.. ÇáãÏÝæÚÇÊ "
               EndProperty
               BeginProperty ButtonMenu9 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "bar2"
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu10 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "xCAshRec"
                  Text            =   "Cash Receipt... ÇáäÞÏíÉ ÇáãÓÊáãÉ "
               EndProperty
               BeginProperty ButtonMenu11 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "xcashpay"
                  Text            =   "Cash Payments... ÇáäÞÏíÉ ÇáãÏÝæÚÉ "
               EndProperty
               BeginProperty ButtonMenu12 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "bar3"
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu13 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "debitnote"
                  Text            =   "Debit Note... ãáÇÍÙÇÊ ÇáãÏíä"
               EndProperty
               BeginProperty ButtonMenu14 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "creditnote"
                  Text            =   "Credit Note... ãáÇÍÙÇÊ ÇáÏÇÆä "
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Post áÕÞ "
            Key             =   "Post"
            Object.ToolTipText     =   "Posting of Journals"
            ImageIndex      =   1
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   13
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "xGenJourPost"
                  Text            =   "General Journal íæãíÉ ÚÇãÉ "
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "xAssetJourPost"
                  Text            =   "Asset Journal íæãíÉ ÇáÇÕæá "
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "xSalesJOurPost"
                  Text            =   "Sales Journal íæãíÉ ÇáãÈíÚÇÊ "
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "xInventoryJOurPost"
                  Text            =   "Inventory Journal íæãíÉ ÇáãÎÒæä "
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "bar1"
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "xPayablesJourPost"
                  Text            =   "Payables Journal íæãíÉ ÇáãÏÝæÚÇÊ "
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "xPurchaseJourpost"
                  Text            =   "Purchase Journal..."
               EndProperty
               BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "bar2"
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu9 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "xCashRecJOurpost"
                  Text            =   "Cash Receipt... ÇáäÞÏíÉ ÇáãÓÊáãÉ "
               EndProperty
               BeginProperty ButtonMenu10 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "CashPayJourPost"
                  Text            =   "Cash Payments... ÇáäÞÏíÉ ÇáãÏÝæÚÉ "
               EndProperty
               BeginProperty ButtonMenu11 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "bar4"
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu12 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "xdebitnote"
                  Text            =   "Debit Note... áÇÍÙÇÊ ÇáãÏíä"
               EndProperty
               BeginProperty ButtonMenu13 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "xcreditnote"
                  Text            =   "Credit Note... ãáÇÍÙÇÊ ÇáÏÇÆä "
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Print ØÈÚÉ"
            Key             =   "Print"
            Object.ToolTipText     =   "Printing of Reports"
            ImageIndex      =   6
            Style           =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Query"
            Key             =   "Sql"
            Object.ToolTipText     =   "Querying records"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Charts"
            Key             =   "Charts"
            Object.ToolTipText     =   "Chart of Accounts"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Graph"
            Key             =   "Graph"
            Object.ToolTipText     =   "Viewing Income Statements in Graph"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Find Accts"
            Key             =   "FindAccts"
            Object.ToolTipText     =   "Find Account Names"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
      EndProperty
      MousePointer    =   99
      MouseIcon       =   "Mainform.frx":41F2
   End
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4155
      Left            =   8760
      ScaleHeight     =   1809.265
      ScaleMode       =   0  'User
      ScaleWidth      =   780
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   72
   End
   Begin MSComctlLib.ListView lvListView 
      Height          =   4080
      Left            =   4200
      TabIndex        =   1
      Top             =   840
      Width           =   4530
      _ExtentX        =   7990
      _ExtentY        =   7197
      View            =   1
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      Icons           =   "ImageList2"
      SmallIcons      =   "ImageList2"
      ForeColor       =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "Mainform.frx":4354
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "TransactionDate"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Total Tickets"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Total Debit Amt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Total Credit Amt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Status"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Remarks"
         Object.Width           =   2434
      EndProperty
   End
   Begin MSComctlLib.TreeView tvTreeView 
      Height          =   4080
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   4170
      _ExtentX        =   7355
      _ExtentY        =   7197
      _Version        =   393217
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "Mainform.frx":44B6
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2400
      Top             =   5160
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   -120
      Top             =   6240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mainform.frx":4618
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mainform.frx":4A6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mainform.frx":4EBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mainform.frx":530E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mainform.frx":5760
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mainform.frx":5BB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mainform.frx":6004
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mainform.frx":6456
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   1200
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   390
      Left            =   0
      TabIndex        =   0
      Top             =   5910
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   688
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7461
            Text            =   "Status: Not Ready"
            TextSave        =   "Status: Not Ready"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "12/05/2004"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "02:46 ã"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "User:"
            TextSave        =   "User:"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   1560
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   20
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mainform.frx":68A8
            Key             =   "Back"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mainform.frx":69BA
            Key             =   "Forward"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mainform.frx":6ACC
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mainform.frx":6BDE
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mainform.frx":6CF0
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mainform.frx":6E02
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mainform.frx":6F14
            Key             =   "Properties"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mainform.frx":7026
            Key             =   "View Large Icons"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mainform.frx":7138
            Key             =   "View Small Icons"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mainform.frx":724A
            Key             =   "View List"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mainform.frx":735C
            Key             =   "View Details"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mainform.frx":746E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mainform.frx":78C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mainform.frx":7D12
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mainform.frx":8164
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mainform.frx":85B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mainform.frx":8A08
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mainform.frx":8E5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mainform.frx":92AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Mainform.frx":96FE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image imgSplitter 
      Height          =   4185
      Left            =   4080
      MousePointer    =   9  'Size W E
      Top             =   840
      Width           =   150
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File ãáÝ"
      Begin VB.Menu Login 
         Caption         =   "Log-in...           ÊÓÌíá ÇáÏÎæá "
         Shortcut        =   ^I
      End
      Begin VB.Menu Logou 
         Caption         =   "Log-out... ÊÓÌíá ÇáÎÑæÌ "
         Enabled         =   0   'False
         Shortcut        =   ^T
      End
      Begin VB.Menu bar8 
         Caption         =   "-"
      End
      Begin VB.Menu new 
         Caption         =   "&New                   ÌÏíÏ"
         Enabled         =   0   'False
         Begin VB.Menu xuser 
            Caption         =   "User... ÇáãÓÊÎÏã "
         End
         Begin VB.Menu bar16 
            Caption         =   "-"
         End
         Begin VB.Menu NewAcct 
            Caption         =   "Chart of Accounts...Ïáíá  ÇáÍÓÇÈÇÊ "
         End
         Begin VB.Menu BankAccount 
            Caption         =   "Bank Account...             ÍÓÇÈ ÇáÈäß"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
         Begin VB.Menu Payee 
            Caption         =   "Payee...                                   ÇáãÏÝæÚ áå"
         End
         Begin VB.Menu currency 
            Caption         =   "Currency ÇáÚãáÉ "
         End
         Begin VB.Menu bar17 
            Caption         =   "-"
         End
         Begin VB.Menu taxdetails 
            Caption         =   "Tax Details              ÇáÈíÇäÇÊ ÇáÖÑíÈÉ "
            Visible         =   0   'False
         End
         Begin VB.Menu PMTCat 
            Caption         =   "Payment Category...ÊÕäíÝ ÇáãÏÝæÚÇÊ  "
         End
         Begin VB.Menu assigningcheque 
            Caption         =   "Assign Cheque... ÇáÔíßÇÊ ÇáãÓÌáÉ"
         End
         Begin VB.Menu bar15 
            Caption         =   "-"
         End
         Begin VB.Menu FinanceBuget 
            Caption         =   "Financial Budget... ÇáãæÇÒäÉ ÇáãÇáíÉ"
         End
         Begin VB.Menu RegisternewAsset 
            Caption         =   "Register Asset...ÊÓÌíá ÇáÇÕæá "
         End
      End
      Begin VB.Menu Opentables 
         Caption         =   "Open                     ÝÊÍ "
         Enabled         =   0   'False
         Shortcut        =   ^O
      End
      Begin VB.Menu closeTables 
         Caption         =   "&Close                  ÛáÞ "
      End
      Begin VB.Menu xREfreshLV 
         Caption         =   "Refresh             ÊÍÏíË"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileBar5 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "E&xit                    ÎÑæÌ"
      End
   End
   Begin VB.Menu xTransaction 
      Caption         =   "&Transactions ÇáÚãáíÇÊ "
      Enabled         =   0   'False
      Begin VB.Menu xGEnJOurn 
         Caption         =   "&General Journal...                       íæãíÉ ÚÇãÉ "
      End
      Begin VB.Menu xFixedASset 
         Caption         =   "&Fixed Assets...                           ÇáÇÕæá ÇáËÇÈÊÉ "
      End
      Begin VB.Menu xSAles 
         Caption         =   "Sales...                                            ÇáãÈíÚÇÊ"
      End
      Begin VB.Menu xInventory 
         Caption         =   "Inventory...                                     ÇáãÎÒæä"
      End
      Begin VB.Menu BankTrans 
         Caption         =   "Bank Transaction...                  ÚãáíÇÊ ÇáÈäß"
      End
      Begin VB.Menu bar9 
         Caption         =   "-"
      End
      Begin VB.Menu xPaySetup 
         Caption         =   "&Payable Setup..                            ÇÌÑÇÁÇÊ ÇáãÏÝæÚÇÊ "
      End
      Begin VB.Menu xPurchaseJOurn 
         Caption         =   "Petty Cash...                                        ÇáÎÒíäÉ  "
      End
      Begin VB.Menu xPurchaseSEtup 
         Caption         =   "Purchase Setup                             ÇÌÑÇÁÇÊ ÇáãÔÊÑíÇÊ "
      End
      Begin VB.Menu bar10 
         Caption         =   "-"
      End
      Begin VB.Menu xcAShRct 
         Caption         =   "Receipt voucher  ÓäÏ ÇáÞÈÖ                     "
      End
      Begin VB.Menu xCashPmt 
         Caption         =   " Payment voucher ÓäÏ ÇáÕÑÝ  "
      End
      Begin VB.Menu dashfive 
         Caption         =   "-"
      End
      Begin VB.Menu debitnote 
         Caption         =   "Debit Note...                  ÇÔÚÇÑ ãÏíä "
      End
      Begin VB.Menu creditnote 
         Caption         =   "Credit Note...                    ÇÔÚÇÑ ÏÇÆä  "
      End
   End
   Begin VB.Menu xTools 
      Caption         =   "Too&ls ÇÏæÇÊ "
      Begin VB.Menu xViewUser 
         Caption         =   "View Users...                         ÚÑÖ ÇáãÓÊÎÏãíä "
      End
      Begin VB.Menu xAcctGrouping 
         Caption         =   "Accounts Grouping...    ÊÞÓíã ÇáÍÓÇÈÇÊ         "
         Enabled         =   0   'False
      End
      Begin VB.Menu xfindaccounts 
         Caption         =   "Find Accounts...                     ÇáÈÍË Úä  ÇáÍÓÇÈÇÊ "
      End
      Begin VB.Menu bar13 
         Caption         =   "-"
      End
      Begin VB.Menu backUp 
         Caption         =   "Backup Database                äÓÎ ÇÍÊíÇØí ááÈíÇäÇÊ "
      End
   End
   Begin VB.Menu xReports 
      Caption         =   "&Reports ÇáÊÞÇÑíÑ"
      Enabled         =   0   'False
      Begin VB.Menu Journals 
         Caption         =   "Journals                                                                    íæãíÇÊ "
         Visible         =   0   'False
         Begin VB.Menu AssetsJournal 
            Caption         =   "Assets Journal                                   íæãíÉ ÇáÇÕæá"
         End
         Begin VB.Menu Sales 
            Caption         =   "Sales Journal                                    íæãíÉ ÇáãÈíÚÇÊ "
         End
         Begin VB.Menu GEnJOurnal 
            Caption         =   "General Journal                                     íæãíÉ ÚÇãÉ "
         End
         Begin VB.Menu INYJournal 
            Caption         =   "Inventory Journal                             íæãíÉ ÇáãÎÒæä  "
         End
         Begin VB.Menu CPR 
            Caption         =   "Receipts Voucher Journal                 íæãíÉ ÇáãÞÈæÖÇÊ ÇáäÞÏíÉ "
         End
         Begin VB.Menu cashpaymentjournal 
            Caption         =   "Payment Voucher Journal      íæãíÉ ÇáãÏÝæÚÇÊ ÇáäÞÏíÉ "
         End
         Begin VB.Menu creditnotejournal 
            Caption         =   "Credit Note Journal íæãíÉ ÇæÑÇÞ ÇáÏÝÚ "
         End
         Begin VB.Menu PAYJ 
            Caption         =   "Payables Journal                           íæãíÉ ÇáãÏÝæÚÇÊ "
         End
         Begin VB.Menu PURJ 
            Caption         =   "Purchase Journal                          íæãíÉ ÇáãÔÊÑíÇÊ "
         End
      End
      Begin VB.Menu CashPosition 
         Caption         =   "Cash Position   ÇáãæÞÝ ÇáãÇáí"
         Enabled         =   0   'False
      End
      Begin VB.Menu global 
         Caption         =   "Global Position æÖÚ ÇáÚÇã ááÎÒíäÉ"
      End
      Begin VB.Menu allcheckcollection 
         Caption         =   "Check Collection  ãÊÍÕáÇÊ ÇáÔíßÇÊ"
      End
      Begin VB.Menu MAturingChkPmtRCT 
         Caption         =   "Uncollected Receipt ChequeÇáÔíßÇÊ ÇáÛíÑ ãÍÕáÉ  "
      End
      Begin VB.Menu paymentcheckcollecion 
         Caption         =   "Uncollected Payments Check       ÔíßÇÊ ãÓÊÍÞÉ ÇáÏÝÚ "
      End
      Begin VB.Menu CashCheckColl 
         Caption         =   "Meturing Check Collection  æÖÚ ÇáÔíßÇÊ ÇáãÓÊáãÉ "
      End
      Begin VB.Menu paymentcheck 
         Caption         =   "Maturing Cheque Payment(h) æÖÚ ÇáÔíßÇÊ ÇáãÏÝæÚÉ "
      End
      Begin VB.Menu dahs 
         Caption         =   "-"
      End
      Begin VB.Menu cashagaints 
         Caption         =   "Cash Receipt Against Invoice       ÇáäÞÏíÉ ÇáãÓÊáãÉ Úä ÝæÇÊíÑ "
      End
      Begin VB.Menu payinvoice 
         Caption         =   "Cash Payment Against Invoice  ÇáäÞÏíÉ ÇáãÏÝæÚÉ Úä ÇáÝæÇÊíÑ "
      End
      Begin VB.Menu glbalance 
         Caption         =   "General Ledger Balances  ÇÑÕÏÉ ÇáÇÓÊÇÐ ÇáÚÇã           "
      End
      Begin VB.Menu trialbalance 
         Caption         =   "Trial Balance   ãíÒÇä ÇáãÑÇÌÚÉ "
      End
      Begin VB.Menu ProfitLost 
         Caption         =   "Profit & Lost Statement  ÞÇÆãÉ ÇáÇÑÈÇÍ æÇáÎÓÇÆÑ"
      End
      Begin VB.Menu balancesheet 
         Caption         =   "Balance Sheet      ÇáãíÒÇäíÉ ÇáÚãæãíÉ"
      End
      Begin VB.Menu Nath 
         Caption         =   "-"
      End
      Begin VB.Menu accountinquery 
         Caption         =   "Statement Of Accounts           ßÔÝ ÍÓÇÈ "
      End
      Begin VB.Menu clientlink 
         Caption         =   "Client information   ãÚáæãÇÊ ÇáÚãíá "
      End
      Begin VB.Menu DebAge 
         Caption         =   "Debtors Aging of Account     ÊÞÑíÑ ÇÚãÇÑ ÇáÏíæä "
      End
      Begin VB.Menu CustLedg 
         Caption         =   "Customer Ledger       ÇÓÊÇÐ ÇáÚãáÇÁ"
      End
      Begin VB.Menu PAyset 
         Caption         =   "Payable Setup    ÊÞÑíÑ ÇáãÏÝæÚÇÊ æÇáãÓÊÍÞÇÊ    "
      End
      Begin VB.Menu paidVouc 
         Caption         =   "Paid Voucher ÊÓäÏÇÊ ÇáÕÑÝ ÇáÕÇÏÑÉ"
      End
      Begin VB.Menu cancVou 
         Caption         =   "Cancelled Voucher                          ÓäÏÇÊ  ÇáãáÛíÉ"
      End
      Begin VB.Menu CrAge 
         Caption         =   "Creditors Aging Of Account      ÊÞÑíÑ ÇÚãÇÑ ÇáÏÇÆäæä"
      End
      Begin VB.Menu Stmtofacc 
         Caption         =   "Statement Of Account              ßÔÝ ÍÓÇÈ "
      End
      Begin VB.Menu bar20 
         Caption         =   "-"
      End
      Begin VB.Menu BAR 
         Caption         =   "Bank Account Report             ÊÞÑíÑ ÍÓÇÈÇÊ ÇáÈäß"
         Begin VB.Menu xBankBalances 
            Caption         =   "Accounts Balances        ÇÑÕÏÉ ÍÓÇÈÇÊ ÇáÈäæß"
         End
         Begin VB.Menu xBankAcctLedger 
            Caption         =   "Accounts Ledger          ÇÓÊÇÐ ÇáÈäæß"
         End
      End
   End
   Begin VB.Menu RepGen 
      Caption         =   "Report Generator"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help ãÓÇÚÏ"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents  ãÍÊæíÇÊ "
      End
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About FINANCE System Úä ÇáäÙÇã ÇáãÇáí "
      End
   End
   Begin VB.Menu Scmenu 
      Caption         =   "SCMenu"
      Visible         =   0   'False
      Begin VB.Menu xPost 
         Caption         =   "Post...                  ÊÑÍíá "
         Begin VB.Menu finalGL 
            Caption         =   "general ledger  ÇáÇÓÊÇÐ ÇáÚÇã"
         End
      End
      Begin VB.Menu xlist 
         Caption         =   "View List...             ÚÑÖ ÞÇÆãÉ"
      End
      Begin VB.Menu xCancelDayJournal 
         Caption         =   "Cancel  Journal  ÇáÛÇÁ Çáíæã"
      End
      Begin VB.Menu xPrint 
         Caption         =   "Print Preview...                ØÈÚÉ"
      End
      Begin VB.Menu xREfresh 
         Caption         =   "Refresh                 ÇÚÇÏÉ ÊÍÏíË"
      End
      Begin VB.Menu bar12 
         Caption         =   "-"
      End
      Begin VB.Menu xView 
         Caption         =   "View                                 ÑÄí "
         Begin VB.Menu xSI 
            Caption         =   "Small Icon ÒÑ ÕÛíÑ "
            Checked         =   -1  'True
         End
         Begin VB.Menu xLI 
            Caption         =   "Large Icon ÒÑ ßÈíÑ "
         End
         Begin VB.Menu xViewList 
            Caption         =   "List ÞÇÆãÉ "
         End
         Begin VB.Menu xDetails 
            Caption         =   "Details áÊÝÇÕíá "
         End
      End
      Begin VB.Menu xBackGround 
         Caption         =   "Background                ÇáÎáÝíÉ "
         Begin VB.Menu xChangeBG 
            Caption         =   "Change Background... ÊÛíÑ ÇáÎáÝíÉ "
         End
         Begin VB.Menu xChangeFC 
            Caption         =   "Change Font Color...ÊÛíÑ áæä ÇáßáãÉ "
         End
         Begin VB.Menu xApperance 
            Caption         =   "Appearance "
            Begin VB.Menu xTopLeft 
               Caption         =   "TopLeft ÞãÉ ÇáÛÑÈ "
            End
            Begin VB.Menu xCenter 
               Caption         =   "Center áÇãÑßÒ "
            End
            Begin VB.Menu xTIle 
               Caption         =   "Tile ÇáÊÝÇÕíá "
               Checked         =   -1  'True
            End
         End
      End
   End
End
Attribute VB_Name = "Mainform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)

Dim mbMoving As Boolean
Dim nodex As Node
Dim rst As ADODB.Recordset
Dim constring As String
Dim MItem As ListItem
Dim xcol As ColumnHeader
Const sglSplitLimit = 540
Dim WhatJOurNCOde As String

Dim recreccheck As New ADODB.Recordset
Dim recpaycheck As New ADODB.Recordset
Dim recpay As New ADODB.Recordset
Dim recop As New ADODB.Recordset
Dim con As New ADODB.Connection

'this is for all variables
Dim openingcash As Currency
Dim openingcheck As Currency
Dim openingcreditcard As Currency

Dim collectioncash As Currency
Dim collectioncheck As Currency
Dim collectioncreditcard As Currency


Dim paymentcash As Currency
Dim paymentcheck11 As Currency
Dim paymentcreditcard As Currency

Dim finalcashtotal As Currency
Dim finalchecktotal As Currency
Dim finalcreditcardtotal As Currency

Public secondvariable As Date
Sub PrintSalesJOurnalPosted(SelectedDate As Date)
            Dim rstSj As New ADODB.Recordset
            Dim rstDlySumSJ As New ADODB.Recordset
            Dim rstJOurcode As New ADODB.Recordset
            Dim rstTOtTRansDly As New ADODB.Recordset
            Dim PrinterReady As Boolean
            Dim cRow As Integer
            
         
          
            rstDlySumSJ.Open "SELECT InvoiceDate,count(TransDate) as TotREc, SUM(TradeRcvble) as TR, SUM(TradeDiscAmt)as TDA , SUM(MgtDiscAmt) as MDA , SUM(GrossSales) as GS, SUM(TranspoCharge)as TC, SUM(NetSales) as NS,  SUM(VAT)as VAT, SUM(SURTaxAmt)as STA From SalesJournal Where transdate=" & "'" & SelectedDate & "'" & " and Remarks is not null " & "GROUP BY Invoicedate", _
                                constring, adOpenKeyset, adLockPessimistic, adCmdText
            
            rstSj.Open "Select * from SalesJOurnal where transdate = " & "'" & SelectedDate & "'" & "and remarks is not null " & " order by INvoiceNo", constring, adOpenKeyset, adLockPessimistic, adCmdText
            Call PrintHeading(PrinterReady)
            If PrinterReady = False Then
                Exit Sub
            End If
            cRow = 13
            cPage = 1
            Dim PtTR, PtTDR, PtTDA, PtMDA, PtGS, PtTC, PtNS, PtVat, PtStr, PtStA          As Currency
            Dim GtTR, GtTDR, GtTDA, GtMDA, GtGS, GtTC, GtNS, GtVat, GtStr, GtStA          As Currency
            i = 0
            'Printer.Orientation = 2
            Do Until rstSj.EOF = True
                  i = i + 1
                 'Take the pagetotal
                 PtTR = PtTR + rstSj!tradercvble ', "###,###,###.#0")
                 PtTDR = PtTDR + IIf(rstSj!TRadedisc = 0, 0, rstSj!TRadedisc) ', "###,###,###.#0"))
                 PtTDA = PtTDA + IIf(rstSj!tradeDiscamt = 0, 0, rstSj!tradeDiscamt) ', , "###,###,###.#0"))
                 PtMDA = PtMDA + IIf(rstSj!MgtDisc = 0, 0, rstSj!MgtDisc)
                 PtGS = PtGS + rstSj!GrossSales ', "###,###,###.#0")
                 PtTC = PtTC + rstSj!transpoCharge ', , "###,###,###.#0")
                 PtNS = PtNS + rstSj!NetSales  ', "###,###,###.#0")
                 PtVat = PtVat + rstSj!vat ', , "###,###,###.#0")
                 PtStA = PtStA + rstSj!SurTaxAmt ', "###,###,###.#0")
                 
                 'Take the GrandTotal
                 GtTR = GtTR + rstSj!tradercvble ', "###,###,###.#0")
                 GtTDR = GtTDR + IIf(rstSj!TRadedisc = 0, 0, rstSj!TRadedisc) ', "###,###,###.#0"))
                 GtTDA = GtTDA + IIf(rstSj!tradeDiscamt = 0, 0, rstSj!tradeDiscamt) ', "###,###,###.#0"))
                 GtMDA = GtMDA + IIf(rstSj!MgtDisc = 0, 0, rstSj!MgtDisc)
                 GtGS = GtGS + rstSj!GrossSales ', "###,###,###.#0")
                 GtTC = GtTC + rstSj!transpoCharge
                 GtNS = GtNS + rstSj!NetSales ', "###,###,###.#0")
                 GtVat = GtVat + rstSj!vat ', "###,###,###.#0")
                 GtStA = GtStA + rstSj!SurTaxAmt ', "###,###,###.#0")
                 
                 TRLen = Len(Format(rstSj!tradercvble, "###,###,###.#0"))
                 TDRLen = Len(Format(rstSj!TRadedisc, "##"))
                 TDALen = Len(Format(rstSj!tradeDiscamt, "###,###,###.#0"))
                 MDRLen = Len(Format(rstSj!MgtDisc, "##"))
                 MDALen = Len(Format(rstSj!MgtDiscAmt, "###,###.#0"))
                 GSLen = Len(Format(rstSj!GrossSales, "###,###,###.#0"))
                 TCLen = Len(Format(rstSj!transpoCharge, "###.#0"))
                 NSLen = Len(Format(rstSj!NetSales, "###,###,###.#0"))
                 VATLen = Len(Format(rstSj!vat, "###,###.#0"))
                 STRLen = Len(Format(rstSj!SurTaxRate, "##"))
                 STALEn = Len(Format(rstSj!SurTaxAmt, "###,###.#0"))
                 cRow = cRow + 1
                 
                 
                 If cRow = 49 Then
                    cRow = 12
                    LenPtTR = Len(Format(PtTR, "###,###,###.#0"))
                    LenPtTDA = Len(Format(PtTDA, "###,###,###.#0"))
                    LenPtMDA = Len(Format(PtMDA, "###,###.#0"))
                    LenPtGS = Len(Format(PtGS, "###,###,###.#0"))
                    LenPtTC = Len(Format(PtTC, "###.#0"))
                    LenPtNS = Len(Format(PtNS, "###,###,###.#0"))
                    LenPtVat = Len(Format(PtVat, "###,###.#0"))
                    LenPtSTA = Len(Format(PtStA, "##"))
                    'Printing page footer
                    Printer.Print ; Tab(0); "-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                    Printer.Print ; Tab(0); "Page Total " _
                   ; Tab(96 - LenPtTR); Format(PtTR, "###,###,###.#0") _
                   ; Tab(114 - LenPtTDA); Format(PtTDA, "###,###,###.#0") _
                   ; Tab(133 - LenPtMDA); Format(PtMDA, "###,###.#0") _
                   ; Tab(147 - LenPtGS); Format(PtGS, "###,###,###.#0") _
                   ; Tab(159 - LenPtTC); Format(PtTC, "###.#0") _
                   ; Tab(173 - LenPtNS); "/" & Format(PtNS, "###,###,###.#0") & "/" _
                   ; Tab(186 - LenPtVat); Format(PtVat, "###,###.#0") _
                   ; Tab(201 - LenPtSTA); Format(PtStA, "###,###.#0")
                    Printer.Print ; Tab(0); "=============================================================================================================================================================================="
                    Printer.Print ; Tab(0); "Page No. " & Trim(cPage)
                    'Reset the pagetotal to 0's
                    PtTR = 0
                    PtTDR = 0
                    PtTDA = 0
                    PtMDA = 0
                    PtGS = 0
                    PtTC = 0
                    PtNS = 0
                    PtVat = 0
                    PtStA = 0
                    cPage = cPage + 1
                    Printer.EndDoc
                    Call PrintHeading(PrinterReady)
                 End If
                 
                'Printing Details
                On Error GoTo 0
                Printer.Print ; Tab(0); Format(rstSj!invoicedate, "dd/mm/yy") _
                              ; Tab(13); (rstSj!invoiceno) _
                              ; Tab(29); (rstSj!ClientCode) _
                              ; Tab(42); Trim(Left(rstSj!accountname, 28)) _
                              ; Tab(77); IIf(IsNull(rstSj!profitcenter) = True, "", rstSj!profitcenter) _
                              ; Tab(96 - TRLen); FormatNumber(rstSj!tradercvble, 2, vbTrue, vbTrue, vbTrue) _
                              ; Tab(102 - TDRLen); IIf(rstSj!TRadedisc = 0, "-", Str(rstSj!TRadedisc)) _
                              ; Tab(114 - TDALen); IIf(rstSj!tradeDiscamt = 0, "", FormatNumber(rstSj!tradeDiscamt, 2, vbTrue, vbTrue, vbTrue)) _
                              ; Tab(120 - MDRLen); IIf(rstSj!MgtDisc = 0, "-", Str(rstSj!MgtDisc)) _
                              ; Tab(134 - MDALen); IIf(rstSj!MgtDiscAmt = 0, "-", FormatNumber(rstSj!MgtDiscAmt, 2, vbTrue, vbTrue, vbTrue)) _
                              ; Tab(148 - GSLen); FormatNumber(rstSj!GrossSales, 2, vbTrue, vbTrue, vbTrue) _
                              ; Tab(159 - TCLen); IIf(rstSj!transpoCharge = 0, "-", FormatNumber(rstSj!transpoCharge, 2, vbTrue, vbTrue, vbTrue)) _
                              ; Tab(173 - NSLen); FormatNumber(rstSj!NetSales, 2, vbTrue, vbTrue, vbTrue) _
                              ; Tab(186 - VATLen); IIf(rstSj!vat = 0, "-", FormatNumber(rstSj!vat, 2, vbTrue, vbTrue, vbTrue)) _
                              ; Tab(193 - STRLen); IIf(rstSj!SurTaxRate = 0, "-", Str(rstSj!SurTaxRate)) _
                              ; Tab(203 - STALEn); IIf(rstSj!SurTaxAmt = 0, "-", FormatNumber(rstSj!SurTaxAmt, 2, vbTrue, vbTrue, vbTrue)) _
                              ; Tab(220); rstSj!accountnumber
                              
                             rstSj.MoveNext
            Loop

'print Report Footer
Printer.Print ; Tab(0); "----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
LenPtTR = Len(Format(PtTR, "###,###,###.#0"))
LenPtTDA = Len(Format(PtTDA, "###,###,###.#0"))
LenPtMDA = Len(Format(PtMDA, "###,###.#0"))
LenPtGS = Len(Format(PtGS, "###,###,###.#0"))
LenPtTC = Len(Format(PtTC, "###.#0"))
LenPtNS = Len(Format(PtNS, "###,###,###.#0"))
LenPtVat = Len(Format(PtVat, "###,###.#0"))
LenPtSTA = Len(Format(PtStA, "##"))
Printer.Print ; Tab(0); "Page Total " _
              ; Tab(96 - LenPtTR); Format(PtTR, "###,###,###.#0") _
              ; Tab(114 - LenPtTDA); Format(PtTDA, "###,###,###.#0") _
              ; Tab(134 - LenPtMDA); Format(PtMDA, "###,###.#0") _
              ; Tab(148 - LenPtGS); Format(PtGS, "###,###,###.#0") _
              ; Tab(159 - LenPtTC); Format(PtTC, "###.#0") _
              ; Tab(173 - LenPtNS); "/" & Format(PtNS, "###,###,###.#0") & "/" _
              ; Tab(186 - LenPtVat); Format(PtVat, "###,###.#0") _
              ; Tab(201 - LenPtSTA); Format(PtStA, "###,###.#0")
Printer.Print ; Tab(0); "=============================================================================================================================================================================="

'Printing of Grand Totals
Printer.FontBold = True
LenGtTR = Len(Format(GtTR, "###,###,###.#0"))
LenGtTDA = Len(Format(GtTDA, "###,###,###.#0"))
LenGtMDA = Len(Format(GtMDA, "###,###.#0"))
LenGtGS = Len(Format(GtGS, "###,###,###.#0"))
LenGtTC = Len(Format(GtTC, "###.#0"))
LenGtNS = Len(Format(GtNS, "###,###,###.#0"))
LenGtVat = Len(Format(GtVat, "###,###.#0"))
LenGtSTA = Len(Format(GtStA, "##"))
Printer.Print ; Tab(0); "Grand Total" _
              ; Tab(87 - LenGtTR); Format(GtTR, "###,###,###.#0") _
              ; Tab(103 - LenGtTDA); Format(GtTDA, "###,###,###.#0") _
              ; Tab(120 - LenGtMDA); Format(GtMDA, "###,###.#0") _
              ; Tab(134 - LenGtGS); Format(GtGS, "###,###,###.#0") _
              ; Tab(143 - LenGtTC); Format(GtTC, "###.#0") _
              ; Tab(156 - LenGtNS); Format(GtNS, "###,###,###.#0") _
              ; Tab(168 - LenGtVat); Format(GtVat, "###,###.#0") _
              ; Tab(179 - LenGtSTA); Format(GtStA, "###,###.#0")
Printer.Print ; Tab(0); "=============================================================================================================================================================================="
Printer.FontBold = False
Printer.Print ; Tab(0); "Page No. " & Trim(cPage)
Printer.EndDoc



Printer.Orientation = 2
Printer.Print ""
Printer.Print ""
Printer.Print ""
Printer.FontBold = True
Printer.Print ; Tab(0); "DAILY SALES JOURNAL by Invoice Date"
Printer.FontBold = False
Printer.Print ; Tab(0); "-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
Printer.Print ; Tab(0); "        "; Tab(14); "       "; Tab(29); "      "; Tab(47); "      "; Tab(78); "       "; Tab(90); "Trade"; Tab(99); "       D   I   S   C   O   U   N   T   S"
Printer.Print ; Tab(0); "        "; Tab(14); "       "; Tab(29); "     "; Tab(47); "No. of "; Tab(78); "      "; Tab(90); "Rcvbl"; Tab(99); "      Trade  "; Tab(115); " Management "; Tab(141); "Gross"; Tab(152); "Transpo"; Tab(164); "Net Sales"; Tab(181); "VAT"; Tab(189); " S U R Tax   "; Tab(210); "  Doc "; Tab(222); "           "
Printer.Print ; Tab(0); "Descriptions"; Tab(14); "       "; Tab(29); "      "; Tab(47); "Trans "; Tab(78); "      "; Tab(90); "     "; Tab(99); "      Amount   "; Tab(118); " Amount   "; Tab(141); "Sales"; Tab(152); "Charges"; Tab(165); "         "; Tab(179); "   "; Tab(189); "  Amount    "; Tab(210); "Stamps"; Tab(222); "           "
Printer.Print ; Tab(0); "-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
Do Until rstDlySumSJ.EOF = True
              Printer.Print ; Tab(0); SelectedDate & " Sales Totals" _
              ; Tab(50 - Len(Format(rstDlySumSJ!Totrec, "###,#0"))); Format(rstDlySumSJ!Totrec, "###,#0") _
              ; Tab(94 - Len(Format(rstDlySumSJ!TR, "###,###,###.#0"))); Format(rstDlySumSJ!TR, "###,###,###.#0") _
              ; Tab(110 - Len(Format(rstDlySumSJ!TDA, "###,###,###.#0"))); Format(rstDlySumSJ!TDA, "###,###,###.#0") _
              ; Tab(127 - Len(Format(rstDlySumSJ!MDA, "###,###.#0"))); Format(rstDlySumSJ!MDA, "###,###.#0") _
              ; Tab(146 - Len(Format(rstDlySumSJ!GS, "###,###,###.#0"))); Format(rstDlySumSJ!GS, "###,###,###.#0") _
              ; Tab(158 - Len(Format(rstDlySumSJ!TC, "###.#0"))); Format(rstDlySumSJ!TC, "###.#0") _
              ; Tab(172 - Len(Format(rstDlySumSJ!NS, "###,###,###.#0"))); Format(rstDlySumSJ!NS, "###,###,###.#0") _
              ; Tab(186 - Len(Format(rstDlySumSJ!vat, "###,###.#0"))); Format(rstDlySumSJ!vat, "###,###.#0") _
              ; Tab(198 - Len(Format(rstDlySumSJ!STA, "###,###.#0"))); Format(rstDlySumSJ!STA, "###,###.#0")
              rstDlySumSJ.MoveNext
Loop
Printer.Print ; Tab(0); "-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
Printer.Print ; Tab(0); "Grand Total"; Tab(49 - Len(i)); (i) _
                  ; Tab(92 - LenGtTR); Format(GtTR, "###,###,###.#0") _
                  ; Tab(110 - LenGtTDA); Format(GtTDA, "###,###,###.#0") _
                  ; Tab(127 - LenGtMDA); Format(GtMDA, "###,###.#0") _
                  ; Tab(146 - LenGtGS); Format(GtGS, "###,###,###.#0") _
                  ; Tab(158 - LenGtTC); Format(GtTC, "###.#0") _
                  ; Tab(172 - LenGtNS); Format(GtNS, "###,###,###.#0") _
                  ; Tab(186 - LenGtVat); Format(GtVat, "###,###.#0") _
                  ; Tab(195 - LenGtSTA); Format(GtStA, "###,###.#0")
                  Printer.Print ; Tab(0); "=============================================================================================================================================================================="


Printer.Print ""
Printer.Print ; Tab(0); "Total Debit Amount           : " & Format(GtTR + GtMDA + GtTDA, "###,###,###.#0"); Tab(90); "___________                                              ____________                                              ___________                       "
Printer.Print ; Tab(0); "Total Credit Amount          : " & Format(GtTR + GtMDA + GtTDA, "###,###,###.#0"); Tab(90); "Prepared by                                                   Checked by                                                  Approved by                     "
Printer.Print ; Tab(0); "=============================================================================================================================================================================="
Printer.Print ; Tab(103); "***End of the Report***"
Printer.Print ""
Printer.Print ""
Printer.Print ""
Printer.Print ""
'Printer.Print "___________                                              ____________                                              ___________                       "
'Printer.Print "Prepared by                                                   Checked by                                                  Approved by                     "



'printing by Profit Center
Dim rsByProfitCenter As New ADODB.Recordset
rsByProfitCenter.Open "SELECT  ProfitCenter, COUNT(ProfitCenter) AS TotalbyProfitCenter, SUM(TradeRcvble) AS TR, SUM(TradeDiscAmt) AS TDA,SUM(MgtDiscAmt) AS MDA," _
             & " SUM(GrossSales) AS GS, SUM(TranspoCharge) AS TC,SUM(NetSales) AS NS,SUM(VAt) AS VAT,SUM(SurTaxAmt) AS STA" _
             & " From SalesJournal where transdate=" & "'" & SelectedDate & "'" & "and Remarks is not null GROUP BY ProfitCenter ORDER BY ProfitCenter", constring, adOpenKeyset, adLockPessimistic, adCmdText

'Printer.Orientation = 2
Printer.Print ""
Printer.Print ""
Printer.Print ""
Printer.FontBold = True
Printer.Print ; Tab(0); "DAILY SALES by Profit Center " & Format(Date, "dd/mm/yyyy") 'Covered Date " & Format(trandate, "dd/mm/yyyy") & "-" & Format(xto, "dd/mm/yyyy")
Printer.FontBold = False
Printer.Print ; Tab(0); "-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
Printer.Print ; Tab(0); "        "; Tab(14); "       "; Tab(29); "      "; Tab(47); "      "; Tab(78); "       "; Tab(90); "Trade"; Tab(99); "       D   I   S   C   O   U   N   T   S"
Printer.Print ; Tab(0); "        "; Tab(14); "       "; Tab(29); "     "; Tab(47); "No. of "; Tab(78); "      "; Tab(90); "Rcvbl"; Tab(99); "      Trade  "; Tab(115); " Management "; Tab(141); "Gross"; Tab(152); "Transpo"; Tab(164); "Net Sales"; Tab(181); "VAT"; Tab(189); " S U R Tax   "; Tab(210); "  Doc "; Tab(222); "           "
Printer.Print ; Tab(0); "Profit Center"; Tab(14); "       "; Tab(29); "      "; Tab(47); "Trans "; Tab(78); "      "; Tab(90); "     "; Tab(99); "      Amount   "; Tab(118); " Amount   "; Tab(141); "Sales"; Tab(152); "Charges"; Tab(165); "         "; Tab(179); "   "; Tab(189); "  Amount    "; Tab(210); "Stamps"; Tab(222); "           "
Printer.Print ; Tab(0); "-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
Do Until rsByProfitCenter.EOF = True
                  Printer.Print ; Tab(0); rsByProfitCenter!profitcenter _
                  ; Tab(50 - Len(Format(rsByProfitCenter!TotalByProfitCenter, "###,#0"))); Format(rsByProfitCenter!TotalByProfitCenter, "###,#0") _
                  ; Tab(94 - Len(Format(rsByProfitCenter!TR, "###,###,###.#0"))); Format(rsByProfitCenter!TR, "###,###,###.#0") _
                  ; Tab(110 - Len(Format(rsByProfitCenter!TDA, "###,###,###.#0"))); Format(rsByProfitCenter!TDA, "###,###,###.#0") _
                  ; Tab(127 - Len(Format(rsByProfitCenter!MDA, "###,###.#0"))); Format(rsByProfitCenter!MDA, "###,###.#0") _
                  ; Tab(146 - Len(Format(rsByProfitCenter!GS, "###,###,###.#0"))); Format(rsByProfitCenter!GS, "###,###,###.#0") _
                  ; Tab(158 - Len(Format(rsByProfitCenter!TC, "###.#0"))); Format(rsByProfitCenter!TC, "###.#0") _
                  ; Tab(173 - Len(Format(rsByProfitCenter!NS, "###,###,###.#0"))); Format(rsByProfitCenter!NS, "###,###,###.#0") _
                  ; Tab(185 - Len(Format(rsByProfitCenter!vat, "###,###.#0"))); Format(rsByProfitCenter!vat, "###,###.#0") _
                  ; Tab(197 - Len(Format(rsByProfitCenter!STA, "###,###.#0"))); Format(rsByProfitCenter!STA, "###,###.#0")
      rsByProfitCenter.MoveNext
    Loop
    Printer.Print ; Tab(0); "-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Printer.Print ; Tab(0); "Grand Total"; Tab(49 - Len(i)); (i) _
                  ; Tab(92 - LenGtTR); Format(GtTR, "###,###,###.#0") _
                  ; Tab(110 - LenGtTDA); Format(GtTDA, "###,###,###.#0") _
                  ; Tab(127 - LenGtMDA); Format(GtMDA, "###,###.#0") _
                  ; Tab(146 - LenGtGS); Format(GtGS, "###,###,###.#0") _
                  ; Tab(158 - LenGtTC); Format(GtTC, "###.#0") _
                  ; Tab(173 - LenGtNS); Format(GtNS, "###,###,###.#0") _
                  ; Tab(185 - LenGtVat); Format(GtVat, "###,###.#0") _
                  ; Tab(195 - LenGtSTA); Format(GtStA, "###,###.#0")
    Printer.Print ; Tab(0); "=============================================================================================================================================================================="

Printer.FontBold = False
Printer.EndDoc
End Sub
Sub PrintHeading(PrinterReady As Boolean)
    On Error GoTo Nelson
    Printer.Orientation = 2
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.FontSize = 10
    Printer.FontName = "Arial Narrow"
    Printer.Print ; Tab(0); "Habitat Furniture & Contract Furnishings"
    Printer.FontBold = True
    Printer.FontSize = 12
    Printer.FontName = "Book Antiqua"
    Printer.Print ; Tab(0); "Sales Journal Registry"
    Printer.FontBold = False
    Printer.FontSize = 9.8
    Printer.FontName = "Arab Transparent"
    Printer.FontItalic = True
    Printer.Print ; Tab(0); Format(Date, "dddd, mmmm dd, yyyy")
    Printer.FontItalic = False
    Printer.FontSize = 8.8
    Printer.FontName = "Arial Narrow"
    Printer.Print ; Tab(0); "-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Printer.Print ; Tab(0); "Invoice "; Tab(14); "Invoice"; Tab(29); "Client"; Tab(42); "Client"; Tab(77); " Profit"; Tab(90); "Trade"; Tab(99); "        D   I   S   C   O   U   N   T   S"
    Printer.Print ; Tab(0); " Date   "; Tab(14); "Number "; Tab(29); "Code "; Tab(42); "Name "; Tab(77); "Center"; Tab(90); "Rcvbl"; Tab(99); "      Trade  "; Tab(120); " Management "; Tab(142); "Gross"; Tab(152); "Transpo"; Tab(164); "Net Sales"; Tab(179); "Sales"; Tab(189); " Addt'l Tax  "; Tab(210); "      "; Tab(222); "     GL    "
    Printer.FontName = "Arial Narrow"
    Printer.Print ; Tab(0); "        "; Tab(14); "       "; Tab(29); "      "; Tab(43); "      "; Tab(78); "      "; Tab(90); "     "; Tab(99); "Rate%   Amount "; Tab(118); "Rate%   Amount"; Tab(142); "Sales"; Tab(152); "Charges"; Tab(165); "         "; Tab(179); "Tax"; Tab(189); "Rate% Amount"; Tab(210); "      "; Tab(222); "Acct Number"
    Printer.Print ; Tab(0); "-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
Nelson:
c = Err.Number
If c = 480 Then
   PrinterReady = False
   msg = MsgBox("Printer not ready", vbExclamation + vbOKOnly, "Message")
  Else
  PrinterReady = True
End If
End Sub


Private Sub anu_Click()
frmcashcollection.Show 1
End Sub

Private Sub accountinquery_Click()
frmStatmentOfAccounts.Show
End Sub

Private Sub allcheckcollection_Click()
On Error Resume Next
dataanu.rscom_allcheckundercollection_Grouping.close
On Error GoTo 0
re_allcheckundercollection.Show
End Sub

Private Sub assigningcheque_Click()
frmassigningcheck.Show
End Sub

Private Sub backUp_Click()
Me.MouseIcon = LoadPicture(APp.Path & "\" & "busy_m.cur")
Me.MousePointer = 99
DoEvents
Dim cnstring As New ADODB.Connection
cnstring.Open constring
Me.sbStatusBar.Panels(1).Text = "Status : Backing up Finance Database..."
cnstring.Execute "BACKUP DATABASE [FINANCE] TO  DISK = N'e:\Finance_Project\FinanceDatabase_DailyBackup' WITH  NOINIT ,  NOUNLOAD ,  NAME = N'FINANCE backup',  NOSKIP ,  STATS = 10,  DESCRIPTION = N'Daily database Backing up',  NOFORMAT"
Me.sbStatusBar.Panels(1).Text = "Status : Ready"
cnstring.close
Me.MousePointer = 0
End Sub

Private Sub balancesheet_Click()
frmbalancesheet.Show
End Sub

Private Sub BankTrans_Click()
BankTransaction1.Show
End Sub

Private Sub cancVou_Click()
DRPcancelled.Show
End Sub

Private Sub cashagaints_Click()
frmreceiptagainsinvoice.Show 1
End Sub

Private Sub CashCheckColl_Click()
frmmaturitingreceiptcheckcollection.Show
End Sub

Private Sub cashpayments_Click()

End Sub

Private Sub cashpaymentjournal_Click()
re_csp_journal.Show
End Sub

Private Sub CashPosition_Click()
 openingcash = 0
 openingcheck = 0
 openingcreditcard = 0

 collectioncash = 0
 collectioncheck = 0
 collectioncreditcard = 0


 paymentcash = 0
 paymentcheck11 = 0
 paymentcreditcard = 0

 finalcashtotal = 0
 finalchecktotal = 0
 finalcreditcardtotal = 0


Dim reccheck As New ADODB.Recordset
Dim con As New ADODB.Connection
con.Open "dsn=finance;uid=sa;"
reccheck.Open "select * from vouchers where okprint <> '1' and receiptdate = '" & Format(DateAdd("d", -1, Date), "mm/dd/yyyy") & "'", con, adOpenKeyset, adLockOptimistic
If reccheck.BOF = False Then
    MsgBox "Please Close Your Cash Position Of Yesterdays", vbInformation, "Cash Position"
    Exit Sub
End If
reccheck.close

Dim recreceived As New ADODB.Recordset
Dim recpayed As New ADODB.Recordset
Dim conpass As New ADODB.Connection

Dim anuclass As New HabitatClass
Dim sqltable As Boolean
Dim constring As String
Dim xtable As String
'this is for opening balnce
recop.Open "Select * from balance", con, adOpenKeyset, adLockOptimistic

openingcash = IIf(IsNull(recop!openingbalance), 0, recop!openingbalance)
openingcheck = IIf(IsNull(recop!openingcheck), 0, recop!openingcheck)
openingcreditcard = IIf(IsNull(recop!openingcreditcard), 0, recop!openingcreditcard)

    prcbeginingcash recashposition.Sections(1).Controls("b1")
    prcbeginingcheck recashposition.Sections(1).Controls("b2")
    prcbeginingcreditcard recashposition.Sections(1).Controls("b3")
'end opening balance

constring = "dsn=finance;uid;sa"
sqltable = True
xtable = "SELECT paymode, svoucher,CASE WHEN LEFT(paymode, 2) = '01' OR LEFT(paymode, 2) = '05' THEN sum(receiptamount * CurrencyRate) END AS realcash, " _
& "CASE WHEN LEFT(paymode, 2) = '03' OR LEFT(paymode, 2) = '04' THEN sum(receiptamount * CurrencyRate) END AS realcheck, " _
& "CASE WHEN LEFT(paymode, 2)= '02' THEN sum(receiptamount * CurrencyRate) END AS realcard " _
& "From vouchers WHERE okprint = '0' and deleted <> '1' AND receiptdate = '" & Format(Date, "mm/dd/yyyy") & "' group by paymode,svoucher"
anuclass.GetTables recreceived, conpass, xtable, constring, sqltable
While recreceived.EOF = False
    If recreceived!svoucher = "Collections" Then
        collectioncash = collectioncash + IIf(IsNull(recreceived!realcash), 0, recreceived!realcash)
        collectioncheck = collectioncheck + IIf(IsNull(recreceived!realcheck), 0, recreceived!realcheck)
        collectioncreditcard = collectioncreditcard + IIf(IsNull(recreceived!realcard), 0, recreceived!realcard)
    Else
        paymentcash = paymentcash + IIf(IsNull(recreceived!realcash), 0, recreceived!realcash)
        paymentcheck11 = paymentcheck11 + IIf(IsNull(recreceived!realcheck), 0, recreceived!realcheck)
        paymentcreditcard = paymentcreditcard + IIf(IsNull(recreceived!realcard), 0, recreceived!realcard)
    End If
    recreceived.MoveNext
Wend
conpass.close

    finalcashtotal = openingcash + collectioncash - paymentcash
    finalchecktotal = openingcheck + collectioncheck - paymentcheck11
    finalcreditcardtotal = openingcreditcard + collectioncreditcard - paymentcreditcard
    
    'sub ending balance after the group printing
     prcendingcash recashposition.Sections(7).Controls("e1")
     prcendingcheck recashposition.Sections(7).Controls("e2")
     prcendingcreditcard recashposition.Sections(7).Controls("e3")
    'end ending balance for the group printing

         FormatLabelAC1 recashposition.Sections(7).Controls("be1")
         FormatLabelAC2 recashposition.Sections(7).Controls("be2")
         FormatLabelAC3 recashposition.Sections(7).Controls("be3")
         FormatLabelAC4 recashposition.Sections(7).Controls("co1")
         FormatLabelAC5 recashposition.Sections(7).Controls("co2")
         FormatLabelAC6 recashposition.Sections(7).Controls("co3")
         FormatLabelAC7 recashposition.Sections(7).Controls("pa1")
         FormatLabelAC8 recashposition.Sections(7).Controls("pa2")
         FormatLabelAC9 recashposition.Sections(7).Controls("pa3")
         FormatLabelAC10 recashposition.Sections(7).Controls("cl1")
         FormatLabelAC11 recashposition.Sections(7).Controls("cl2")
         FormatLabelAC12 recashposition.Sections(7).Controls("cl3")

On Error Resume Next
dataanu.rscom_Cashposition_Grouping.close
On Error GoTo 0

dataanu.com_Cashposition_Grouping Format(Date, "mm/dd/yyyy")
recashposition.Show 1

'this is for check to ask question for closing
reccheck.Open "select * from vouchers where okprint <> '1' and receiptdate = '" & Format(Date, "mm/dd/yyyy") & "'", con, adOpenKeyset, adLockOptimistic
If reccheck.BOF = False Then
    If MsgBox("Are You Sure Want to Close the Cash Position ? åá ÇäÊ ãÊÇßÏ ÊÑíÏ ÇáÛáÞ ßÇÔíÑ  ?", vbYesNo, "Confirm to Close") = vbYes Then
        
            recop.MoveFirst
            With recop
            !balancedate = Date
            !openingbalance = finalcashtotal
            !openingcheck = finalchecktotal
            !openingcreditcard = finalcreditcardtotal
            recop.Update
            End With
                    'update the currency table
                    Dim rectakevoucollection As New ADODB.Recordset
                    Dim rectakevoupayment As New ADODB.Recordset
                    Dim recupdate As New ADODB.Recordset
                    
                    'this is to update beginning
                    recupdate.Open "update currencytable set beginning = ending", con, adOpenKeyset, adLockOptimistic
                    'end update
                    
                    'this is for and collections
                    rectakevoucollection.Open "Select CurrencyMark, SUM(receiptamount) AS voucollections from vouchers where svoucher = 'Collections' and okprint <> '1' and deleted <> '1' and receiptdate =" & "'" & Format(Date, "mm/dd/yyyy") & "' group by currencymark", con, adOpenKeyset, adLockOptimistic
                    If rectakevoucollection.BOF = False Then
                        While rectakevoucollection.EOF = False
                             recupdate.Open "update currencytable set collections=" & Val(rectakevoucollection!voucollections) & "" _
                            & " where currency = '" & Trim(rectakevoucollection!currencymark) & "'", con, adOpenKeyset, adLockOptimistic
                            rectakevoucollection.MoveNext
                        Wend
                    End If
                    rectakevoucollection.close
                    
                    'this is for payments
                    rectakevoupayment.Open "Select CurrencyMark, SUM(receiptamount) AS voucollections from vouchers where svoucher <> 'Collections' and okprint <> '1' and deleted <> '1' and receiptdate =" & "'" & Format(Date, "mm/dd/yyyy") & "' group by currencymark", con, adOpenKeyset, adLockOptimistic
                    If rectakevoupayment.BOF = False Then
                        While rectakevoupayment.EOF = False
                            recupdate.Open "update currencytable set payment=" _
                            & Val(rectakevoupayment!voucollections) & "where currency = '" _
                            & Trim(rectakevoupayment!currencymark) & "'", con, adOpenKeyset, adLockOptimistic
                            rectakevoupayment.MoveNext
                        Wend
                    End If
                    'this is for ending balance
                    recupdate.Open "update currencytable set ending = beginning+collections-payment", con, adOpenKeyset, adLockOptimistic
                    'end update
                    
            Dim reccloseall As New ADODB.Recordset
            Dim rechostname As New ADODB.Recordset
            
            reccloseall.Open "update vouchers set okprint = '1' where okprint <> '1' and deleted <> '1' and receiptdate =" & "'" & Format(Date, "mm/dd/yyyy") & "'", con, adOpenKeyset, adLockOptimistic
            rechostname.Open "Select Host_name() as takehostname", con, adOpenKeyset, adLockOptimistic
            
            takehostname = rechostname!takehostname
            
            On Error Resume Next
            reccloseall.close
            rechostname.close
            On Error GoTo 0
     
            'this is to keep the allopening balance in the allbalance table
            reccloseall.Open "allbalance", con, adOpenKeyset, adLockOptimistic
            
            reccloseall.addnew
                reccloseall!balancedate = Date
                reccloseall!openingbalance = finalcashtotal
                reccloseall!openingcheck = finalchecktotal
                reccloseall!openingcreditcard = finalcreditcardtotal
                reccloseall!HostName = takehostname
                reccloseall!LogUser = cLogUser
            reccloseall.Update
            reccloseall.close
    End If
End If
On Error Resume Next
reccheck.close
recop.close
con.close
conpass.close
On Error GoTo 0
End Sub

Private Sub clientlink_Click()
connectclientcode.Show
End Sub

Private Sub closeTables_Click()
'Dim rstL6 As New ADODB.Recordset
'Dim rstFM As New ADODB.Recordset
'rstFM.Open "financeMaster", conString, adOpenKeyset, adLockPessimistic, adCmdTable
'
'Do Until rstFM.EOF = True
'  an = rstFM!AccountNameEng
'  ac = rstFM!AccountCode
'  rstL6.Open "Update Level6 set AccountNameEng=" & "'" & an & "'" & "Where AccountCode = " & " '" & ac & "'", conString, adOpenKeyset, adLockPessimistic, adCmdText
'  'rstL6.Close
'  rstFM.MoveNext
'Loop

Me.tvTreeView.Nodes.clear
'Me.Opentables.Enabled = True
'SCMenu.xMain.Enabled = False
End Sub

Private Sub CoolBar1_Click()
'W = Me.CoolBar1.Bands(2).Width
'l = Me.CoolBar1.Bands(1).Width
'Me.ProgressBar1.Width = W - 100
'Me.ProgressBar1.Left = l + 150
End Sub

Private Sub CoolBar1_Resize()
'W = Me.CoolBar1.Bands(2).Width
'l = Me.CoolBar1.Bands(1).Width
'Me.ProgressBar1.Width = W - 100
'Me.ProgressBar1.Left = l + 150
End Sub

Private Sub Deposit_Click()
ProcessDep.Show 1
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Command1_KeyPress(KeyAscii As Integer)

End Sub

Private Sub Command1_KeyUp(KeyCode As Integer, Shift As Integer)
X = KeyCode
End Sub

Private Sub CPR_Click()
re_csr_jourmal.Show 1
End Sub

Private Sub CrAge_Click()
AgingCreditors.Show 1
End Sub

Private Sub creditnote_Click()
frmcreditnote.Show
End Sub

Private Sub creditnotejournal_Click()
On Error Resume Next
dataanu.rscom_creditnotejournal.Requery
On Error GoTo 0
re_creditnotejournal.Show
End Sub

Private Sub currency_Click()
frmtempcurrency.Show 1
End Sub

Private Sub DebAge_Click()
CustLedger.Show 1
End Sub

Private Sub debitnote_Click()
frmdebitnote.Show
End Sub

Private Sub exit_Click()
Unload Me

End Sub

Private Sub finalGL_Click()
        Dim rsInvj As New ADODB.Recordset
        Dim MItem As ListItem
        Dim SelectedDate As Date
        mess = MsgBox("Do you want to continue? ", vbQuestion + vbYesNo + vbDefaultButton2, "Please confirm")
        If mess = vbYes Then
            PostingJournal.Text1.Text = WhatJOurNCOde
            If WhatJOurNCOde = "SAL" Then
                xtable = "SalesJournal"
             ElseIf WhatJOurNCOde = "GEN" Then
                xtable = "GEnJournalTRans"
             ElseIf WhatJOurNCOde = "GEN" Then
                xtable = "BankJournal"
             ElseIf WhatJOurNCOde = "IVY" Then
                xtable = "InventoryJournal"
             ElseIf WhatJOurNCOde = "AST" Then
                xtable = "AssetJournal"
             ElseIf WhatJOurNCOde = "PYB" Then
                xtable = "payJournal"
             ElseIf WhatJOurNCOde = "SRL" Then
                xtable = "creditnote"
             ElseIf WhatJOurNCOde = "SPA" Then
                xtable = "debitnote"
            ElseIf WhatJOurNCOde = "SRL" Then
                xtable = "debitnote"
             ElseIf WhatJOurNCOde = "CSR" Or WhatJOurNCOde = "CSP" Then
                xtable = "CashJournal"
                    If WhatJOurNCOde = "CSR" Then
                        anutype = "R"
                    Else
                        anutype = "P"
                    End If
           
             ElseIf WhatJOurNCOde = "PTC" Then
                xtable = "PettyJournal"
            End If
            
            If xtable = "SalesJournal" Then
              rsInvj.Open "SELECT  TransDate, COUNT(TransDate) AS TotalTRan,SUM(tradercvble+tradeDiscAmt) AS DrAmt, SUM(GrossSales+Vat+transpoCharge+surtaxamt) AS CrAmt" _
              & " From " & xtable & " where Remarks is null GROUP BY TransDate ORDER BY TransDate", constring, adOpenKeyset, adLockPessimistic, adCmdText
             ElseIf UCase(xtable) = UCase("PayJournal") Then
               rsInvj.Open "SELECT  confirmedDate, COUNT(confirmeddate) AS TotalTRan,SUM(dbAmount) AS DrAmt, SUM(CrAmount) AS CrAmt" _
               & " From " & xtable & " where Remarks is null GROUP BY confirmedDate ORDER BY Confirmeddate", constring, adOpenKeyset, adLockPessimistic, adCmdText
               
            ElseIf UCase(xtable) = UCase("cashjournal") And anutype = "R" Then
               rsInvj.Open "SELECT  transdate, COUNT(transdate) AS TotalTRan,SUM(debitAmount) AS DrAmt, SUM(CreditAmount) AS CrAmt" _
               & " From " & xtable & " where Remarks is null and trantype = 'R' GROUP BY transdate ORDER BY transdate", constring, adOpenKeyset, adLockPessimistic, adCmdText
               
            ElseIf UCase(xtable) = UCase("cashjournal") And anutype = "P" Then
               rsInvj.Open "SELECT  transdate, COUNT(transdate) AS TotalTRan,SUM(debitAmount) AS DrAmt, SUM(CreditAmount) AS CrAmt" _
               & " From " & xtable & " where Remarks is null and trantype = 'P' GROUP BY transdate ORDER BY transdate", constring, adOpenKeyset, adLockPessimistic, adCmdText

            ElseIf UCase(xtable) = UCase("Pettyjournal") Then
               rsInvj.Open "SELECT  datex, COUNT(datex) AS TotalTRan,SUM(dbAmount) AS DrAmt, SUM(CrAmount) AS CrAmt" _
               & " From " & xtable & " where PostMark is null GROUP by datex ORDER BY datex", constring, adOpenKeyset, adLockPessimistic, adCmdText

             Else
                rsInvj.Open "SELECT  TRansDate, COUNT(TransDate) AS TotalTRan,SUM(debitAmount) AS DrAmt, SUM(CreditAmount) AS CrAmt" _
               & " From " & xtable & " where Remarks is null GROUP BY transDate ORDER BY transdate", constring, adOpenKeyset, adLockPessimistic, adCmdText
            End If
           
            Do Until rsInvj.EOF = True
                If UCase(xtable) = UCase("PayJournal") Then
                   Set MItem = PostingJournal.ListView1.ListItems.Add(, , Format(rsInvj!confirmeddate, "dd/mm/yyyy"))
                 ElseIf UCase(xtable) = UCase("PettyJournal") Then
                   Set MItem = PostingJournal.ListView1.ListItems.Add(, , Format(rsInvj!Datex, "dd/mm/yyyy"))
                 Else
                 Set MItem = PostingJournal.ListView1.ListItems.Add(, , Format(rsInvj!TRansDate, "dd/mm/yyyy"))
                End If
                MItem.SubItems(1) = rsInvj!TotalTRan
                MItem.SubItems(2) = FormatNumber(rsInvj!Dramt, 2, vbTrue, vbTrue, vbTrue)
                MItem.SubItems(3) = FormatNumber(rsInvj!CrAmt, 2, vbTrue, vbTrue, vbTrue)
                MItem.SubItems(4) = "Waiting"
                rsInvj.MoveNext
            Loop
            
            PostingJournal.Show 1
        End If
      End Sub

Private Sub FinanceBuget_Click()
financialBudget.Show
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
'Unload ScreenSaver
End Sub

Private Sub Form_Paint()
    'lvListView.View = Val(GetSetting(App.Title, "Settings", "ViewMode", "3"))
    Select Case lvListView.View
        Case lvwIcon
           ' tbToolBar.Buttons(LISTVIEW_MODE0).Value = tbrPressed
        Case lvwSmallIcon
           ' tbToolBar.Buttons(LISTVIEW_MODE1).Value = tbrPressed
        Case lvwList
           ' tbToolBar.Buttons(LISTVIEW_MODE2).Value = tbrPressed
        Case lvwReport
           ' tbToolBar.Buttons(LISTVIEW_MODE3).Value = tbrPressed
    End Select
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    If Me.Logou.Enabled = True Then
     Call Logou_Click
    End If
    'If Trim(Me.sbStatusBar.Panels(4).Text) <> "User:" Then
    '    mss = MsgBox("To complete close this program Please Click Logout", vbExclamation + vbOKOnly, "Message")
    '    Cancel = -1
    'End If

    'close all sub forms
    For i = Forms.Count - 1 To 1 Step -1
        Unload Forms(i)
    Next
    If Me.WindowState <> vbMinimized Then
        SaveSetting APp.Title, "Settings", "MainLeft", Me.Left
        SaveSetting APp.Title, "Settings", "MainTop", Me.Top
        SaveSetting APp.Title, "Settings", "MainWidth", Me.Width
        SaveSetting APp.Title, "Settings", "MainHeight", Me.Height
    End If
    'SaveSetting App.Title, "Settings", "ViewMode", lvListView.View = lvwReport
    End
End Sub



Private Sub Form_Resize()
'Me.CoolBar1.Width = Me.Width - 100
    On Error Resume Next
    If Me.Width < 3000 Then Me.Width = 3000
    SizeControls imgSplitter.Left
End Sub


Private Sub Gi_Click()
End Sub

Private Sub Ic_Click()
Printcheq.Show 1
End Sub

Private Sub glbalance_Click()
frmglstatement.Show
End Sub

Private Sub global_Click()

Dim xtable As String
Dim sqltable As Boolean
Dim xClass As New HabitatClass
Dim rectemp As New ADODB.Recordset
Dim contemp As New ADODB.Connection

xtable = "Select * from vouchers where deleted = '0' and okprint = '0' and receiptdate = '" & Format(Date, "mm/dd/yyyy") & "'"
constring = "Dsn=finance;uid=sa;"
sqltable = True
xClass.GetTables rectemp, contemp, xtable, constring, sqltable

If rectemp.BOF = False Then
    MsgBox "You Have to Close the Cash Position Before You Print the Global Position", vbInformation, "Global Position"
    rectemp.close
    contemp.close
    Exit Sub
End If
rectemp.close
contemp.close

xtable = "Select * from currencytable where closedate = '" & Format(Date, "mm/dd/yyyy") & "'"
sqltable = True
xClass.GetTables rectemp, contemp, xtable, constring, sqltable

If rectemp.BOF = False Then
    On Error Resume Next
    dataanu.rscom_foreigncurrency_closing_Grouping.close
    On Error GoTo 0
    re_allcashbalance.Show 1
Else
    MsgBox "You Have to Close the Cash Position Before You Print the Global Position", vbInformation, "Global Position"
End If
rectemp.close
contemp.close

End Sub

Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    With imgSplitter
        picSplitter.Move .Left, .Top, .Width \ 2, .Height - 20
    End With
    picSplitter.Visible = True
    mbMoving = True
End Sub


Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    Dim sglPos As Single
    If mbMoving Then
        sglPos = X + imgSplitter.Left
        If sglPos < sglSplitLimit Then
            picSplitter.Left = sglSplitLimit
        ElseIf sglPos > Me.Width - sglSplitLimit Then
            picSplitter.Left = Me.Width - sglSplitLimit
        Else
            picSplitter.Left = sglPos
        End If
    End If
End Sub


Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    SizeControls picSplitter.Left
    picSplitter.Visible = False
    mbMoving = False
End Sub


Private Sub TreeView1_DragDrop(Source As Control, X As Single, y As Single)
    If Source = imgSplitter Then
        SizeControls X
    End If
End Sub


Sub SizeControls(X As Single)
    On Error Resume Next
    

    'set the width
    If X < 1540 Then X = 1540
    If X > (Me.Width - 1540) Then X = Me.Width - 1540
    tvTreeView.Width = X + 10
    imgSplitter.Left = X
    lvListView.Left = X + 40
   ' ListView1.Left = x + 40
    lvListView.Width = Me.Width - (tvTreeView.Width + 140)
   ' Me.ListView1.Width = Me.Width - (tvTreeView.Width + 140)
    'set the top
  

   ' If tbToolBar.Visible Then
        tvTreeView.Top = tbToolBar.Height + picTitles.Height
   ' Else
   '     tvTreeView.Top = picTitles.Height
   ' End If

  lvListView.Top = tvTreeView.Top - 20
  'ListView1.Top = tvTreeView.Top

    'set the height
   ' If sbStatusBar.Visible Then
        tvTreeView.Height = Me.Height - 2020 '1500 ' - (picTitles.Top + picTitles.Height + sbStatusBar.Height)
   ' Else
   '     tvTreeView.Height = Me.ScaleHeight - (picTitles.Top + picTitles.Height)
   ' End If
    

    lvListView.Height = tvTreeView.Height + 20
    'istView1.Height = tvTreeView.Height
    imgSplitter.Top = tvTreeView.Top
    imgSplitter.Height = tvTreeView.Height
End Sub

Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub ListView1_DblClick()
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)

End Sub

Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)

End Sub

Private Sub NI_Click()
ProcessPmt.Show 1
End Sub

Private Sub List1_Click()

End Sub

Private Sub Login_Click()


LogUser.Show 1
If LogSucess = True Then
    EnablesAllObject
    Call Opentables_Click
End If
End Sub
Sub EnablesAllObject()
Me.Login.Enabled = False
Me.Logou.Enabled = True
Me.Toolbar1.Buttons(2).Enabled = False
Me.Toolbar1.Buttons(3).Enabled = True
Me.Toolbar1.Buttons(4).Enabled = True
Mainform.Toolbar1.Buttons(5).Enabled = True
Mainform.Toolbar1.Buttons(6).Enabled = True
Mainform.Toolbar1.Buttons(7).Enabled = True
Mainform.Toolbar1.Buttons(8).Enabled = True
Mainform.Toolbar1.Buttons(9).Enabled = True
Mainform.Toolbar1.Buttons(10).Enabled = True
Mainform.Toolbar1.Buttons(11).Enabled = True
Mainform.Toolbar1.Buttons(12).Enabled = True
Mainform.xReports.Enabled = True
Me.xReports.Enabled = True
Me.Opentables.Enabled = True
Me.new.Enabled = True
Me.xrefresh.Enabled = True
Me.xTransaction.Enabled = True
LogSucess = False
Me.sbStatusBar.Panels(1).Text = "Status: Ready"
End Sub
Private Sub Logou_Click()
Me.Login.Enabled = True
Me.Logou.Enabled = False
Me.Toolbar1.Buttons(2).Enabled = True
Me.Toolbar1.Buttons(3).Enabled = False
Me.Toolbar1.Buttons(4).Enabled = False
Mainform.Toolbar1.Buttons(5).Enabled = False
Mainform.Toolbar1.Buttons(6).Enabled = False
Mainform.Toolbar1.Buttons(7).Enabled = False
Mainform.Toolbar1.Buttons(8).Enabled = False
Mainform.Toolbar1.Buttons(9).Enabled = False
Mainform.Toolbar1.Buttons(10).Enabled = False
Mainform.Toolbar1.Buttons(11).Enabled = False
Mainform.Toolbar1.Buttons(12).Enabled = False
Mainform.CashPosition.Enabled = False
Mainform.xAcctGrouping.Enabled = False
Me.xReports.Enabled = False
Me.Opentables.Enabled = False
Me.new.Enabled = False
Me.xrefresh.Enabled = False
Me.xTransaction.Enabled = False
Me.tvTreeView.Nodes.clear
Me.lvListView.ListItems.clear
LogSucess = False

cUser = Trim(sbStatusBar.Panels(4).Text)
cUser = Trim(Mid(cUser, 6, 20))
rstUser.MoveFirst
Dim rsLOUser As New ADODB.Recordset
rsLOUser.Open "UserS", constring, adOpenKeyset, adLockPessimistic, adCmdTable
Do Until rsLOUser.EOF = True
  If UCase(Trim(rsLOUser!Userid)) = UCase(cUser) Then
    rsLOUser!logouttime = Time
    rsLOUser!logged = "No"
    rsLOUser.Update
    Exit Do
  End If
 rsLOUser.MoveNext
Loop
rsLOUser.close
rstUser.close
Me.sbStatusBar.Panels(4).Text = "User:"

Mainform.xGEnJOurn.Enabled = True
Mainform.xFixedASset.Enabled = True
Mainform.xSAles.Enabled = True
Mainform.xInventory.Enabled = True
Mainform.xPaySetup.Enabled = True
Mainform.xPurchaseJOurn.Enabled = True
Mainform.xPurchaseSEtup.Enabled = True
Mainform.xcAShRct.Enabled = True
Mainform.xCashPmt.Enabled = True
Mainform.NewAcct.Enabled = True
Mainform.RegisternewAsset.Enabled = True

'Under File Menu
Mainform.BankAccount.Enabled = True
Mainform.Payee.Enabled = True
Mainform.taxdetails.Enabled = True
Mainform.PMTCat.Enabled = True
Mainform.assigningcheque.Enabled = True

'under Transaction button
Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(1).Enabled = True
Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(2).Enabled = True
Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(3).Enabled = True
Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(4).Enabled = True
Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(5).Enabled = True
Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(6).Enabled = True
Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(7).Enabled = True
Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(8).Enabled = True
Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(9).Enabled = True
Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(10).Enabled = True

 'under Post button
Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(1).Enabled = True
Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(2).Enabled = True
Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(3).Enabled = True
Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(4).Enabled = True
Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(5).Enabled = True
Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(6).Enabled = True
Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(7).Enabled = True
Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(8).Enabled = True
Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(9).Enabled = True
Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(10).Enabled = True
Me.Timer2.Interval = 500


End Sub

Private Sub lvListView_DblClick()
If WhatJOurNCOde = "SAL" Then
 SalesJournal.Show 1
End If
End Sub

Private Sub lvListView_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call xPrint_Click
End If
End Sub

Private Sub lvListView_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
'Me.xPost.Enabled = False
'Me.xPrint.Enabled = False
'Me.xlist.Enabled = False
If Me.lvListView.ListItems.Count <> 0 Then
    If Trim(Me.lvListView.SelectedItem.SubItems(4)) = "Unpost" Then
      If Trim(UserRole) = "Accounts1" Or Trim(UserRole) = "Accounts2" Or Trim(UserRole) = "Admin" Or Trim(UserRole) = "Sales" Then
         'Me.xPost.Enabled = True
         If WhatJOurNCOde = "SAL" And Trim(UserRole) = "Sales" Or Trim(UserRole) = "Accounts1" Then
            Me.xPost.Enabled = True
            Me.xCancelDayJournal.Enabled = True
          ElseIf WhatJOurNCOde <> "SAL" And Trim(UserRole) <> "Sales" Then
           Me.xPost.Enabled = True
           Me.xCancelDayJournal.Enabled = True
          ElseIf Trim(UserRole) = "Admin" Then
           Me.xPost.Enabled = True
           Me.xCancelDayJournal.Enabled = True
          Else
          Me.xPost.Enabled = False
          Me.xCancelDayJournal.Enabled = False
         End If
        'If Trim(UserRole) = "Accounts1" Or Trim(UserRole) = "Accounts2" Or Trim(UserRole) = "Admin" Or Trim(UserRole) = "Sales" Then
         'Me.xPost.Enabled = True
        'End If
        'Me.xList.Enabled = True
        'Me.xPost.Enabled = True
       Else
        'Me.xList.Enabled = False
        Me.xPost.Enabled = False
        Me.xCancelDayJournal.Enabled = False
      End If
     Else
      'Me.xList.Enabled = False
      Me.xPost.Enabled = False
      Me.xCancelDayJournal.Enabled = False
    End If
  
    If Left(WhatJOurNCOde, 3) = "SAL" Then
            Me.xPrint.caption = "Print"
       Else
       Me.xPrint.caption = "Print Preview"
    End If
 If Button = 2 Then
    PopupMenu Me.Scmenu
 End If
 Else

End If
End Sub

Private Sub MAturingChkPmtRCT_Click()
dataanu.comuncollectedcheck_Grouping Format(Date, "mm/dd/yyyy")
reuncollectedchecks.Show 1
End Sub

Private Sub NewAcct_Click()
Unload NewAccts
FormNo = 0
xCountry = ""
xCountryNAMe = ""
xBranchName = ""
xClientCode = ""
TopLevelCode = ""
TopLevelName = ""
Level1Code = ""
Level2Code = ""
level3Code = ""
level4Code = ""
level5Code = ""
Level1Name = ""
Level2Name = ""
level3Name = ""
level4Name = ""
level5Name = ""
cTotalItems = ""

CancelAll = False
NewAccts.Show 1
End Sub

Private Sub PO_Click()
ProcessPO.Show 1
End Sub

Private Sub PPO_Click()
ProcessPO.Show 1
End Sub

Private Sub Opentables_Click()
Dim rsJourCode As New ADODB.Recordset
Me.tvTreeView.Nodes.clear
rsJourCode.Open "Select * from JOurnalCode order by Code", constring, adOpenKeyset, adLockPessimistic, adCmdText
Set nodex = Me.tvTreeView.Nodes.Add(, , "a", "Finance ÇáãÇáíÉ", 2)

nodex.Image = 4

Set nodex = Me.tvTreeView.Nodes.Add("a", tvwChild, "b", "Journals ÇáíæãíÇÊ ", 1)
Set nodex = Me.tvTreeView.Nodes.Add("a", tvwChild, "c", "Users ÇáãÓÊÎÏãíä ", 1)

Do Until rsJourCode.EOF = True
    Set nodex = Me.tvTreeView.Nodes.Add("b", tvwChild, , rsJourCode!Code & "-" & rsJourCode!JOurnalName, 1)
    rsJourCode.MoveNext
Loop
rsJourCode.close

Dim rsUsers As New ADODB.Recordset
rsUsers.Open "select * from Users order by Userid", constring, adOpenKeyset, adLockPessimistic, adCmdText
Do Until rsUsers.EOF = True
    Set nodex = Me.tvTreeView.Nodes.Add("c", tvwChild, , rsUsers!Userid & "[" & rsUsers!role & "]", IIf(rsUsers!logged = "Yes", 8, 7))
    rsUsers.MoveNext
Loop

End Sub

Private Sub paidVouc_Click()
DRPpaidvoucher.Show 1
End Sub

Private Sub Payee_Click()
frmpayee.Show 1
End Sub

Private Sub payinvoice_Click()
frmpaymentagaintsinvoice.Show 1
End Sub

Private Sub PAYJ_Click()
DataReport4.Show 1
End Sub

Private Sub paymentcheck_Click()
frmmaturingpaymentcheck.Show
End Sub

Private Sub paymentcheckcollecion_Click()
Command3_Click

End Sub

Private Sub PAYRCTColl_Click()

End Sub

Private Sub Payreq_Click()
PaymentRequestLast.Show 1
End Sub

Private Sub PAyset_Click()
DataReport2.Show 1
End Sub

Private Sub PMTCat_Click()
frmPayCat.Show 1
End Sub

Private Sub ProfitLost_Click()
frmprofitandlost.Show 1
End Sub

Private Sub RegisternewAsset_Click()
NewASset.Show
End Sub

Private Sub RepGen_Click()
FrmReportGenarator.Show 1
End Sub

Private Sub sbStatusBar_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
'Unload ScreenSaver
End Sub

Private Sub SOa_Click()
SOAPrn.Show 1
End Sub

Private Sub supplier_Click()
newSupplier.Text1.Text = "1"
newSupplier.Show 1
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "      "
        Case Is = "COA"
           AccTreeView.Show 1
    End Select
 
End Sub

Private Sub mnuHelpAbout_Click()
MsgBox "Coded by Ahmed-Al-Agroudy, Egypt"
End Sub

Private Sub mnuHelpSearchForHelpOn_Click()
    Dim nRet As Integer


    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(APp.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, APp.HelpFile, 261, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub

Private Sub mnuHelpContents_Click()
    Dim nRet As Integer


    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(APp.HelpFile) = 0 Then
        'MsgBox "Unable to display Help Contents. There is no Help associated with this system.", vbInformation, Me.Caption
        xmsg = MsgBox("We have no help file to display but I show you the following Function Keys to use to alternate the using of mouse." & vbCrLf & _
                                        "CTRL+O= Opentables" & vbCrLf & _
                                        "CTRL+F= Find Item in listview screen" & vbCrLf & _
                                        "CTRL+P= Process New Purchase Orders" & vbCrLf & _
                                        "CTRL+I= Process  New Invoice" & vbCrLf & _
                                        "ESC = Redisplay Previous List" & vbCrLf & _
                                        "DEL =  Delete an item" & vbCrLf & _
                                        "F2 =   Save Entries when make a payment" & vbCrLf & _
                                        "F3 =   Print Entries when make a payment", vbInformation + vbOKOnly, "Help")

                                        
                                        
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, APp.HelpFile, 3, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub


Private Sub mnuToolsOptions_Click()
    'ToDo: Add 'mnuToolsOptions_Click' code.
    MsgBox "Add 'mnuToolsOptions_Click' code."
End Sub

Private Sub mnuFileClose_Click()
    'unload the form
    Unload Me

End Sub

Private Sub mnuFileProperties_Click()
    'ToDo: Add 'mnuFileProperties_Click' code.
    MsgBox "Add 'mnuFileProperties_Click' code."
End Sub

Private Sub mnuFileRename_Click()
    'ToDo: Add 'mnuFileRename_Click' code.
    MsgBox "Add 'mnuFileRename_Click' code."
End Sub

Private Sub mnuFileDelete_Click()
    'ToDo: Add 'mnuFileDelete_Click' code.
    MsgBox "Add 'mnuFileDelete_Click' code."
End Sub

Private Sub mnuFileNew_Click()
    'ToDo: Add 'mnuFileNew_Click' code.
    MsgBox "Add 'mnuFileNew_Click' code."
End Sub

Private Sub mnuFileSendTo_Click()
    'ToDo: Add 'mnuFileSendTo_Click' code.
    MsgBox "Add 'mnuFileSendTo_Click' code."
End Sub

Private Sub mnuFileFind_Click()
    'ToDo: Add 'mnuFileFind_Click' code.
    MsgBox "Add 'mnuFileFind_Click' code."
End Sub

Private Sub mnuFileOpen_Click()
    Dim sFile As String


    With dlgCommonDialog
        .DialogTitle = "Open"
        .CancelError = False
        'ToDo: set the flags and attributes of the common dialog control
        .Filter = "All Files (*.*)|*.*"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
    'ToDo: add code to process the opened file

End Sub

Private Sub tbToolBar_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Select Case ButtonMenu.Key
    Case "GJProcess"
        GenJournalEntry.Show 1
    Case "AssetProcess"
        AssetSetup.Show 1
End Select
End Sub

Private Sub tbToolBar_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
'Unload ScreenSaver

End Sub

Private Sub Stmtofacc_Click()
RepStatementOfAcc.Show
End Sub

Private Sub taxdetails_Click()
frmtax.Show
End Sub

Private Sub Temrs_Click()
frmterms.Show
End Sub



Private Sub Timer1_Timer()
On Error Resume Next
ScreenSaver.Text1.Text = (ScreenSaver.Text1.Text) + 1
If Val(ScreenSaver.Text1) > 50 Then
    ScreenSaver.Show
 End If
End Sub

Private Sub Timer2_Timer()
'Dim rsdate As New ADODB.Recordset
'Dim serverdate As Date
'
'rsdate.Open "SELECT GETDATE() AS ServerDate ", constring, adOpenKeyset, adLockPessimistic, adCmdText
'serverdate = Format(rsdate!serverdate, "mm/dd/yyyy")
'If Date <> serverdate Then
'    mess = MsgBox("Your Machine Date is not valid!, You are not able to take up any transactions. Date Today is: " & serverdate & " áíÓ ÇáÇáÉ ÊÚãá Ýí ÇáÊÇÑíÎ  ÛíÑ ÞÇÏÑ Úáí Úãá Çí ÊÍæíá ", vbExclamation + vbOKOnly, "Invalid Date ÊÇÑíÎ ÎØÇð")
'    Exit Sub
'End If

End Sub

Private Sub Timer3_Timer()
On Error GoTo nel
Dim JustTEstConnection As New ADODB.Connection
JustTEstConnection.Open "dsN=fINANCe;UID=Sa;PWD=;"
JustTEstConnection.close
Dim RstBC As New ADODB.Recordset
RstBC.Open "Setup", constring, adOpenKeyset, adLockPessimistic, adCmdTable
xBranchCode = RstBC!BranchCode
xBranchName = RstBC!BranchNameENg
Me.caption = "Finance System-" & xBranchName
xCountry = RstBC!countrycode
xCountryNAMe = RstBC!CountryName
RstBC.close

nel:
n = Err.Number
c = Err.Description
Debug.Print c
If Trim(c) = "[Microsoft][ODBC Driver Manager] Data source name not found and no default driver specified" Then
   msg = MsgBox("DSN not found. Check it at ODBC", vbExclamation + vbOKOnly, "Connection Error")
   Unload Me
 ElseIf Trim(c) = "[Microsoft][ODBC Driver Manager] Cannot generate SSPI context" Then
   msg = MsgBox("Please Logoff or restart the machine. Cannot generate SSPI context", vbExclamation + vbOKOnly, "Connection Error")
   Unload Me
End If

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case Is = "login"
        LogUser.caption = "Log-in User ÇÏÎá ÇÓã ÇáãÓÊÎÏã"
        LogUser.Show 1
        If LogSucess = True Then
            EnablesAllObject
            Call Opentables_Click
        End If

    Case Is = "Sql"
        FrmSQL.Show
    Case Is = "FindAccts"
        FindAcctNAme = False
        FindAcctNames.Show
    Case Is = "Charts"
             
        Me.sbStatusBar.Panels(1).Text = "Status : Now Loading Chart of Accounts..."
        AccTreeView.Show
        Me.sbStatusBar.Panels(1).Text = "Status : Ready"
End Select
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)

End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Dim rsInvj As New ADODB.Recordset
Dim MItem As ListItem
Select Case ButtonMenu.Key
    Case Is = "xGJ"
            GenJournalEntry.Show
    Case Is = "xAJ"
            AssetSetup.Show
    Case Is = "xSJ"
            DownLoadinvoices.Show 1
    Case Is = "xIJ"
            InvJournalEntry.Show
     Case Is = "BT"
            BankTransaction1.Show
    Case Is = "xCAshRec"
            frmrecieptvou.Show
    Case Is = "xcashpay"
            frmpaymentvou.Show
    Case Is = "xPayables"
            FrmPayableSetup.Show
    Case Is = "xPettyCash"
            frmPettyCash.Show
    'for post button
    Case Is = "xGenJourPost"
            mess = MsgBox("Do you want to continue?", vbQuestion + vbOKCancel + vbDefaultButton2, "Please confirm")
            If mess = vbOK Then
                PostingJournal.Text1.Text = "GEN"
                rsInvj.Open "SELECT  TransDate, COUNT(TransDate) AS TotalTRan, SUM(DebitAmount) AS DrAmt, SUM(CreditAmount) AS CrAmt" _
                & " From GEnJournalTrans where remarks is null GROUP BY TransDate ORDER BY TransDate", constring, adOpenKeyset, adLockPessimistic, adCmdText
            Call OPenDayTran(rsInvj)
            End If
    
    Case Is = "xAssetJourPost"
            mess = MsgBox("Do you want to continue?", vbQuestion + vbOKCancel + vbDefaultButton2, "Please confirm")
            If mess = vbOK Then
                 PostingJournal.Text1.Text = "AST"
                rsInvj.Open "SELECT  TransDate, COUNT(TransDate) AS TotalTRan, SUM(DebitAmount) AS DrAmt, SUM(CreditAmount) AS CrAmt" _
                & " From AssetJournal where Remarks is null GROUP BY TransDate ORDER BY TransDate", constring, adOpenKeyset, adLockPessimistic, adCmdText
                Call OPenDayTran(rsInvj)
            End If
            
    Case Is = "xInventoryJOurPost"
            mess = MsgBox("Do you want to continue?", vbQuestion + vbOKCancel + vbDefaultButton2, "Please confirm")
            If mess = vbOK Then
                 PostingJournal.Text1.Text = "IVY"
                rsInvj.Open "SELECT  TransDate, COUNT(TransDate) AS TotalTRan, SUM(DebitAmount) AS DrAmt, SUM(CreditAmount) AS CrAmt" _
                & " From InventoryJournal where Remarks is null GROUP BY TransDate ORDER BY TransDate", constring, adOpenKeyset, adLockPessimistic, adCmdText
                Call OPenDayTran(rsInvj)
            End If
    Case Is = "xSalesJOurPost"
            mess = MsgBox("Do you want to continue?", vbQuestion + vbOKCancel + vbDefaultButton2, "Please confirm")
            If mess = vbOK Then
                 PostingJournal.Text1.Text = "IVY"
                rsInvj.Open "SELECT  TransDate, COUNT(TransDate) AS TotalTRan, SUM(DebitAmount) AS DrAmt, SUM(CreditAmount) AS CrAmt" _
                & " From InventoryJournal where Remarks is null GROUP BY TransDate ORDER BY TransDate", constring, adOpenKeyset, adLockPessimistic, adCmdText
                Call OPenDayTran(rsInvj)
            End If
            
            
    Case Is = "xCashRecJOurpost"
            mess = MsgBox("Do you want to continue?", vbQuestion + vbOKCancel + vbDefaultButton2, "Please confirm")
            If mess = vbOK Then
                 PostingJournal.Text1.Text = "CSR"
                 rsInvj.Open "SELECT  TransDate, COUNT(TransDate) AS TotalTRan, SUM(DebitAmount) AS DrAmt, SUM(CreditAmount) AS CrAmt" _
                 & " From CAshJournal where Remarks is null and trantype = 'R' GROUP BY TransDate ORDER BY TransDate", constring, adOpenKeyset, adLockPessimistic, adCmdText
                 Call OPenDayTran(rsInvj)
            End If
    Case Is = "CashPayJourPost"
            mess = MsgBox("Do you want to continue?", vbQuestion + vbOKCancel + vbDefaultButton2, "Please confirm")
            If mess = vbOK Then
                PostingJournal.Text1.Text = "CSP"
                rsInvj.Open "SELECT  TransDate, COUNT(TransDate) AS TotalTRan, SUM(DebitAmount) AS DrAmt, SUM(CreditAmount) AS CrAmt" _
                & " From CAshJournal where Remarks is null and trantype = 'P' GROUP BY TransDate ORDER BY TransDate", constring, adOpenKeyset, adLockPessimistic, adCmdText
                Call OPenDayTran(rsInvj)
            End If
    Case Is = "xCreditNote"
            mess = MsgBox("Do you want to continue?", vbQuestion + vbOKCancel + vbDefaultButton2, "Please confirm")
            If mess = vbOK Then
                 PostingJournal.Text1.Text = "SRL"
                 rsInvj.Open "SELECT  TransDate, COUNT(TransDate) AS TotalTRan, SUM(DebitAmount) AS DrAmt, SUM(CreditAmount) AS CrAmt" _
                 & " From salesreturnjournal where Remarks is null GROUP BY TransDate ORDER BY TransDate", constring, adOpenKeyset, adLockPessimistic, adCmdText
                 'MsgBox rsInvj.RecordCount
                 Call OPenDayTran(rsInvj)
            End If

End Select
End Sub
Sub OPenDayTran(rsInvj As Recordset)
    If rsInvj.EOF = True Then
      mess = MsgBox("No Transactions to post", vbExclamation + vbOKOnly, "Message")
      Exit Sub
    End If
    
    Do Until rsInvj.EOF = True
                    Set MItem = PostingJournal.ListView1.ListItems.Add(, , rsInvj!TRansDate)
                    MItem.SubItems(1) = rsInvj!TotalTRan
                    MItem.SubItems(2) = FormatNumber(rsInvj!Dramt, 2, vbTrue, vbTrue, vbTrue)
                    MItem.SubItems(3) = FormatNumber(rsInvj!CrAmt, 2, vbTrue, vbTrue, vbTrue)
                    MItem.SubItems(4) = "Waiting"
                    rsInvj.MoveNext
    Loop
    PostingJournal.Show 1
End Sub
Private Sub trialbalance_Click()
On Error Resume Next
frmtrialbalance.Show 1
On Error GoTo 0
End Sub

Private Sub tvTreeView_Collapse(ByVal Node As MSComctlLib.Node)
cindex = Node.Index
If cindex = 1 Then
 Me.tvTreeView.Nodes.Item(cindex).Image = 4
 Else
  Me.tvTreeView.Nodes.Item(cindex).Image = 1
End If
End Sub

Private Sub tvTreeView_Expand(ByVal Node As MSComctlLib.Node)
cindex = Node.Index
If cindex = 1 Then
 Me.tvTreeView.Nodes.Item(cindex).Image = 5
 Else
  Me.tvTreeView.Nodes.Item(cindex).Image = 2
End If
End Sub

Private Sub tvTreeView_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
'If Me.tvTreeView.Nodes.Count <> 0 Then
' If Button = 2 Then
'    PopupMenu Scmenu.ModifyTV
' End If
'End If
End Sub

Private Sub tvTreeView_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
'Unload ScreenSaver
End Sub

Private Sub tvTreeView_NodeClick(ByVal Node As MSComctlLib.Node)

i = 0
For i = 1 To Me.tvTreeView.Nodes.Count
    If Mid(Me.tvTreeView.Nodes.Item(i).Text, 4, 1) = "-" Then
      Me.tvTreeView.Nodes.Item(i).Image = 1
    End If
Next
If Mid(Node.Text, 4, 1) = "-" Then
  Node.Image = 2
End If
Dim rstJournals As New ADODB.Recordset
Dim rstJournals2 As New ADODB.Recordset
Dim rstJournals3 As New ADODB.Recordset

If WhatJOurNCOde = Left(Node.Text, 3) Then
    Exit Sub
End If
Me.lvListView.ListItems.clear
On Error Resume Next
rstJournals.close
rstJournals2.close
On Error GoTo 0
    
WhatJOurNCOde = Left(Node.Text, 3)

If Left(Node.Text, 3) = "AST" Then
    'query unposted transactions
    rstJournals.Open "SELECT  TransDate, COUNT(TransDate) AS TotalTRan, SUM(debitamount) AS DrAmt, SUM(CreditAmount) AS CrAmt  " _
             & " From AssetJournal where remarks is null GROUP BY TransDate ORDER BY TransDate", constring, adOpenKeyset, adLockPessimistic, adCmdText
    'query also posted transactions
    rstJournals2.Open "SELECT  TransDate, COUNT(TransDate) AS TotalTRan, SUM(DebitAmount) AS DrAmt, SUM(CreditAmount) AS CrAmt  " _
             & " From AssetJournal where remarks is not null GROUP BY TransDate ORDER BY TransDate", constring, adOpenKeyset, adLockPessimistic, adCmdText
    'query for posted transactions by JN
'    rstJournals3.Open "SELECT COUNT(DISTINCT SerialNo) AS cTOtal from assetJournal WHERE (Remarks IS NULL) where remarks is not null GROUP BY TransDate ORDER BY TransDate", constring, adOpenKeyset, adLockPessimistic, adCmdText
             
             
ElseIf Left(Node.Text, 3) = "SAL" Then
    'query unposted transactions
    rstJournals.Open "SELECT  TransDate, COUNT(TransDate) AS TotalTRan, SUM(tradercvble+tradeDiscAmt) AS DrAmt, SUM(GrossSales+Vat+transpoCharge+surtaxamt) AS CrAmt  " _
             & " From SalesJournal where remarks is null GROUP BY TransDate ORDER BY TransDate", constring, adOpenKeyset, adLockPessimistic, adCmdText
    'query also posted transactions
    rstJournals2.Open "SELECT  TransDate, COUNT(TransDate) AS TotalTRan, SUM(tradercvble+tradeDiscAmt) AS DrAmt, SUM(GrossSales+Vat+transpoCharge+surtaxamt) AS CrAmt  " _
             & " From SalesJournal where remarks is not null GROUP BY TransDate ORDER BY TransDate", constring, adOpenKeyset, adLockPessimistic, adCmdText
             
ElseIf Left(Node.Text, 3) = "IVY" Then
    'query unposted transactions
    rstJournals.Open "SELECT  TransDate, COUNT(TransDate) AS TotalTRan, SUM(DebitAmount) AS DrAmt, SUM(CreditAmount) AS CrAmt  " _
             & " From InventoryJournal where remarks is null GROUP BY TransDate ORDER BY TransDate", constring, adOpenKeyset, adLockPessimistic, adCmdText
    'query also posted transactions
    rstJournals2.Open "SELECT  TransDate, COUNT(TransDate) AS TotalTRan, SUM(DebitAmount) AS DrAmt, SUM(CreditAmount) AS CrAmt  " _
             & " From InventoryJournal where remarks is not null GROUP BY TransDate ORDER BY TransDate", constring, adOpenKeyset, adLockPessimistic, adCmdText
             
ElseIf Left(Node.Text, 3) = "GEN" Then
    'query unposted transactions
    rstJournals.Open "SELECT  TransDate, COUNT(TransDate) AS TotalTRan, SUM(DebitAmount) AS DrAmt, SUM(CreditAmount) AS CrAmt  " _
             & " From GenJournalTrans where remarks is null GROUP BY TransDate ORDER BY TransDate", constring, adOpenKeyset, adLockPessimistic, adCmdText
    'query also posted transactions
    rstJournals2.Open "SELECT  TransDate, COUNT(TransDate) AS TotalTRan, SUM(DebitAmount) AS DrAmt, SUM(CreditAmount) AS CrAmt  " _
             & " From GenJournalTrans where remarks is not null GROUP BY TransDate ORDER BY TransDate", constring, adOpenKeyset, adLockPessimistic, adCmdText
             
             
ElseIf Left(Node.Text, 3) = "PYB" Then
    'query unposted transactions
    rstJournals.Open "SELECT  confirmedDate, COUNT(confirmedDate) AS TotalTRan, SUM(DbAmount) AS DrAmt, SUM(CrAmount) AS CrAmt  " _
             & " From PayJournal where Status = 'Unposted'  GROUP BY confirmedDate ORDER BY confirmedDate", constring, adOpenKeyset, adLockPessimistic, adCmdText
    'query also posted transactions
    rstJournals2.Open "SELECT  confirmedDate, COUNT(confirmedDate) AS TotalTRan, SUM(DbAmount) AS DrAmt, SUM(CrAmount) AS CrAmt  " _
             & " From PayJournal where status ='Posted' GROUP BY confirmedDate ORDER BY confirmedDate", constring, adOpenKeyset, adLockPessimistic, adCmdText
                          
ElseIf Left(Node.Text, 3) = "CSP" Then
    'query unposted transactions
    rstJournals.Open "SELECT  TransDate, COUNT(TransDate) AS TotalTRan, SUM(DebitAmount) AS DrAmt, SUM(CreditAmount) AS CrAmt  " _
             & " From cashjournal where remarks is null and trantype = 'P' GROUP BY TransDate ORDER BY TransDate", constring, adOpenKeyset, adLockPessimistic, adCmdText
    'query also posted transactions
    rstJournals2.Open "SELECT  TransDate, COUNT(TransDate) AS TotalTRan, SUM(DebitAmount) AS DrAmt, SUM(CreditAmount) AS CrAmt  " _
             & " From cashjournal where remarks is not null and trantype = 'P' GROUP BY TransDate ORDER BY TransDate", constring, adOpenKeyset, adLockPessimistic, adCmdText


ElseIf Left(Node.Text, 3) = "CSR" Then
    'query unposted transactions
    rstJournals.Open "SELECT  TransDate, COUNT(TransDate) AS TotalTRan, SUM(DebitAmount) AS DrAmt, SUM(CreditAmount) AS CrAmt  " _
             & " From cashjournal where remarks is null and trantype = 'R' GROUP BY TransDate ORDER BY TransDate", constring, adOpenKeyset, adLockPessimistic, adCmdText
    'query also posted transactions
    rstJournals2.Open "SELECT  TransDate, COUNT(TransDate) AS TotalTRan, SUM(DebitAmount) AS DrAmt, SUM(CreditAmount) AS CrAmt  " _
             & " From cashjournal where remarks is not null and trantype = 'R' GROUP BY TransDate ORDER BY TransDate", constring, adOpenKeyset, adLockPessimistic, adCmdText
             
ElseIf Left(Node.Text, 3) = "SRL" Then
    'query unposted transactions
    rstJournals.Open "SELECT  TransDate, COUNT(TransDate) AS TotalTRan, SUM(DebitAmount) AS DrAmt, SUM(CreditAmount) AS CrAmt  " _
             & " From creditnote where remarks is null and trantype = 'SRL' GROUP BY TransDate ORDER BY TransDate", constring, adOpenKeyset, adLockPessimistic, adCmdText
    'query also posted transactions
    rstJournals2.Open "SELECT  TransDate, COUNT(TransDate) AS TotalTRan, SUM(DebitAmount) AS DrAmt, SUM(CreditAmount) AS CrAmt  " _
             & " From creditnote where remarks is not null and trantype = 'SRL' GROUP BY TransDate ORDER BY TransDate", constring, adOpenKeyset, adLockPessimistic, adCmdText

ElseIf Left(Node.Text, 3) = "SPA" Then
    'query unposted transactions
    rstJournals.Open "SELECT  TransDate, COUNT(TransDate) AS TotalTRan, SUM(DebitAmount) AS DrAmt, SUM(CreditAmount) AS CrAmt  " _
             & " From debitnote where remarks is null and trantype = 'SPA' GROUP BY TransDate ORDER BY TransDate", constring, adOpenKeyset, adLockPessimistic, adCmdText
    'query also posted transactions
    rstJournals2.Open "SELECT  TransDate, COUNT(TransDate) AS TotalTRan, SUM(DebitAmount) AS DrAmt, SUM(CreditAmount) AS CrAmt  " _
             & " From debitnote  where remarks is not null and trantype = 'SPA' GROUP BY TransDate ORDER BY TransDate", constring, adOpenKeyset, adLockPessimistic, adCmdText

ElseIf Left(Node.Text, 3) = "PTC" Then
    'query unposted transactions
    rstJournals.Open "SELECT  ConfirmedDate, COUNT(ConfirmedDate) AS TotalTRan, SUM(DbAmount) AS DrAmt, SUM(CrAmount) AS CrAmt  " _
             & " From pettyjournal where postMark is null  GROUP BY ConfirmedDate ORDER BY ConfirmedDate", constring, adOpenKeyset, adLockPessimistic, adCmdText
    'query also posted transactions
     rstJournals2.Open "SELECT  ConfirmedDate, COUNT(ConfirmedDate) AS TotalTRan, SUM(DbAmount) AS DrAmt, SUM(CrAmount) AS CrAmt  " _
             & " From pettyjournal where postMark is not null  GROUP BY ConfirmedDate ORDER BY ConfirmedDate", constring, adOpenKeyset, adLockPessimistic, adCmdText
             
ElseIf Left(Node.Text, 3) = "BNK" Then
    'query unposted transactions
    rstJournals.Open "SELECT  TransDate, COUNT(TransDate) AS TotalTRan, SUM(DebitAmount) AS DrAmt, SUM(CreditAmount) AS CrAmt  " _
             & " From BankJOurnal where remarks is null GROUP BY TransDate ORDER BY TransDate", constring, adOpenKeyset, adLockPessimistic, adCmdText
    'query also posted transactions
    rstJournals2.Open "SELECT  TransDate, COUNT(TransDate) AS TotalTRan, SUM(DebitAmount) AS DrAmt, SUM(CreditAmount) AS CrAmt  " _
             & " From BankJOurnal where remarks is not null GROUP BY TransDate ORDER BY TransDate", constring, adOpenKeyset, adLockPessimistic, adCmdText
             
Else
  Exit Sub
End If



'display also posted days

Do Until rstJournals2.EOF = True
    If Left(Node.Text, 3) = "PYB" Then
        trDate = Format(rstJournals2!confirmeddate, "dd/mm/yyyy")
     ElseIf Left(Node.Text, 3) = "PTC" Then
        trDate = Format(rstJournals2!confirmeddate, "dd/mm/yyyy")
     Else
        trDate = Format(rstJournals2!TRansDate, "dd/mm/yyyy")
    End If
    Set MItem = Me.lvListView.ListItems.Add(, , trDate, 3, 3)
    MItem.SubItems(1) = rstJournals2!TotalTRan
    MItem.SubItems(2) = FormatNumber(rstJournals2!Dramt, 2, vbTrue, vbTrue, vbTrue)
    MItem.SubItems(3) = FormatNumber(rstJournals2!CrAmt, 2, vbTrue, vbTrue, vbTrue)
    MItem.SubItems(4) = "Posted"
    rstJournals2.MoveNext
    DoEvents
Loop

'display unposted days
Do Until rstJournals.EOF = True
    If Left(Node.Text, 3) = "PYB" Then
        trDate = Format(rstJournals!confirmeddate, "dd/mm/yyyy")
     ElseIf Left(Node.Text, 3) = "PTC" Then
        trDate = Format(rstJournals!confirmeddate, "dd/mm/yyyy")
     Else
        trDate = Format(rstJournals!TRansDate, "dd/mm/YYYy")
    End If
    Set MItem = Me.lvListView.ListItems.Add(, , trDate, 4, 4)
    MItem.SubItems(1) = rstJournals!TotalTRan
    MItem.SubItems(2) = FormatNumber(rstJournals!Dramt, 2, vbTrue, vbTrue, vbTrue)
    MItem.SubItems(3) = FormatNumber(rstJournals!CrAmt, 2, vbTrue, vbTrue, vbTrue)
    MItem.SubItems(4) = "Unpost"
    rstJournals.MoveNext
    DoEvents
Loop


Nelson:
Exit Sub
End Sub

Private Sub uncollectedpaycheck_Click()
End Sub

Private Sub uncollectedrecheck_Click()
End Sub


'***************************  anushath



Private Sub cmdcheckcollections_Click()

End Sub


'Private Sub Command2_Click()
'secondvariable = DateAdd("d", 7, Date)
''MsgBox Format(secondvariable, "dd/mm/yyyy")
''MsgBox Format(Date, "dd/mm/yyyy")
'dataanu.compaymentcheck_Grouping Format(secondvariable, "mm/dd/yyyy"), Format(Date, "mm/dd/yyyy")
'FormatLabelcheckdate repaymentcheck.Sections(2).Controls("label20"), _
'        "Company Report "
'repaymentcheck.Show 1
'
'End Sub

Private Sub Command3_Click()
dataanu.comuncollectedpaymentscheck_Grouping Format(Date, "mm/dd/yyyy")
reuncollectedpaymentchecks.Show 1

End Sub

Private Sub Form_Load()
'con.Open "dsN=fINANCE;UID=Sa;PWD=;"
constring = "dsN=fINANCE;UID=SA;PWD="

End Sub

'this is for beginning
Private Sub prcbeginingcash(lblX As RptLabel)
      lblX.caption = Format(openingcash, "###,###,###,##0.#0")
End Sub
Private Sub prcbeginingcheck(lblX As RptLabel)
      lblX.caption = Format(openingcheck, "###,###,###,##0.#0")
End Sub
Private Sub prcbeginingcreditcard(lblX As RptLabel)
      lblX.caption = Format(openingcreditcard, "###,###,###,##0.#0")
End Sub
'end beginning

'this is for ending
Private Sub prcendingcash(lblX As RptLabel)
      lblX.caption = Format(finalcashtotal, "###,###,###,##0.#0")
End Sub
Private Sub prcendingcheck(lblX As RptLabel)
      lblX.caption = Format(finalchecktotal, "###,###,###,##0.#0")
End Sub
Private Sub prcendingcreditcard(lblX As RptLabel)
      lblX.caption = Format(finalcreditcardtotal, "###,###,###,##0.#0")
End Sub
'end ending

Private Sub FormatLabelAC1(lblX As RptLabel)
      lblX.caption = Format(openingcash, "###,###,###,##0.#0")
End Sub
Private Sub FormatLabelAC2(lblX As RptLabel)
      lblX.caption = Format(openingcheck, "###,###,###,##0.#0")
End Sub
Private Sub FormatLabelAC3(lblX As RptLabel)
      lblX.caption = Format(openingcreditcard, "###,###,###,##0.#0")
End Sub


Private Sub FormatLabelAC4(lblX As RptLabel)
      lblX.caption = Format(collectioncash, "###,###,###,##0.#0")
End Sub
Private Sub FormatLabelAC5(lblX As RptLabel)
      lblX.caption = Format(collectioncheck, "###,###,###,##0.#0")
End Sub
Private Sub FormatLabelAC6(lblX As RptLabel)
      lblX.caption = Format(collectioncreditcard, "###,###,###,##0.#0")
End Sub


Private Sub FormatLabelAC7(lblX As RptLabel)
      lblX.caption = Format(paymentcash, "###,###,###,##0.#0")
End Sub
Private Sub FormatLabelAC8(lblX As RptLabel)
      lblX.caption = Format(paymentcheck11, "###,###,###,##0.#0")
End Sub
Private Sub FormatLabelAC9(lblX As RptLabel)
      lblX.caption = Format(paymentcreditcard, "###,###,###,##0.#0")
End Sub


Private Sub FormatLabelAC10(lblX As RptLabel)
      lblX.caption = Format(finalcashtotal, "###,###,###,##0.#0")
End Sub
Private Sub FormatLabelAC11(lblX As RptLabel)
      lblX.caption = Format(finalchecktotal, "###,###,###,##0.#0")
End Sub
Private Sub FormatLabelAC12(lblX As RptLabel)
      lblX.caption = Format(finalcreditcardtotal, "###,###,###,##0.#0")
End Sub

Private Sub FormatLabelAC13(lblX As RptLabel)
      lblX.caption = Format(finalcashtotal, "###,###,###,##0.#0")
End Sub
Private Sub FormatLabelAC14(lblX As RptLabel)
      lblX.caption = Format(finalchecktotal, "###,###,###,##0.#0")
End Sub
Private Sub FormatLabelAC15(lblX As RptLabel)
      lblX.caption = Format(finaltotal, "###,###,###,##0.#0")
End Sub
'this is for doller
Private Sub FormatLabelAC16(lblX As RptLabel)
      lblX.caption = Format(opendoller, "###,###,###,##0.#0")
End Sub
Private Sub FormatLabelAC17(lblX As RptLabel)
      lblX.caption = Format(receiptdoller, "###,###,###,##0.#0")
End Sub

Private Sub FormatLabelAC18(lblX As RptLabel)
      lblX.caption = Format(middoller, "###,###,###,##0.#0")
End Sub
Private Sub FormatLabelAC19(lblX As RptLabel)
      lblX.caption = Format(paydoller, "###,###,###,##0.#0")
End Sub
Private Sub FormatLabelAC20(lblX As RptLabel)
      lblX.caption = Format(closedoller, "###,###,###,##0.#0")
End Sub
' this is for receipt check collections

Private Sub xAcctGrouping_Click()
Grouping.Show
End Sub

Private Sub xBankAcctLedger_Click()
BankAccountLedger.Show

End Sub

Private Sub xBankBalances_Click()
Dim xClass As New HabitatClass
Dim rsBA As New ADODB.Recordset
Dim conBA As New ADODB.Connection
Dim rsBABal As New ADODB.Recordset
Dim rsBalance As New ADODB.Recordset
Dim xtable As String
Dim sqltable As Boolean
Dim TOtdR As Currency
Dim TOTcr As Currency
rsBABal.Open "Delete BankAccountBalances", constring, adOpenKeyset, adLockPessimistic, adCmdText
rsBABal.Open "BankAccountBalances", constring, adOpenKeyset, adLockPessimistic, adCmdTable
xtable = "SELECT * from Level6 WHERE LEFT(AccountCode, 7) = '1110201'" 'current account
sqltable = True
xClass.GetTables rsBA, conBA, xtable, constring, sqltable
While rsBA.EOF = False
    acctNo = Trim(rsBA!AccountCode)
    
    'chek the balance in GLMaster Table
    rsBalance.Open "SElect sum(Debitamount) as TotDr,sum(creditamount) as TotCr from GLMaster where accountcode=" & "'" & acctNo & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
    
    TOtdR = IIf(IsNull(rsBalance!TOtdR) = True, 0, rsBalance!TOtdR)
    TOTcr = IIf(IsNull(rsBalance!TOTcr) = True, 0, rsBalance!TOTcr)
    Balances = TOtdR - TOTcr
    rsBalance.close
    'chek the the Last Trans date
    rsBalance.Open "SElect Max(recordDate) as AsOf from GLMaster where accountcode=" & "'" & acctNo & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
    Asof = rsBalance!Asof
    rsBalance.close
    If rsBA!level5Code <> 5 Then
        With rsBABal
            .addnew
            !rateCode = rsBA!level5Code
            !BankAcctNameEng = rsBA!accountnameeng
            !BankAcctNamearab = rsBA!accountnamearab
            !GlAcctCode = rsBA!AccountCode
            !AccountType = rsBA!level4Name
            !currency = rsBA!level5Name
            If rsBA!level5Code = 1 Then
              !LE1 = Balances
             ElseIf rsBA!level5Code = 2 Then
              !Le2 = Balances
             ElseIf rsBA!level5Code = 3 Then
              !Le3 = Balances
             ElseIf rsBA!level5Code = 4 Then
              !LE4 = Balances
            End If
            !Balances = Balances
            !Asof = Asof
            .Update
            Balances = 0
        End With
     End If
     rsBA.MoveNext
Wend
rsBABal.close
'rate conversion
Dim rsRate As New ADODB.Recordset
rsBABal.Open "sELect * from BankAccountBalances order by ratecode", constring, adOpenKeyset, adLockPessimistic, adCmdText
With rsBABal
    '.MoveFirst
   Do Until .EOF = True
      xrateCode = Trim(rsBABal!rateCode)
      rsRate.Open "Select * from CurrencyRate where Code=" & "'" & xrateCode & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
      If xrateCode = 2 Then
         convertedamt = !Le2 / rsRate!EgyptpoundRate
         !UsDollar2 = convertedamt
         .Update
       ElseIf xrateCode = 3 Then
         convertedamt = !Le3 / rsRate!EgyptpoundRate
         !Euro = convertedamt
         .Update
        ElseIf xrateCode = 4 Then
         convertedamt = !Le3 / rsRate!EgyptpoundRate
         !UkPound = convertedamt
         .Update
       End If
       .MoveNext
       rsRate.close
    Loop
End With
On Error Resume Next
FinanceDE.rsBankBalances.close
FinanceDE.BankBalances
BankBalances.Show
End Sub

Private Sub xCancelDayJournal_Click()

TRansDate = Trim(Mainform.lvListView.SelectedItem.Text) ' , "mm/dd/yyyy")
msg = MsgBox("This will delete all " & TRansDate & " Journal" & vbCrLf _
 & "Do you want to delete ?", vbQuestion + vbOKCancel + vbDefaultButton2, "Please confirm")
If msg = vbOK Then
    Dim rsJn As New ADODB.Recordset
    cindex = Me.lvListView.SelectedItem.Index
    
    SelectedDate = Mid(TRansDate, 4, 2) & "/" & Left(TRansDate, 2) & "/" & Mid(TRansDate, 7, 4)
    If WhatJOurNCOde = "SAL" Then
        xtable = "SalesJournal"
     Else
        Exit Sub
    End If
    
   
    Dim rsInv As New ADODB.Recordset
    Dim rsCon As New ADODB.Connection
    Dim TAbleStr As String
    Dim sqltable As Boolean
    Dim xcls As New HabitatClass
    Dim xSelectedDate As Date
    sqltable = True
    xSelectedDate = SelectedDate
    TAbleStr = "Select * from invcmain where invc_Date=" & "ctod" & "('" & xSelectedDate & "')"
    'TAbleStr = "Select invc_Date, download from invcmain"
    'xcls.GetTables rsInv, rsCon, TAbleStr, "DSN=Invoices;Uid=sa;pwd=;", sqltable
    rsInv.Open "Select * from invcmain where invc_Date=" & "ctod" & "('" & xSelectedDate & "')", "DSN=Invoices;Uid=sa;pwd=;", adOpenKeyset, adLockPessimistic, adCmdText
    If rsInv.EOF = True Then
        msg = MsgBox("Can't reset the invoices")
        rsInv.close
        Exit Sub
     Else
      While rsInv.EOF = False
       'If rsInv!invc_date = xSelectedDate Then
        rsInv!Download = False
        rsInv.Update
       'End If
       rsInv.MoveNext
       
      Wend
      rsInv.close
     End If
     
    rsJn.Open "Delete " & xtable & " where transdate = " & "'" & SelectedDate & "'", constring, adOpenKeyset, adLockPessimistic, adCmdText
    Me.lvListView.ListItems.Remove cindex
    msg = MsgBox("Successfully deleted", vbInformation + vbOKOnly, "Message")
     
    
End If
End Sub

Private Sub xCashPmt_Click()
frmpaymentvou.Show 1
End Sub

Private Sub xcAShRct_Click()
frmrecieptvou.Show 1
End Sub

Private Sub xDetails_Click()
Me.lvListView.View = lvwReport
Me.xLI.Checked = False
Me.xSI.Checked = False
Me.xViewList.Checked = False
Me.xDetails.Checked = True

End Sub

Private Sub xfindaccounts_Click()
FindAcctNAme = False
FindAcctNames.Show
End Sub

Private Sub xFixedASset_Click()
AssetSetup.Show
End Sub

Private Sub xGEnJOurn_Click()
 GenJournalEntry.Show
End Sub

Private Sub xInventory_Click()
InvJournalEntry.Show
End Sub

Private Sub xLI_Click()
Me.lvListView.View = lvwIcon
Me.xLI.Checked = True
Me.xSI.Checked = False
Me.xViewList.Checked = False
Me.xDetails.Checked = False

End Sub

Private Sub xList_Click()


If WhatJOurNCOde = "SAL" Then
 SalesJournal.Show 1
End If
End Sub

Private Sub xPaySetup_Click()
FrmPayableSetup.Show
End Sub

Private Sub xPrint_Click()
'Dim transdate As Date
If Me.lvListView.ListItems.Count <> 0 Then
    Dim SelectedDate As Date
    TRansDate = Trim(Mainform.lvListView.SelectedItem.Text) ' , "mm/dd/yyyy")
    SelectedDate = Mid(TRansDate, 4, 2) & "/" & Left(TRansDate, 2) & "/" & Mid(TRansDate, 7, 4)
    remark = Trim(Me.lvListView.SelectedItem.SubItems(4))
    
    On Error Resume Next
    'for sales Journal
    If WhatJOurNCOde = "SAL" Then
      On Error GoTo CancelPrn
      Me.dlgCommonDialog.CancelError = True
      Me.dlgCommonDialog.ShowPrinter
    End If
    
        
    'for General Journal
    If WhatJOurNCOde = "GEN" And remark = "Posted" Then
       'PuCaption GenJOurnalPosted.Sections(1).Controls("Label7"), _
        "Company Report "
       On Error Resume Next
       FinanceDE.rsGenJournalPosted.close
       FinanceDE.GenJOurnalPosted SelectedDate, WhatJOurNCOde
       GenJOurnalPosted.Show
       
    
       
'     ElseIf WhatJOurNCOde = "GEN" And remark = "Unpost" Then
'        Dim rsTRans As New ADODB.Recordset
'        rsTRans.Open "select Count(*) as GTotal,sum(Debitamount) as DRTotal,Sum(CreditaMount) as CrTotal from genjournaltrans where transdate=" & "'" & SelectedDate & "'" & " and remarks is null", constring, adOpenKeyset, adLockPessimistic, adCmdText
'        If rsTRans.EOF = False Then
'        gTotal = rsTRans!gTotal
'        dRTotal = FormatNumber(rsTRans!dRTotal, 2, vbTrue, vbTrue, vbTrue)
'        cRTotal = FormatNumber(rsTRans!cRTotal, 2, vbTrue, vbTrue, vbTrue)
'        GenJournalTotalTrn = gTotal
'        End If
'        Dim xcaption As String
'        PuCaption byGenDescrition.Sections(2).Controls("Label13"), "Company Report "
'        xcaption = gTotal
'        PutGTotal byGenDescrition.Sections(7).Controls("Label16"), "Company Report", xcaption
'        xcaption = dRTotal
'        PutGTotal byGenDescrition.Sections(7).Controls("Label17"), "Company Report", xcaption
'        xcaption = cRTotal
'        PutGTotal byGenDescrition.Sections(7).Controls("Label18"), "Company Report", xcaption
'        xcaption = cLogUser
'        PutGTotal byGenDescrition.Sections(7).Controls("Label27"), "Company Report", xcaption
'        On Error Resume Next
'        FinanceDE.rsByGenDescription_Grouping.Close
'        FinanceDE.ByGenDescription_Grouping SelectedDate  ', cLogUser
'        byGenDescrition.Show 1
'
'
'    'for Inventory Journal
'    ElseIf WhatJOurNCOde = "IVY" And remark = "Posted" Then
'       PuCaption GenJOurnalPosted.Sections(1).Controls("Label7"), _
'        "Company Report "
'       On Error Resume Next
'       On Error Resume Next
'       FinanceDE.rsInventoryJournal.Close
'       FinanceDE.InventoryJournal SelectedDate
'       IVYJounralposted.Show
'     ElseIf WhatJOurNCOde = "IVY" And remark = "Unpost" Then
'       On Error Resume Next
'       FinanceDE.rsInventoryJournal.Close
'       FinanceDE.InventoryJournal SelectedDate
'       IVYJounralUnpost.Show
'
'    'for Asset Journal
'    ElseIf WhatJOurNCOde = "AST" And remark = "Posted" Then
'       PuCaption GenJOurnalPosted.Sections(1).Controls("Label7"), _
'        "Company Report "
'       On Error Resume Next
'       On Error Resume Next
'       FinanceDE.rsAssetJournal.Close
'       FinanceDE.AssetJournal SelectedDate
'       AssetJOurnAlRep.Show
'     ElseIf WhatJOurNCOde = "AST" And remark = "Unpost" Then
'       On Error Resume Next
'       FinanceDE.rsAssetJournal.Close
'       FinanceDE.AssetJournal SelectedDate
'       'AssetJOurnAlRep.Show
'
'
'    'for payable Journal
'    ElseIf WhatJOurNCOde = "PYB" And remark = "Posted" Then
'       PuCaption PayPosted.Sections(1).Controls("Label7"), _
'       "Company Report "
'       'On Error Resume Next
'       DataEnvironment1.rsPayPost.Close
'       DataEnvironment1.PayPost SelectedDate, WhatJOurNCOde
'       PayPosted.Show
'
'    ElseIf WhatJOurNCOde = "PYB" And remark = "Unpost" Then
'       PuCaption PayJOurnalUnpost.Sections(1).Controls("Label3"), _
'       "Company Report "
'       On Error Resume Next
'       DataEnvironment1.rsPayJounralUnposted.Close
'       If UserRole = "Admin" Then
'         DataEnvironment1.rsPayJounralUnposted.Open "Select * from Payjournal where confirmedDate =" & "'" & SelectedDate & "'" & " and remarks is null order by serialno", constring, adOpenKeyset, adLockPessimistic, adCmdText
'         DataEnvironment1.PayJounralUnposted SelectedDate, cLogUser
'        Else
'         DataEnvironment1.PayJounralUnposted SelectedDate, cLogUser
'        End If
'       PayJOurnalUnpost.Show
'
'
'   'for Bank Journal
'    ElseIf WhatJOurNCOde = "BNK" And remark = "Posted" Then
'      PuCaption GenJOurnalPosted.Sections(1).Controls("Label7"), _
'        "Company Report "
'       On Error Resume Next
'       'FinanceDE.rsGenJournalPosted.Close
'       'FinanceDE.BankJOurnalUnpost SelectedDate, WhatJOurNCOde
'       'BankJOurnalPosted.Show
'     ElseIf WhatJOurNCOde = "BNK" And remark = "Unpost" Then
'       On Error Resume Next
'       FinanceDE.rsBankJournalUnpost.Close
'       If UserRole = "Admin" Then
'         FinanceDE.rsBankJournalUnpost.Open "Select * from Bankjournal where transdate =" & "'" & SelectedDate & "'" & " and remarks is null order by serialno", constring, adOpenKeyset, adLockPessimistic, adCmdText
'         FinanceDE.BankJOurnalUnpost SelectedDate, cLogUser
'        Else
'         FinanceDE.BankJOurnalUnpost SelectedDate, cLogUser
'        End If
'       BankJOurnalUnpost.Show
'
'
'     'for Petty Journal
'    ElseIf WhatJOurNCOde = "PTC" And remark = "Posted" Then
'      PuCaption PettyPosted.Sections(1).Controls("Label7"), _
'        "Company Report "
'       On Error Resume Next
'       DataEnvironment1.rsPettyPost.Close
'       DataEnvironment1.PettyPost SelectedDate, WhatJOurNCOde
'       PettyPosted.Show
'
'     ElseIf WhatJOurNCOde = "PTC" And remark = "Unpost" Then
'       PuCaption RepPettyUnposted.Sections(1).Controls("Label7"), _
'        "Company Report "
'       On Error Resume Next
'       DataEnvironment1.rsPettyUnpost.Close
'       DataEnvironment1.rsPettyUnpost.Requery
'       If UserRole = "Admin" Then
'         DataEnvironment1.rsPettyUnpost.Open "Select * from Pettyjournal where confirmeddate =" & "'" & SelectedDate & "'" & " and postmark is null order by journo", constring, adOpenKeyset, adLockPessimistic, adCmdText
'         DataEnvironment1.PettyUnpost SelectedDate, cLogUser
'        Else
'         DataEnvironment1.PettyUnpost SelectedDate, cLogUser
'        End If
'       RepPettyUnposted.Show
'
'
'    'for cashJournal
    ElseIf WhatJOurNCOde = "CSR" And remark = "Posted" Then
      'PuCaption com_csr_journal_posted.Sections(1).Controls("Label7"),"Company Report "
       
        On Error Resume Next
        dataanu.rscom_csr_journal_posted_Grouping.close
        dataanu.com_csr_journal_posted_Grouping SelectedDate, SelectedDate
        re_csr_Journal_posted.Show
     ElseIf WhatJOurNCOde = "CSR" And remark = "Unpost" Then
       'PuCaption re_csr_jourmal.Sections(1).Controls("Label7"), _
        "Company Report "
        On Error Resume Next
        dataanu.rscom_csr_journal_Grouping.close
        dataanu.com_csr_journal_Grouping SelectedDate, SelectedDate
        re_csr_jourmal.Show
    'End If
    
 'for cashJournal
    ElseIf WhatJOurNCOde = "CSP" And remark = "Posted" Then
      'PuCaption re_csp_Journal_posted.Sections(1).Controls("Label7"), _
        "Company Report "
      On Error Resume Next
      dataanu.rscom_csp_Journal_Posted_Grouping.close
      dataanu.com_csp_Journal_Posted_Grouping SelectedDate, SelectedDate
      re_csp_Journal_posted.Show
     ElseIf WhatJOurNCOde = "CSP" And remark = "Unpost" Then
       'PuCaption re_csp_Journal_posted.Sections(1).Controls("Label7"),
       ' "Company Report "
        On Error Resume Next
        dataanu.rscom_csp_journal_Grouping.close
        dataanu.com_csp_journal_Grouping SelectedDate, SelectedDate
       re_csp_journal.Show
       
' 'for sales return journal
'    ElseIf WhatJOurNCOde = "SRL" And remark = "Posted" Then
'      On Error Resume Next
'      dataanu.rscom_creditnotejournal_posted.Close
'      dataanu.com_creditnotejournal_posted SelectedDate, SelectedDate
'      re_creditnotejournal_posted.Show
'
'     ElseIf WhatJOurNCOde = "SRL" And remark = "Unpost" Then
'        On Error Resume Next
'        dataanu.rscom_creditnotejournal.Close
'        dataanu.com_creditnotejournal SelectedDate, SelectedDate
'       re_creditnotejournal.Show
'
''for sales return journal
'    ElseIf WhatJOurNCOde = "SPA" And remark = "Posted" Then
'      On Error Resume Next
'      dataanu.rscom_debitnotejournal_posted.Close
'      dataanu.com_debitnotejournal_posted SelectedDate, SelectedDate
'      re_debitnoteJournal_Posted.Show
'
'     ElseIf WhatJOurNCOde = "SPA" And remark = "Unpost" Then
'        On Error Resume Next
'        dataanu.rscom_debitnotejournal.Close
'        dataanu.com_debitnoteJournal SelectedDate, SelectedDate
'        re_debitnoteJournal.Show
    End If
End If



CancelPrn:
        X = Err.Number
        If X = 32755 Then
           Exit Sub
          Else
              If WhatJOurNCOde = "SAL" And remark = "Posted" Then
                  PrintSalesJOurnalPosted (SelectedDate)
                ElseIf WhatJOurNCOde = "SAL" And remark = "Unpost" Then
                   Call PrintSalesJOurnal(SelectedDate)
              End If
       End If
End Sub
'Private Sub PutGTotal(lblX As RptLabel, xcaption As String)
'   With lblX
'      .CanGrow = True
'      .caption = xcaption
'   End With
'End Sub
'
'Private Sub PuCaption(lblX As RptLabel)
'   With lblX
'      .CanGrow = True
'      .caption = "As of : " & Me.lvListView.SelectedItem ', "mm/dd/yyyy")
'   End With
'End Sub

Private Sub xREfresh_Click()
Dim rstJournals As New ADODB.Recordset
Dim rstJournals2 As New ADODB.Recordset
Dim rstJournals3 As New ADODB.Recordset
If WhatJOurNCOde = "AST" Then
    'query unposted transactions
    rstJournals.Open "SELECT  TransDate, COUNT(TransDate) AS TotalTRan, SUM(debitamount) AS DrAmt, SUM(CreditAmount) AS CrAmt  " _
             & " From AssetJournal where remarks is null GROUP BY TransDate ORDER BY TransDate", constring, adOpenKeyset, adLockPessimistic, adCmdText
    'query also posted transactions
    rstJournals2.Open "SELECT  TransDate, COUNT(TransDate) AS TotalTRan, SUM(DebitAmount) AS DrAmt, SUM(CreditAmount) AS CrAmt  " _
             & " From AssetJournal where remarks is not null GROUP BY TransDate ORDER BY TransDate", constring, adOpenKeyset, adLockPessimistic, adCmdText
    'query for posted transactions by JN
'    rstJournals3.Open "SELECT COUNT(DISTINCT SerialNo) AS cTOtal from assetJournal WHERE (Remarks IS NULL) where remarks is not null GROUP BY TransDate ORDER BY TransDate", constring, adOpenKeyset, adLockPessimistic, adCmdText
             
             
ElseIf WhatJOurNCOde = "SAL" Then
    'query unposted transactions
    rstJournals.Open "SELECT  TransDate, COUNT(TransDate) AS TotalTRan, SUM(tradercvble+tradeDiscAmt) AS DrAmt, SUM(GrossSales+Vat+transpoCharge+surtaxamt) AS CrAmt  " _
             & " From SalesJournal where remarks is null GROUP BY TransDate ORDER BY TransDate", constring, adOpenKeyset, adLockPessimistic, adCmdText
    'query also posted transactions
    rstJournals2.Open "SELECT  TransDate, COUNT(TransDate) AS TotalTRan, SUM(tradercvble+tradeDiscAmt) AS DrAmt, SUM(GrossSales+Vat+transpoCharge+surtaxamt) AS CrAmt  " _
             & " From SalesJournal where remarks is not null GROUP BY TransDate ORDER BY TransDate", constring, adOpenKeyset, adLockPessimistic, adCmdText
             
ElseIf WhatJOurNCOde = "IVY" Then
    'query unposted transactions
    rstJournals.Open "SELECT  TransDate, COUNT(TransDate) AS TotalTRan, SUM(DebitAmount) AS DrAmt, SUM(CreditAmount) AS CrAmt  " _
             & " From InventoryJournal where remarks is null GROUP BY TransDate ORDER BY TransDate", constring, adOpenKeyset, adLockPessimistic, adCmdText
    'query also posted transactions
    rstJournals2.Open "SELECT  TransDate, COUNT(TransDate) AS TotalTRan, SUM(DebitAmount) AS DrAmt, SUM(CreditAmount) AS CrAmt  " _
             & " From InventoryJournal where remarks is not null GROUP BY TransDate ORDER BY TransDate", constring, adOpenKeyset, adLockPessimistic, adCmdText
             
ElseIf WhatJOurNCOde = "GEN" Then
    'query unposted transactions
    rstJournals.Open "SELECT  TransDate, COUNT(TransDate) AS TotalTRan, SUM(DebitAmount) AS DrAmt, SUM(CreditAmount) AS CrAmt  " _
             & " From GenJournalTrans where remarks is null GROUP BY TransDate ORDER BY TransDate", constring, adOpenKeyset, adLockPessimistic, adCmdText
    'query also posted transactions
    rstJournals2.Open "SELECT  TransDate, COUNT(TransDate) AS TotalTRan, SUM(DebitAmount) AS DrAmt, SUM(CreditAmount) AS CrAmt  " _
             & " From GenJournalTrans where remarks is not null GROUP BY TransDate ORDER BY TransDate", constring, adOpenKeyset, adLockPessimistic, adCmdText
             
             
ElseIf WhatJOurNCOde = "PYB" Then
    'query unposted transactions
    rstJournals.Open "SELECT  confirmedDate, COUNT(confirmedDate) AS TotalTRan, SUM(DbAmount) AS DrAmt, SUM(CrAmount) AS CrAmt  " _
             & " From PayJournal where Status = 'Unposted'  GROUP BY confirmedDate ORDER BY confirmedDate", constring, adOpenKeyset, adLockPessimistic, adCmdText
    'query also posted transactions
    rstJournals2.Open "SELECT  confirmedDate, COUNT(confirmedDate) AS TotalTRan, SUM(DbAmount) AS DrAmt, SUM(CrAmount) AS CrAmt  " _
             & " From PayJournal where status ='Posted' GROUP BY confirmedDate ORDER BY confirmedDate", constring, adOpenKeyset, adLockPessimistic, adCmdText
                          
ElseIf WhatJOurNCOde = "CSP" Then
    'query unposted transactions
    rstJournals.Open "SELECT  TransDate, COUNT(TransDate) AS TotalTRan, SUM(DebitAmount) AS DrAmt, SUM(CreditAmount) AS CrAmt  " _
             & " From cashjournal where remarks is null and trantype = 'P' GROUP BY TransDate ORDER BY TransDate", constring, adOpenKeyset, adLockPessimistic, adCmdText
    'query also posted transactions
    rstJournals2.Open "SELECT  TransDate, COUNT(TransDate) AS TotalTRan, SUM(DebitAmount) AS DrAmt, SUM(CreditAmount) AS CrAmt  " _
             & " From cashjournal where remarks is not null and trantype = 'P' GROUP BY TransDate ORDER BY TransDate", constring, adOpenKeyset, adLockPessimistic, adCmdText


ElseIf WhatJOurNCOde = "CSR" Then
    'query unposted transactions
    rstJournals.Open "SELECT  TransDate, COUNT(TransDate) AS TotalTRan, SUM(DebitAmount) AS DrAmt, SUM(CreditAmount) AS CrAmt  " _
             & " From cashjournal where remarks is null and trantype = 'R' GROUP BY TransDate ORDER BY TransDate", constring, adOpenKeyset, adLockPessimistic, adCmdText
    'query also posted transactions
    rstJournals2.Open "SELECT  TransDate, COUNT(TransDate) AS TotalTRan, SUM(DebitAmount) AS DrAmt, SUM(CreditAmount) AS CrAmt  " _
             & " From cashjournal where remarks is not null and trantype = 'R' GROUP BY TransDate ORDER BY TransDate", constring, adOpenKeyset, adLockPessimistic, adCmdText
             
ElseIf WhatJOurNCOde = "SRL" Then
    'query unposted transactions
    rstJournals.Open "SELECT  TransDate, COUNT(TransDate) AS TotalTRan, SUM(DebitAmount) AS DrAmt, SUM(CreditAmount) AS CrAmt  " _
             & " From creditnote where remarks is null and trantype = 'SRL' GROUP BY TransDate ORDER BY TransDate", constring, adOpenKeyset, adLockPessimistic, adCmdText
    'query also posted transactions
    rstJournals2.Open "SELECT  TransDate, COUNT(TransDate) AS TotalTRan, SUM(DebitAmount) AS DrAmt, SUM(CreditAmount) AS CrAmt  " _
             & " From creditnote where remarks is not null and trantype = 'SRL' GROUP BY TransDate ORDER BY TransDate", constring, adOpenKeyset, adLockPessimistic, adCmdText

ElseIf WhatJOurNCOde = "SPA" Then
    'query unposted transactions
    rstJournals.Open "SELECT  TransDate, COUNT(TransDate) AS TotalTRan, SUM(DebitAmount) AS DrAmt, SUM(CreditAmount) AS CrAmt  " _
             & " From debitnote where remarks is null and trantype = 'SPA' GROUP BY TransDate ORDER BY TransDate", constring, adOpenKeyset, adLockPessimistic, adCmdText
    'query also posted transactions
    rstJournals2.Open "SELECT  TransDate, COUNT(TransDate) AS TotalTRan, SUM(DebitAmount) AS DrAmt, SUM(CreditAmount) AS CrAmt  " _
             & " From debitnote  where remarks is not null and trantype = 'SPA' GROUP BY TransDate ORDER BY TransDate", constring, adOpenKeyset, adLockPessimistic, adCmdText

ElseIf WhatJOurNCOde = "PTC" Then
    'query unposted transactions
    rstJournals.Open "SELECT  confirmeddate, COUNT(confirmeddate) AS TotalTRan, SUM(DbAmount) AS DrAmt, SUM(CrAmount) AS CrAmt  " _
             & " From pettyjournal where postMark is null  GROUP BY confirmeddate ORDER BY confirmeddate", constring, adOpenKeyset, adLockPessimistic, adCmdText
    'query also posted transactions
     rstJournals2.Open "SELECT  confirmeddate, COUNT(confirmeddate) AS TotalTRan, SUM(confirmeddate) AS DrAmt, SUM(confirmeddate) AS CrAmt  " _
             & " From pettyjournal where postMark is not null  GROUP BY Datex ORDER BY confirmeddate", constring, adOpenKeyset, adLockPessimistic, adCmdText
             
ElseIf WhatJOurNCOde = "BNK" Then
    'query unposted transactions
    rstJournals.Open "SELECT  TransDate, COUNT(TransDate) AS TotalTRan, SUM(DebitAmount) AS DrAmt, SUM(CreditAmount) AS CrAmt  " _
             & " From BankJOurnal where remarks is null GROUP BY TransDate ORDER BY TransDate", constring, adOpenKeyset, adLockPessimistic, adCmdText
    'query also posted transactions
    rstJournals2.Open "SELECT  TransDate, COUNT(TransDate) AS TotalTRan, SUM(DebitAmount) AS DrAmt, SUM(CreditAmount) AS CrAmt  " _
             & " From BankJOurnal where remarks is not null GROUP BY TransDate ORDER BY TransDate", constring, adOpenKeyset, adLockPessimistic, adCmdText
             
Else
  Exit Sub
End If



'display also posted days
Me.lvListView.ListItems.clear
Do Until rstJournals2.EOF = True
    If WhatJOurNCOde = "PYB" Then
        trDate = Format(rstJournals2!confirmeddate, "dd/mm/yyyy")
     ElseIf WhatJOurNCOde = "PTC" Then
        trDate = Format(rstJournals2!confirmeddate, "dd/mm/yyyy")
     Else
        trDate = Format(rstJournals2!TRansDate, "dd/mm/yyyy")
    End If
    Set MItem = Me.lvListView.ListItems.Add(, , trDate, 3, 3)
    MItem.SubItems(1) = rstJournals2!TotalTRan
    MItem.SubItems(2) = FormatNumber(rstJournals2!Dramt, 2, vbTrue, vbTrue, vbTrue)
    MItem.SubItems(3) = FormatNumber(rstJournals2!CrAmt, 2, vbTrue, vbTrue, vbTrue)
    MItem.SubItems(4) = "Posted"
    rstJournals2.MoveNext
    DoEvents
Loop

'display unposted days
Do Until rstJournals.EOF = True
    If WhatJOurNCOde = "PYB" Then
        trDate = Format(rstJournals!confirmeddate, "dd/mm/yyyy")
     ElseIf WhatJOurNCOde = "PTC" Then
        trDate = Format(rstJournals!confirmeddate, "dd/mm/yyyy")
     Else
        trDate = Format(rstJournals!TRansDate, "dd/mm/YYYy")
    End If
    Set MItem = Me.lvListView.ListItems.Add(, , trDate, 4, 4)
    MItem.SubItems(1) = rstJournals!TotalTRan
    MItem.SubItems(2) = FormatNumber(rstJournals!Dramt, 2, vbTrue, vbTrue, vbTrue)
    MItem.SubItems(3) = FormatNumber(rstJournals!CrAmt, 2, vbTrue, vbTrue, vbTrue)
    MItem.SubItems(4) = "Unpost"
    rstJournals.MoveNext
    DoEvents
Loop
End Sub

Private Sub xSAles_Click()
DownLoadinvoices.Show
End Sub

Private Sub xSI_Click()

Me.lvListView.View = lvwSmallIcon
Me.xLI.Checked = False
Me.xSI.Checked = True
Me.xViewList.Checked = False
Me.xDetails.Checked = False
End Sub

Private Sub xuser_Click()
NewUser.Show 1
End Sub

Private Sub xViewList_Click()
Me.lvListView.View = lvwList
Me.xLI.Checked = False
Me.xSI.Checked = False
Me.xViewList.Checked = True
Me.xDetails.Checked = False

End Sub

Private Sub xViewUser_Click()
ViewLogUser.Show
End Sub

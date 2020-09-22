VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmreceiptagainsinvoice 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cash Receipts Vs Invoice"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4950
   ControlBox      =   0   'False
   Icon            =   "frmreceiptagainsinvoice.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   4950
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optdate 
      Caption         =   "By date ÈæÇÓØÉ ÇáÊÇÑíÎ"
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   120
      Width           =   1935
   End
   Begin VB.OptionButton optnumber 
      Caption         =   "By Number ÈæÇÓØÉ ÇáÑÞã "
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   360
      Value           =   -1  'True
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Caption         =   "Enter the Receipt Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4695
      Begin VB.CommandButton cmdok 
         Caption         =   "&Ok"
         Height          =   350
         Left            =   1800
         TabIndex        =   11
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton cmdclose 
         Caption         =   "C&lose"
         Height          =   350
         Left            =   2880
         TabIndex        =   10
         Top             =   960
         Width           =   975
      End
      Begin MSComCtl2.DTPicker dateto 
         Height          =   300
         Left            =   2640
         TabIndex        =   9
         Top             =   600
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         Format          =   52035585
         CurrentDate     =   37573
      End
      Begin MSComCtl2.DTPicker datefrom 
         Height          =   300
         Left            =   2640
         TabIndex        =   8
         Top             =   240
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         Format          =   52035585
         CurrentDate     =   37573
      End
      Begin MSMask.MaskEdBox txtfrom 
         Height          =   300
         Left            =   2640
         TabIndex        =   0
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   8
         Mask            =   "########"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtto 
         Height          =   300
         Left            =   2640
         TabIndex        =   2
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   8
         Mask            =   "########"
         PromptChar      =   " "
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Ending Number"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   650
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Starting Number "
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   300
         Width           =   1335
      End
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Choose Option ÇÎÊíÇÑ ÇáæÙÇÆÝ"
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
      Left            =   150
      TabIndex        =   7
      Top             =   45
      Width           =   1575
   End
End
Attribute VB_Name = "frmreceiptagainsinvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Option1_Click()

End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdok_Click()
  If optnumber.Value = True Then
    If Trim(txtfrom.Text) = "" Or Trim(txtto.Text) = "" Then
        MsgBox "Please Enter The Receipt Number", vbInformation, "Empty Receipt Number"
        txtfrom.SetFocus
        Exit Sub
    End If

        If Val(Trim(txtfrom.Text)) > Val(Trim(txtto.Text)) Then
            MsgBox "You Enterd the Receipt Number in Wrong Format", vbInformation, "Invalid Number"
            txtfrom.SetFocus
            Exit Sub
        End If
        Dim from As Long
        Dim too As Long
        from = Val(Trim(txtfrom.Text))
        too = Val(Trim(txtto.Text))
        On Error Resume Next
        dataanu.rscominvoice_Grouping.Close
        On Error GoTo 0
        dataanu.cominvoice_Grouping from, too
        rereceiptvsinvoice.Show 1
    End If
    
    
    If optdate.Value = True Then
        If Format(datefrom.Value, "mm/dd/yyyy") > Format(dateto.Value, "mm/dd/yyyy") Then
            MsgBox "Plese check Dates Should be From and to Method", vbInformation, "Incorrect Date Format"
            Exit Sub
        End If
         On Error Resume Next
        dataanu.rscominvioce2_Grouping.Close
        On Error GoTo 0
        dataanu.cominvioce2_Grouping datefrom.Value, dateto.Value
        rereceiptvsinvoice2.Show 1
    End If
End Sub

Private Sub Form_Load()
datefrom.Value = "01/01/" & DatePart("yyyy", Date)
dateto.Value = Date
End Sub

Private Sub optdate_Click()
txtfrom.Visible = False
txtto.Visible = False
datefrom.Visible = True
dateto.Visible = True
Label1.caption = "Starting Date"
Label2.caption = "Ending Date"
Frame1.caption = "Enter the Date"
End Sub

Private Sub optnumber_Click()
txtfrom.Visible = True
txtto.Visible = True
txtfrom.Text = ""
txtto.Text = ""
datefrom.Visible = False
dateto.Visible = False
Label1.caption = "Starting Number"
Label2.caption = "Ending Number"
Frame1.caption = "Enter the Receipt Number"
txtfrom.SetFocus
End Sub

Private Sub txtfrom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Trim(txtfrom.Text) <> "" Then
    txtto.SetFocus
End If
End Sub

Private Sub txtto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Trim(txtto.Text) <> "" Then
    cmdok_Click
End If
End Sub

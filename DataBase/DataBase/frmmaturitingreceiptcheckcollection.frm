VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmmaturitingreceiptcheckcollection 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maturiting Collection Check."
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4230
   Icon            =   "frmmaturitingreceiptcheckcollection.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   4230
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      Begin VB.CommandButton cmdclose 
         Caption         =   "&Close"
         Height          =   350
         Left            =   3000
         TabIndex        =   5
         Top             =   960
         Width           =   765
      End
      Begin VB.CommandButton cmdactivity 
         Caption         =   "&Preview"
         Height          =   350
         Left            =   3000
         TabIndex        =   1
         Top             =   600
         Width           =   765
      End
      Begin MSComCtl2.DTPicker dtpfrom 
         Height          =   315
         Left            =   1200
         TabIndex        =   2
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Format          =   67174401
         CurrentDate     =   37530
      End
      Begin MSComCtl2.DTPicker dtpto 
         Height          =   315
         Left            =   1200
         TabIndex        =   6
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Format          =   67174401
         CurrentDate     =   37621
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Çáí "
         Height          =   195
         Left            =   360
         TabIndex        =   7
         Top             =   1080
         Width           =   525
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Choose The Period ÇÎÊíÇÑ ÇáÝÊÑÉ "
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
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From ãä "
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmmaturitingreceiptcheckcollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdactivity_Click()
If dtpfrom.Value > dtpto.Value Then
    MsgBox "Please Enter Your Date Correctly", vbInformation, "Disorder Date"
    Exit Sub
End If

dataanu.comcollectedreceiptcheck_Grouping Format(dtpto.Value, "mm/dd/yyyy"), Format(dtpfrom.Value, "mm/dd/yyyy")
FormatLabelcheckdate recheckcollections.Sections(2).Controls("label20"), _
        "Company Report "
FormatLabelcheckdate2 recheckcollections.Sections(2).Controls("label19"), _
        "Company Report "
recheckcollections.Show 1

End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub Form_Load()
dtpfrom.Value = Date
dtpto.Value = DateAdd("d", 7, dtpfrom.Value)
End Sub

Private Sub FormatLabelcheckdate(lblX As RptLabel, caption As String)
      lblX.caption = Format(dtpto.Value, "dd/mm/yyyy")
End Sub
Private Sub FormatLabelcheckdate2(lblX As RptLabel, caption As String)
      lblX.caption = Format(dtpfrom.Value, "dd/mm/yyyy") & "  To "
End Sub


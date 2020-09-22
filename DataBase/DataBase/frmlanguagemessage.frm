VERSION 5.00
Begin VB.Form frmlanguagemessage 
   Appearance      =   0  'Flat
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Language Option"
   ClientHeight    =   885
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3810
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   885
   ScaleWidth      =   3810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Arabic ÚÑÈí "
      Default         =   -1  'True
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "English ÇäÌáíÒí "
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmlanguagemessage.frx":0000
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ÇáÑÌÇÁ ÇÎÊíÇÑ áÛÉ ØÈÇÚÉ ÇáãÓÊäÏ "
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
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   1740
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please Choose Your Language To Print"
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
      Left            =   720
      TabIndex        =   0
      Top             =   0
      Width           =   2835
   End
End
Attribute VB_Name = "frmlanguagemessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmpaymentvou.langopt.Value = 0
frmrecieptvou.langopt.Value = 0
Unload Me
End Sub

Private Sub Command2_Click()
 frmpaymentvou.langopt.Value = 1
 frmrecieptvou.langopt.Value = 1
 Unload Me
End Sub


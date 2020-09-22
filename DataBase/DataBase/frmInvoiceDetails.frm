VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmInvoiceDetails 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Invoice Details"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   5805
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Height          =   450
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   5775
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   4320
         TabIndex        =   5
         Top             =   120
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   120
         Width           =   1455
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   2880
         TabIndex        =   3
         Top             =   120
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   0
         TabIndex        =   2
         Top             =   120
         Width           =   1455
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   3201
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483634
      Appearance      =   0
      NumItems        =   0
   End
End
Attribute VB_Name = "frmInvoiceDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Load()

Dim Pundamora
Dim Sevena
Dim CON1 As ADODB.Connection
Set CON1 = New ADODB.Connection
Dim RstInv As New ADODB.Recordset
conStr = "Provider=MSDASQL;DSN=Finance;UID=sa; PWD=;"

CON1.Open conStr

Set AR = Me.ListView1.ColumnHeaders.Add(, , "Invoice No")
Set AR = Me.ListView1.ColumnHeaders.Add(, , "Invoice Date")
Set AR = Me.ListView1.ColumnHeaders.Add(, , "Due Date")
Set AR = Me.ListView1.ColumnHeaders.Add(, , "Invoice Amount")
Sevena = FrmPayableSetup.txtSerialNo.Text

Pundamora = "select * from PayInvoiceDetails where SerialNo = " & "'" & Sevena & "'" & ""
RstInv.Open Pundamora, CON1, adOpenDynamic, adLockOptimistic

If RstInv.EOF = False Then
RstInv.MoveFirst
End If
  
     While RstInv.EOF = False
     Set MItem = Me.ListView1.ListItems.Add(, , Format(RstInv!InvNo))
     MItem.SubItems(1) = Format(RstInv!InvDate, "dd/mm/yyyy")
     MItem.SubItems(2) = Format(RstInv!duedate, "dd/mm/yyyy")
     MItem.SubItems(3) = Format(RstInv!invAmt)
     RstInv.MoveNext
     Wend



End Sub

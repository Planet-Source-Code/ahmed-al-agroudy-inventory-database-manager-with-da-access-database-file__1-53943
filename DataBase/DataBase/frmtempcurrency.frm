VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmtempcurrency 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Currency  ÇáÚãáÉ"
   ClientHeight    =   2715
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   7815
   Icon            =   "frmtempcurrency.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   7815
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmtempcurrency 
      Height          =   2700
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7815
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   1920
         Top             =   1440
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483639
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmtempcurrency.frx":0582
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmtempcurrency.frx":09D4
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmtempcurrency.frx":0E26
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmtempcurrency.frx":1278
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmtempcurrency.frx":180A
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmddelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   350
         Left            =   5760
         TabIndex        =   20
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox txtRate 
         Height          =   285
         Left            =   4080
         TabIndex        =   13
         Top             =   1320
         Width           =   1320
      End
      Begin VB.CommandButton cmdclose 
         Caption         =   "Cl&ose"
         Height          =   350
         Left            =   6720
         TabIndex        =   12
         Top             =   2160
         Width           =   855
      End
      Begin VB.CommandButton cmdnewrecord 
         Caption         =   "&New"
         Height          =   350
         Left            =   2880
         TabIndex        =   11
         Top             =   2160
         Width           =   855
      End
      Begin VB.CommandButton cmdsave 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   350
         Left            =   4800
         TabIndex        =   10
         Top             =   2160
         Width           =   855
      End
      Begin VB.CommandButton cmdcancel 
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   350
         Left            =   3840
         TabIndex        =   9
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox txtCurrency 
         Height          =   285
         Left            =   4080
         TabIndex        =   4
         Top             =   255
         Width           =   1035
      End
      Begin VB.TextBox txtlongdetails 
         Height          =   285
         Left            =   4080
         ScrollBars      =   3  'Both
         TabIndex        =   3
         Top             =   960
         Width           =   2655
      End
      Begin VB.ComboBox comaccountnumber 
         Height          =   315
         Left            =   4080
         TabIndex        =   2
         Text            =   "comaccountnumber"
         Top             =   1680
         Width           =   2655
      End
      Begin VB.ComboBox comtype 
         Height          =   315
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   1095
      End
      Begin MSComctlLib.TreeView treecurrency 
         Height          =   2535
         Left            =   50
         TabIndex        =   21
         Top             =   120
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   4471
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "ImageList1"
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ÑÞã ÇáÍÓÇÈ"
         Height          =   195
         Index           =   7
         Left            =   6930
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   1740
         Width           =   795
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ÇáÇÓã ÇáßÇãá"
         Height          =   195
         Index           =   6
         Left            =   6840
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   1020
         Width           =   825
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÇáäæÚ"
         Height          =   195
         Index           =   5
         Left            =   7320
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   660
         Width           =   360
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ÇáÚãáÉ"
         Height          =   195
         Index           =   3
         Left            =   7335
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ÓÚÑ ÇáÊÍæíá"
         Height          =   195
         Index           =   2
         Left            =   6855
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   1380
         Width           =   855
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rate:"
         Height          =   195
         Index           =   1
         Left            =   2880
         TabIndex        =   14
         Top             =   1440
         Width           =   390
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Currency"
         Height          =   195
         Index           =   0
         Left            =   2880
         TabIndex        =   8
         Top             =   300
         Width           =   630
      End
      Begin VB.Label lblFieldLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         Height          =   195
         Index           =   4
         Left            =   2895
         TabIndex        =   7
         Top             =   720
         Width           =   360
      End
      Begin VB.Label lblFieldLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Full Name"
         Height          =   195
         Index           =   9
         Left            =   2880
         TabIndex        =   6
         Top             =   1080
         Width           =   705
      End
      Begin VB.Label lblFieldLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Accountnumber"
         Height          =   195
         Index           =   11
         Left            =   2880
         TabIndex        =   5
         Top             =   1800
         Width           =   1125
      End
   End
   Begin VB.Menu changecurrency 
      Caption         =   "File"
      Begin VB.Menu addnew 
         Caption         =   "Add New"
      End
      Begin VB.Menu dash 
         Caption         =   "-"
      End
      Begin VB.Menu setasdefault 
         Caption         =   "Set As Default  ÇáÚãáÉ ÇáãÓÊÚãáÉ "
         Checked         =   -1  'True
      End
      Begin VB.Menu dash1 
         Caption         =   "-"
      End
      Begin VB.Menu close 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "frmtempcurrency"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xtable As String
Dim sqltable As Boolean
Dim myclass As New HabitatClass
Dim constring As String

Dim nodex As Node

Dim rectemp As New ADODB.Recordset
Dim contemp As New ADODB.Connection
Dim Status As String


Private Sub addnew_Click()
cmdnewrecord_Click
End Sub

Private Sub close_Click()
cmdclose_Click
End Sub

Private Sub cmdcancel_Click()
cmdnewrecord.Enabled = True
cmdcancel.Enabled = False
cmdclose.Enabled = True
cmdsave.Enabled = False
cmddelete.Enabled = False
txtCurrency.Enabled = True

txtCurrency.Text = ""
comtype.ListIndex = 0
txtlongdetails.Text = ""
txtRate.Text = ""
comaccountnumber.ListIndex = 0
Status = "nodate" ' just assign for nothing

End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub Command1_Click()
X = treecurrency.Nodes.Count
For X = 1 To treecurrency.Nodes.Count
treecurrency.Nodes.Remove X
Next
treecurrency.Nodes.Clear
End Sub

Private Sub cmddelete_Click()
If MsgBox("Are You Sure You Want to Delete", vbYesNo, "Conformation") = vbYes Then
    Dim recupdate As New ADODB.Recordset
    Dim conupdate As New ADODB.Connection
    conupdate.Open "dsn=finance;uid=sa;"
    recupdate.Open "delete from currencytable where currency = '" & Trim(txtCurrency.Text) & "'", conupdate, adOpenKeyset, adLockOptimistic
    conupdate.Close
    MsgBox "This currency Deleted", vbInformation, "Deleted"
    Call prcrecalltreeview
    cmdcancel_Click
End If
End Sub

Private Sub cmdnewrecord_Click()
txtCurrency.Text = ""
comtype.ListIndex = 0
txtlongdetails.Text = ""
txtRate.Text = ""
comaccountnumber.ListIndex = 0
cmdnewrecord.Enabled = False
cmdcancel.Enabled = True
cmdclose.Enabled = False
cmdsave.Enabled = True
Status = "new"

End Sub

Private Sub cmdsave_Click()
If Trim(txtCurrency.Text) = "" Or _
     Trim(txtRate.Text) = "" Or _
     Trim(txtlongdetails.Text) = "" Or _
     Trim(comaccountnumber.Text) = "" Then
   MsgBox "Please Fill All The Details to Save", vbInformation, "Found Empty"
   Exit Sub
End If
   
If Val(txtRate.Text) < 0 Then
    MsgBox "You Must Enter Currency Rate Greater Than Zero", vbInformation, "Invalid Rate"
    txtRate.SetFocus
    Exit Sub
End If

Dim recaddnew As New ADODB.Recordset
Dim conaddnew As New ADODB.Connection
conaddnew.Open "Dsn=finance;uid=sa;"

If Status = "new" Then
    recaddnew.Open "INSERT INTO currencytable VALUES ('" & _
                    Trim(txtCurrency.Text) & _
                    "'," & Trim(txtRate.Text) & _
                    ",'" & Date & _
                    "',NULL,'" & Trim(comtype.Text) & _
                    "'," & 0 & _
                    "," & 0 & _
                    "," & 0 & _
                    "," & 0 & _
                    ",'" & Trim(txtlongdetails.Text) & _
                    "','" & Date & _
                    "','" & Trim(comaccountnumber.Text) & "')", conaddnew, adOpenKeyset, adLockOptimistic
End If

If Status = "edit" Then
    If Trim(comtype.Text) = "Card" Then
        recaddnew.Open "update currencytable set rate ='" & Trim(txtRate.Text) & _
                    "',status = NULL,latestupdate=getdate(),Detail='" & Trim(comtype.Text) & _
                    " ',longdetails ='" & Trim(txtlongdetails.Text) & _
                    "',accountnumber='" & Trim(comaccountnumber.Text) & "' where currency = '" & Trim(txtCurrency.Text) & "'", conaddnew, adOpenKeyset, adLockOptimistic
    Else
        recaddnew.Open "update currencytable set rate ='" & Trim(txtRate.Text) & _
                    "',latestupdate=getdate(),Detail='" & Trim(comtype.Text) & _
                    " ',longdetails ='" & Trim(txtlongdetails.Text) & _
                    "',accountnumber='" & Trim(comaccountnumber.Text) & "' where currency = '" & Trim(txtCurrency.Text) & "'", conaddnew, adOpenKeyset, adLockOptimistic
    End If
End If

txtCurrency.Enabled = True

conaddnew.Close
cmdnewrecord.Enabled = True
cmdcancel.Enabled = False
cmdclose.Enabled = True
cmdsave.Enabled = False
Status = "nodata" ' just assign for nothing
Call prcrecalltreeview


End Sub

Private Sub Form_Load()
comtype.AddItem "Cash"
comtype.AddItem "Card"
Dim confinance As New ADODB.Connection
constring = "Dsn=finance;pwd=;"
sqltable = True
xtable = "select * from financemaster"

myclass.GetTables rectemp, confinance, xtable, constring, sqltable


While rectemp.EOF = False
    comaccountnumber.AddItem Trim(rectemp!AccountCode)
    rectemp.MoveNext
Wend
rectemp.Close

Set nodex = treecurrency.Nodes.Add(, , "a", "Currency", 2)
Set nodex = treecurrency.Nodes.Add("a", tvwChild, "a1", "Cash", 2)
Set nodex = treecurrency.Nodes.Add("a", tvwChild, "a2", "Credit Card", 2)

contemp.Open "Dsn=finance;uid=sa"
rectemp.Open "select * from currencytable where detail = 'Cash'", contemp, adOpenKeyset, adLockOptimistic

While rectemp.EOF = False
i = i + 1
    If Trim(rectemp!Status) = "default" Then
        Set nodex = treecurrency.Nodes.Add("a1", tvwChild, "ac1" & i, Trim(rectemp!Currency) & "  ", 5)
    Else
        Set nodex = treecurrency.Nodes.Add("a1", tvwChild, "ac1" & i, Trim(rectemp!Currency) & "  ", 4)
    End If
    
    rectemp.MoveNext
Wend
rectemp.Close

i = 0
rectemp.Open "select * from currencytable where detail = 'Card'", contemp, adOpenKeyset, adLockOptimistic

While rectemp.EOF = False
i = i + 1
    Set nodex = treecurrency.Nodes.Add("a2", tvwChild, "ac2" & i, Trim(rectemp!Currency) & "  ", 1)
    rectemp.MoveNext
Wend
rectemp.Close
End Sub


Private Sub setasdefault_Click()
If treecurrency.Nodes.Count > 0 Then
    'make all currency as not default
    'MsgBox Trim(treecurrency.SelectedItem.Parent.Text)
    If Trim(treecurrency.SelectedItem.Parent.Text) <> "Cash" Then
        MsgBox "You Can Not Make Credit Card As Default", vbInformation, "Invalid Selection"
        Exit Sub
    End If
    
    On Error Resume Next
    contemp.Open "Dsn=finance;uid=sa;'"
    On Error GoTo 0
    
    rectemp.Open "UPDATE currencytable SET Status = NULL", contemp, adOpenKeyset, adLockOptimistic
    'update the table for default
    rectemp.Open "update currencytable set status = 'default' where currency ='" & Trim(treecurrency.SelectedItem.Text) & "'", contemp, adOpenKeyset, adLockOptimistic
    contemp.Close
    'clear the treeview
    Call prcrecalltreeview
End If
End Sub

Private Sub treecurrency_Click()
If treecurrency.Nodes.Count > 0 Then
    On Error Resume Next
    selectedtext = treecurrency.SelectedItem.Text
    On Error GoTo 0
    rectemp.Open "select * from currencytable where currency = '" & selectedtext & "'", contemp, adOpenKeyset, adLockOptimistic
    If rectemp.BOF = False Then
        If Trim(rectemp!longdetails) <> "" Then
            treecurrency.ToolTipText = Trim(rectemp!longdetails)
        End If
    End If
    rectemp.Close
End If
End Sub
Private Sub treecurrency_DblClick()
If treecurrency.Nodes.Count > 0 Then
    On Error Resume Next
    selectedtext = treecurrency.SelectedItem.Text
    On Error GoTo 0
    rectemp.Open "select * from currencytable where currency = '" & selectedtext & "'", contemp, adOpenKeyset, adLockOptimistic
    If rectemp.BOF = False Then

        If Trim(rectemp!longdetails) <> "" Then
            treecurrency.ToolTipText = Trim(rectemp!longdetails)
        End If

        txtCurrency.Text = Trim(rectemp!Currency)
        txtRate.Text = Trim(rectemp!Rate)
        If Trim(rectemp!detail) = "Cash" Then
            comtype.ListIndex = 0
        Else
            comtype.ListIndex = 1
        End If
       txtlongdetails = Trim(rectemp!longdetails)
       comaccountnumber.Text = Trim(rectemp!accountnumber)

       Status = "edit"
       cmdnewrecord.Enabled = False
       cmdclose.Enabled = False
       cmdsave.Enabled = True
       cmdcancel.Enabled = True
       cmddelete.Enabled = True
       txtCurrency.Enabled = False
    Else
         cmdnewrecord.Enabled = True
       cmdclose.Enabled = True
       cmdsave.Enabled = False
       cmdcancel.Enabled = False
    End If
    rectemp.Close
End If
End Sub

Private Sub treecurrency_Collapse(ByVal Node As MSComctlLib.Node)
    Node.Image = 2
End Sub



Private Sub treecurrency_Expand(ByVal Node As MSComctlLib.Node)
    Node.Image = 3
End Sub

Public Sub prcrecalltreeview()
    treecurrency.Nodes.Clear
    'adding all currency again
    Set nodex = treecurrency.Nodes.Add(, , "a", "Currency", 2)
    Set nodex = treecurrency.Nodes.Add("a", tvwChild, "a1", "Cash", 2)
    Set nodex = treecurrency.Nodes.Add("a", tvwChild, "a2", "Credit Card", 2)
    On Error Resume Next
    contemp.Open "Dsn=finance;uid=sa"
    On Error GoTo 0
    rectemp.Open "select * from currencytable where detail = 'Cash'", contemp, adOpenKeyset, adLockOptimistic
    
    While rectemp.EOF = False
    i = i + 1
        If Trim(rectemp!Status) = "default" Then
            Set nodex = treecurrency.Nodes.Add("a1", tvwChild, "ac1" & i, Trim(rectemp!Currency) & "  ", 5)
        Else
            Set nodex = treecurrency.Nodes.Add("a1", tvwChild, "ac1" & i, Trim(rectemp!Currency) & "  ", 4)
        End If
        rectemp.MoveNext
    Wend
    rectemp.Close
    
    i = 0
    rectemp.Open "select * from currencytable where detail = 'Card'", contemp, adOpenKeyset, adLockOptimistic
    While rectemp.EOF = False
    i = i + 1
        Set nodex = treecurrency.Nodes.Add("a2", tvwChild, "ac2" & i, Trim(rectemp!Currency) & "  ", 1)
        rectemp.MoveNext
    Wend
    rectemp.Close

End Sub

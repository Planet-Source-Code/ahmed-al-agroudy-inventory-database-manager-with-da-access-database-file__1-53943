VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form SCMenu 
   Caption         =   "Form2"
   ClientHeight    =   2115
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4080
   Icon            =   "SCMenu.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   2115
   ScaleWidth      =   4080
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1200
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   4800
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   5280
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   5280
      Width           =   855
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SCMenu.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SCMenu.frx":075C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SCMenu.frx":0BAE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu xMain 
      Caption         =   "Main     ÑÆíÓí"
      Begin VB.Menu xview 
         Caption         =   "View      ÚÑÖ"
         Begin VB.Menu Li 
            Caption         =   "Large Icons           ÃíÞæäÇÊ ßÈíÑÉ"
            Checked         =   -1  'True
         End
         Begin VB.Menu SI 
            Caption         =   "Small Icons          ÃíÞæäÇÊ ÕÛíÑÉ"
         End
         Begin VB.Menu List 
            Caption         =   "List                                   ÚÑÖ"
         End
         Begin VB.Menu details 
            Caption         =   "Details                           ÊÝÇÕíá"
         End
      End
      Begin VB.Menu Open 
         Caption         =   "Modify/Delete      ÊÚÏíá       "
      End
      Begin VB.Menu Post 
         Caption         =   "&Post..."
      End
      Begin VB.Menu xrefresh 
         Caption         =   "Refresh               ÊÍÏíË"
         Enabled         =   0   'False
      End
      Begin VB.Menu bar1 
         Caption         =   "-"
      End
      Begin VB.Menu find 
         Caption         =   "Find..."
      End
      Begin VB.Menu CB 
         Caption         =   "Background"
         Begin VB.Menu changes 
            Caption         =   "Change..."
         End
         Begin VB.Menu FontColor 
            Caption         =   "Font Color..."
         End
         Begin VB.Menu Position 
            Caption         =   "Position"
            Begin VB.Menu Strech 
               Caption         =   "Top Left"
            End
            Begin VB.Menu Center 
               Caption         =   "Center"
            End
            Begin VB.Menu Tile 
               Caption         =   "Tile"
               Checked         =   -1  'True
            End
         End
      End
   End
   Begin VB.Menu Lv 
      Caption         =   "ListView "
      Begin VB.Menu xdelete 
         Caption         =   "&Delete Selected item    ÇáÛÇÁ ÇáØÑÝ ÇáãÙáá"
      End
      Begin VB.Menu xPrint 
         Caption         =   "&Print Entries      ØÈÇÚÉ ÇáØÑÝ ÇáãÙáá"
      End
      Begin VB.Menu FV 
         Caption         =   "&Full View"
      End
   End
   Begin VB.Menu ModifyTV 
      Caption         =   "ModifyTV"
      Begin VB.Menu colaps 
         Caption         =   "Expand"
      End
      Begin VB.Menu DeleteLevel 
         Caption         =   "Delete "
      End
      Begin VB.Menu Preview 
         Caption         =   "&Print Preview..."
      End
      Begin VB.Menu xFind 
         Caption         =   "Find..."
      End
   End
   Begin VB.Menu Graph 
      Caption         =   "&Graph"
      Begin VB.Menu ChartType 
         Caption         =   "Chart Type"
         Begin VB.Menu cBar 
            Caption         =   "Bar"
         End
         Begin VB.Menu cPie 
            Caption         =   "Pie"
         End
         Begin VB.Menu cLine 
            Caption         =   "Line"
            Checked         =   -1  'True
         End
         Begin VB.Menu cArea 
            Caption         =   "Area"
         End
      End
      Begin VB.Menu PrintChart 
         Caption         =   "Print"
      End
   End
End
Attribute VB_Name = "SCMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sqltable As Boolean
Dim mConnek As Boolean
Dim xTimes As Integer

Dim MItem As ListItem
Dim xcol As ColumnHeader


Dim cTOtalDebit As Currency
Dim xtable As String
Dim CON1 As ADODB.Connection
Dim rst As ADODB.Recordset

 

Private Sub cArea_Click()
Me.cArea.Checked = False
Me.cBar.Checked = False
Me.cLine.Checked = False
Me.cPie.Checked = False
Me.cArea.Checked = True
GraphExpense.MSChart1.chartType = VtChChartType2dArea
End Sub

Private Sub cBar_Click()
Me.cArea.Checked = False
Me.cBar.Checked = True
Me.cLine.Checked = False
Me.cPie.Checked = False
Me.cArea.Checked = False
GraphExpense.MSChart1.chartType = VtChChartType2dBar
End Sub

Private Sub Center_Click()
'Me.Center.Checked = True
'Me.Strech.Checked = False
'Me.Tile.Checked = False
''Mainform.ListView1.PictureAlignment = lvwCenter

End Sub

Private Sub changes_Click()
On Error GoTo CancelSelect
With Me.CommonDialog1
    Me.CommonDialog1.CancelError = True
    .DialogTitle = "Open Picture for Ledger Background ×ÍËì ºåÄÚÞË ÈÎÞ /ËíáËÞ áÂÔÄäáÞÎÚìí"
    .Filter = "Bitmap (*.Bmp)|*.Bmp|Jpeg(*.Jpg)|*.Jpg|"

    .ShowOpen
    X = .FileName
    
    

CancelSelect:
cErr = Err.Number
If cErr = 32755 Then
   Exit Sub
  Else
    On Error GoTo nopicture:
   'Mainform.ListView1.Picture = LoadPicture(x)
   
  End If
End With
 
nopicture:
itserr = Err.Number
cc = Err.Description
If itserr = 53 Then
 X = ""
 'Mainform.ListView1.Picture = LoadPicture(x)
End If

'Save bakground file
 SaveSetting APp.Title, "BackGround", "Background", X
End Sub

Private Sub cLine_Click()
Me.cArea.Checked = False
Me.cBar.Checked = False
Me.cLine.Checked = True
Me.cPie.Checked = False
Me.cArea.Checked = False
GraphExpense.MSChart1.chartType = VtChChartType2dLine
End Sub

Private Sub colaps_Click()
Dim xindex As Long

xindex = AccTreeView.TreeView1.SelectedItem.Index
If AccTreeView.TreeView1.SelectedItem.Text = "Accounts" Then
 If Me.colaps.caption = "Expand" Then
    i = 0
    For i = 1 To AccTreeView.TreeView1.Nodes.Count
        AccTreeView.TreeView1.Nodes.Item(i).Expanded = True
    Next
    Me.colaps.caption = "Collapse"
  Else
  i = xindex
  For i = 1 To AccTreeView.TreeView1.Nodes.Count
        AccTreeView.TreeView1.Nodes.Item(i).Expanded = False
    Next
    Me.colaps.caption = "Expand"
  End If
 Else
 totalItem = AccTreeView.TreeView1.SelectedItem.Children
 If AccTreeView.TreeView1.SelectedItem.Expanded = False Then
  On Error Resume Next
  totalItem = totalItem + AccTreeView.TreeView1.SelectedItem.Next.Children
  On Error GoTo 0
  i = 0
    For i = xindex To xindex + totalItem ' AccTreeView.TreeView1.Nodes.Count
        AccTreeView.TreeView1.Nodes.Item(i).Expanded = True
        X = AccTreeView.TreeView1.Nodes.Item(i).Children
        
    Next
    Me.colaps.caption = "Collapse"
  Else
   i = 0
    For i = xindex To xindex + totalItem ' AccTreeView.TreeView1.Nodes.Count
        AccTreeView.TreeView1.Nodes.Item(i).Expanded = False
    Next
    Me.colaps.caption = "Expand"
 End If
End If
End Sub

Private Sub cPie_Click()
Me.cArea.Checked = False
Me.cBar.Checked = False
Me.cLine.Checked = False
Me.cPie.Checked = True
Me.cArea.Checked = False
GraphExpense.MSChart1.chartType = VtChChartType2dPie
End Sub

Private Sub DeleteLevel_Click()
xItem = AccTreeView.TreeView1.SelectedItem.Text
xindex = AccTreeView.TreeView1.SelectedItem.Index
i = 0
'let's get the account number selected
For i = 1 To Len(xItem)
   If Mid(xItem, i, 1) = " " Then
     Exit For
    Else
    Ac = Ac & Mid(xItem, i, 1)
   End If
Next i
Ac = Trim(Ac)
xmsg = MsgBox("Are you sure you want to delete " & xItem & "?", vbQuestion + vbYesNo + vbDefaultButton2, "Please confirm")
If xmsg = vbYes Then
    Dim rstDel As New ADODB.Recordset
    con = "DSN=Ledger;UID=sa;pwd=;"
    If Len(Ac) = 15 Then
        rstDel.Open "Delete from accounts where accountnumber=" & "'" & Ac & "'", con
     ElseIf Len(Ac) = 11 Then
         rstDel.Open "Delete from level3 where code=" & "'" & Ac & "'", con
    ElseIf Len(Ac) = 7 Then
         rstDel.Open "Delete from level2 where code=" & "'" & Ac & "'", con
    ElseIf Len(Ac) = 3 Then
         rstDel.Open "Delete from level1 where code=" & "'" & Ac & "'", con
    End If
    AccTreeView.TreeView1.Nodes.Remove xindex
  
 End If
End Sub

Private Sub details_Click()
  SCMenu.Li.Checked = False
  SCMenu.details.Checked = False
  SCMenu.SI.Checked = False
  SCMenu.details.Checked = True
  SCMenu.List.Checked = False
  'ListTAbles.ListView1.View = lvwReport
  'Mainform.ListView1.View = lvwReport
End Sub

Private Sub find_Click()
FindItemLV.Show 1
End Sub

Private Sub FontColor_Click()
On Error GoTo CancelSelect
With Me.CommonDialog1
    .DialogTitle = "Select Color for Listview"
    .ShowColor
    X = .Color
End With
CancelSelect:
cErr = Err.Number
If cErr = 32755 Then
   Exit Sub
  Else
   'Mainform.ListView1.ForeColor = x
   
End If
'Mainform.ListView1.ForeColor = x
SaveSetting APp.Title, "FontColor", "fontcolor", X

End Sub

Private Sub Fv_Click()
If Me.FV.caption = "&Normal View" Then
  SCMenu.FV.caption = "&Full View"
   GenJournalEntry.Frame5.Top = 4920
   GenJournalEntry.ListView1.Top = 5160
   GenJournalEntry.ListView1.Height = 1935
   GenJournalEntry.Frame5.Height = 2295
  Else
   SCMenu.FV.caption = "&Normal View"
   GenJournalEntry.Frame5.Top = 350
   GenJournalEntry.ListView1.Top = 630
   GenJournalEntry.Frame5.Height = 7600
   GenJournalEntry.ListView1.Height = 7200
End If


End Sub

Private Sub Li_Click()
  SCMenu.Li.Checked = True
  SCMenu.details.Checked = False
  SCMenu.SI.Checked = False
  SCMenu.details.Checked = False
  SCMenu.List.Checked = False
  'ListTAbles.ListView1.View = lvwIcon
  'Mainform.ListView1.View = lvwIcon
  
End Sub

Private Sub List_Click()
  SCMenu.Li.Checked = False
  SCMenu.details.Checked = False
  SCMenu.SI.Checked = False
  SCMenu.details.Checked = False
  SCMenu.List.Checked = True
'  ListTAbles.ListView1.View = lvwList
  'Mainform.ListView1.View = lvwList
End Sub

Private Sub refresh_Click()
VoucherPayment.Timer1.Enabled = True
End Sub

Private Sub reneamelevel_Click()
End Sub

Private Sub Open_Click()
'Check the temptable if it is empty. if not abort.
'If Mainform.ListView1.ListItems.Count <> 0 Then
 Dim rsttran As ADODB.Recordset
 Set rsttran = New ADODB.Recordset
 con = "DSN=Ledger;UID=sa;pwd=;"
 rsttran.Open "TempTable ", con, adOpenDynamic, adLockOptimistic, adCmdTable
 If rsttran.EOF = False Then
    xmsg = MsgBox("There was unsaved transactions that you must save first before Deleting or Modifying any saved transactions.", vbInformation + vbOKOnly, "Message")
     Exit Sub
 End If
 rsttran.Close
 'Mainform.Text2.Text = "Edit"
 Createtrans.Show 1
'End If
End Sub

Private Sub Post_Click()
Call xREfresh_Click
  'If Mainform.ListView1.ListItems.Count <> 0 Then
    TRansPosting.Show 1
    'Else
      xmsg = MsgBox("There's No Transactions to post", vbOKOnly + vbInformation, "Message")
    Exit Sub
  'End If
End Sub

Private Sub Preview_Click()
AccountCharts.Show 1
End Sub

Private Sub PrintChart_Click()
GraphExpense.PrintForm
End Sub

Private Sub SI_Click()
  SCMenu.Li.Checked = False
  SCMenu.details.Checked = False
  SCMenu.SI.Checked = True
  SCMenu.details.Checked = False
  SCMenu.List.Checked = False
  'ListTAbles.ListView1.View = lvwSmallIcon
  'Mainform.ListView1.View = lvwSmallIcon
End Sub

Sub ConvertNO(xTimes, xTimesCaption)
 If Trim(Str(Right(xTimes, 1))) = "1" Then
    xTimesCaption = "st"
  ElseIf Trim(Str(Right(xTimes, 1))) = "2" Then
    xTimesCaption = "nd"
 ElseIf Trim(Str(Right(xTimes, 1))) = "3" Then
    xTimesCaption = "rd"
 ElseIf Trim(Str(Right(xTimes, 1))) >= "4" Then
    xTimesCaption = "th"
 End If
End Sub

Private Sub Timer1_Timer()
End Sub

Private Sub Strech_Click()
Me.Center.Checked = False
Me.Strech.Checked = True
Me.Tile.Checked = False
'Mainform.ListView1.PictureAlignment = lvwTopLeft

End Sub

Private Sub Tile_Click()
Me.Center.Checked = False
Me.Strech.Checked = False
Me.Tile.Checked = True
'Mainform.ListView1.PictureAlignment = lvwTile

End Sub

Private Sub xdelete_Click()
Dim xtemp As New ADODB.Recordset
xmsg = MsgBox("Delete TN " & GenJournalEntry.ListView1.SelectedItem & "?", vbQuestion + vbYesNo, "Please confirm")
        If xmsg = vbYes Then
            TN = Createtrans.ListView1.SelectedItem.SubItems(1)
            xtemp.Open "delete temptable  where ticket=" & "'" & TN & "'", constring, adOpenKeyset, adLockOptimistic, adCmdText
            
            cindex = Val(Createtrans.ListView1.SelectedItem.Index)
            If Val(Createtrans.ListView1.ListItems.Item(1).SubItems(1)) <> 1 Then
              Newtn = Val(Createtrans.ListView1.ListItems.Item(1).SubItems(1))
              xItem = Val(cindex)
             Else
              xItem = 1
            End If
 
            
            Createtrans.ListView1.ListItems.Remove cindex
            
            Createtrans.Combo6 = ""
            Createtrans.Combo12 = ""
            Createtrans.MaskEdBox1.Text = ""
            Createtrans.Combo8 = ""
            Createtrans.Combo9 = ""
            Createtrans.Combo7 = ""
            Createtrans.Combo13 = ""
            Createtrans.MaskEdBox2.Text = ""
            Createtrans.Combo12 = ""
            Createtrans.Combo10 = ""
            'get the total no. of rec on transaction table
            'and in temptable
            'Dim TN As Long
            Dim TN1 As Long
            xtemp.Open "SElect count(ticket) as [tn] from transactions where transdate=" & "'" & Date & "'" & " and Deletemark=" & "'" & 0 & "'", constring, adOpenDynamic, adLockOptimistic, adCmdText
            If Createtrans.Text4 = "Edit" Then 'dont change the TN if it is editing
                'TN1 = xtemp!TN '+ 1
              Else
                TN1 = xtemp!TN '+ 1
            End If
            xtemp.Close
           
           
            i = 0
            For i = xItem To Createtrans.ListView1.ListItems.Count
               TN = TN1 + i ' createtrans.ListView1.ListItems.Item(i).SubItems(1)
               Dim GetNewTn As Boolean
               If GetNewTn = False Then
                  'If Val(Me.ListView1.ListItems.Item(1).SubItems(1)) <> 1 Then
                    'newtn = Val(Me.ListView1.ListItems.Item(1).SubItems(1))
                    GetNewTn = True
                    If xItem > 1 Then
                     TN = Newtn + 1
                     Newtn = Newtn + 1
                     Else
                     TN = IIf(IsEmpty(Newtn) = True, 1, Newtn)
                    End If
                 Else
                 Newtn = IIf(IsEmpty(Newtn) = True, TN, Newtn + 1)
                 TN = Newtn
               End If
               
               Ac = Createtrans.ListView1.ListItems.Item(i).SubItems(2)
               an = Createtrans.ListView1.ListItems.Item(i).SubItems(3)
               dr = Createtrans.ListView1.ListItems.Item(i).SubItems(5)
               cR = Createtrans.ListView1.ListItems.Item(i).SubItems(6)
               Jn = Createtrans.ListView1.ListItems.Item(i).SubItems(7)
               fcy = Createtrans.ListView1.ListItems.Item(i).SubItems(8)
               descr = Createtrans.ListView1.ListItems.Item(i).SubItems(9)
               xTN = Createtrans.ListView1.ListItems.Item(i).SubItems(1)
               NextTn = TN
               cindex = Val(Createtrans.ListView1.ListItems.Item(i).Index)
               Createtrans.ListView1.ListItems.Remove cindex
               
               xtemp.Open "delete temptable  where ticket=" & "'" & xTN & "'", constring, adOpenKeyset, adLockOptimistic, adCmdText
               
               'add again but with the new TN
               Set MItem = Createtrans.ListView1.ListItems.Add(, , NextTn, , 1)
               MItem.SubItems(1) = NextTn
               MItem.SubItems(2) = Ac
               MItem.SubItems(3) = an
               MItem.SubItems(4) = Date
               MItem.SubItems(5) = dr
               MItem.SubItems(6) = cR
               MItem.SubItems(7) = Jn
               MItem.SubItems(8) = fcy
               MItem.SubItems(9) = descr
               Createtrans.ListView1.SortKey = 1
               Createtrans.ListView1.Sorted = True
               On Error Resume Next
               With rstTemp
                  .AddNew
                  !ticket = NextTn
                  !accountnumber = Ac
                  !accountname = an
                  !TRansDate = Date
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
               On Error GoTo 0
           Next
          End If
          Exit Sub
          End Sub

Private Sub xFind_Click()
Find1.Show 1
End Sub

Private Sub xPrint_Click()
'tsik the setting for right to left printing
           Set RstBA = New ADODB.Recordset
           'conString = "DSN=Ledger;Uid=sa;pwd=;"
           RstBA.Open "Setup", constring, adOpenKeyset, adLockPessimistic, adCmdTable
           If RstBA!RightToLeftPrn = "Yes" Then
               Printer.RightToLeft = True
              Else
                Printer.RightToLeft = False
           End If
           RstBA.Close
           
           Printer.Orientation = 1
           Printer.FontName = "Arabic Transparent"
           Printer.Print
           Printer.Print
           Printer.Print
           Printer.Print
           Printer.FontSize = 12
           Printer.Print ; Tab(50); Createtrans.Combo11
           Printer.Print
           Printer.Print ; Tab(8); Date
           Printer.Print
           Printer.Print
           Printer.FontSize = 10
           
           i = 0
           For i = 1 To Createtrans.ListView1.ListItems.Count
                  TN = Createtrans.ListView1.ListItems.Item(i).SubItems(1)
                  Ac = Createtrans.ListView1.ListItems.Item(i).SubItems(2)
                  an = Createtrans.ListView1.ListItems.Item(i).SubItems(3)
                  dr = FormatNumber(Createtrans.ListView1.ListItems.Item(i).SubItems(5), 2, vbTrue, vbTrue, vbTrue)
                  cR = FormatNumber(Createtrans.ListView1.ListItems.Item(i).SubItems(6), 2, vbTrue, vbTrue, vbTrue)
                  Jn = Createtrans.ListView1.ListItems.Item(i).SubItems(7)
                  fcy = Createtrans.ListView1.ListItems.Item(i).SubItems(8)
                  descr = Createtrans.ListView1.ListItems.Item(i).SubItems(9)
                  drCol = 90 - Len(cR)
                  CrCol = 110 - Len(dr)
                  Printer.Print ; Tab(10); an; Tab(50); Ac _
                              ; Tab(drCol - Len(dr)); IIf(cR <> 0, cR, "") _
                              ; Tab(CrCol - Len(cR)); IIf(dr <> 0, dr, "")
                  If dr = 0 Then
                    Printer.Print ; Tab(10); "Desc : " & descr
                  End If
           Next i
           Printer.EndDoc
End Sub

Private Sub xREfresh_Click()
Dim MItem As ListItem
Dim xcol As ColumnHeader
'Mainform.ListView1.ColumnHeaders.Clear
'Mainform.ListView1.ListItems.Clear
'Set xcol = Mainform.ListView1.ColumnHeaders.Add(, "h", "!", 300)
'Set xcol = Mainform.ListView1.ColumnHeaders.Add(, "A", "TN ÈíÈíÓ", 600)
'Set xcol = Mainform.ListView1.ColumnHeaders.Add(, "b", "Account No.ÑÞã ÇáÍÓÇÈ ", 2100)
'Set xcol = Mainform.ListView1.ColumnHeaders.Add(, "c", "AccountName ÑÞã ÇáÍÓÇÈ ", 2000)
'Set xcol = Mainform.ListView1.ColumnHeaders.Add(, "d", "Trans Date ËÕÖÞ", 1550)
'Set xcol = Mainform.ListView1.ColumnHeaders.Add(, "e", "Debit ËÕÖÞ", 1200)
'Set xcol = Mainform.ListView1.ColumnHeaders.Add(, "f", "Credit ËÕÖÞ", 1200)
'Set xcol = Mainform.ListView1.ColumnHeaders.Add(, "g", "SerialNo ÊÝÇÕíá", 1100)
'Set xcol = Mainform.ListView1.ColumnHeaders.Add(, "i", "Details ÊÝÇÕíá", 2000)

Dim rst As ADODB.Recordset
Set rst = New ADODB.Recordset
Set CON1 = New ADODB.Connection
'conString = "Provider=MSDASQL;DSN=Ledger;UID=sa;pwd=;"
xtable = "Unposted"
Status = "Unposted"
xKey = "SELECT * From Transactions  Where  Status =" & "'" & Status & "'" & "order by ticket"
rst.Open xKey, constring, adOpenDynamic, adLockOptimistic, adCmdText
'Mainform.ListView1.ColumnHeaders(2).Alignment = lvwColumnRight
'Mainform.ListView1.ColumnHeaders(4).Alignment = lvwColumnRight
'Mainform.ListView1.ColumnHeaders(5).Alignment = lvwColumnCenter
'Mainform.ListView1.ColumnHeaders(6).Alignment = lvwColumnRight
'Mainform.ListView1.ColumnHeaders(7).Alignment = lvwColumnRight
'Mainform.ListView1.ColumnHeaders(9).Alignment = lvwColumnRight
Dim Xdate As Date
Xdate = Format(Date, "mm/dd/yyyy")
Do Until rst.EOF
  If rst!Status = "Unposted" Then
   rstTRansDate = Format(rst!TRansDate, "mm/dd/yyyy")
   cAmount = IIf(rst!DebitAmount > 0, Format(rst!DebitAmount, "###,###,###.#0"), Format(rst!creditamount, "###,###,###.#0"))
   'Set mitem = Mainform.ListView1.ListItems.Add(, , rst!Ticket, IIf(rstTRansDate = Xdate, 1, 1), IIf(rstTRansDate = Xdate, 1, IIf(rstTRansDate = Xdate - 1, 2, IIf(rstTRansDate = Xdate - 2, 3, IIf(rstTRansDate = Xdate - 3, 4, IIf(rst!TransDate = Date - 4, 5, 5))))))
    MItem.SubItems(1) = IIf(Len(rst!ticket) = 1, "  " _
                        & (rst!ticket), IIf(Len(rst!ticket) = 2, _
                       " " & rst!ticket, rst!ticket))
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
rst.Close


End Sub

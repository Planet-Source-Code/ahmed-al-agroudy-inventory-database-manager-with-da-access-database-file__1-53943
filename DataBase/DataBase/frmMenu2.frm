VERSION 5.00
Begin VB.Form frmMyMenu2 
   Caption         =   "Form2"
   ClientHeight    =   3135
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   7440
   LinkTopic       =   "Form2"
   ScaleHeight     =   3135
   ScaleWidth      =   7440
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu menu1 
      Caption         =   "Menu1"
      Begin VB.Menu Add 
         Caption         =   "Add"
      End
      Begin VB.Menu edit 
         Caption         =   "Edit"
      End
      Begin VB.Menu del 
         Caption         =   "Delete"
      End
      Begin VB.Menu clear 
         Caption         =   "Clear"
      End
   End
   Begin VB.Menu Status 
      Caption         =   "status"
      Begin VB.Menu sAdd 
         Caption         =   "Add"
      End
      Begin VB.Menu sEdit 
         Caption         =   "Edit"
      End
      Begin VB.Menu sDel 
         Caption         =   "Delete"
      End
      Begin VB.Menu sclear 
         Caption         =   "clear"
      End
   End
   Begin VB.Menu Shift 
      Caption         =   "Shift"
      Begin VB.Menu ShAdd 
         Caption         =   "Add"
      End
      Begin VB.Menu shedit 
         Caption         =   "Edit"
      End
      Begin VB.Menu Shdel 
         Caption         =   "Delete"
      End
      Begin VB.Menu shclear 
         Caption         =   "Clear"
      End
   End
End
Attribute VB_Name = "frmMyMenu2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstpos As ADODB.Recordset
Dim RstStat As ADODB.Recordset
Dim rstShift As ADODB.Recordset
Dim xvar As String

Private Sub Add_Click()
frmPosition.Text1.SetFocus
End Sub

Private Sub clear_Click()
If frmPosition.cmdedit.caption <> "&Edit" Then
frmPosition.cmdedit.caption = "&Edit"
frmPosition.Command4.Enabled = True
frmPosition.cmdsave.Enabled = True
frmPosition.Command5.caption = "E&xit"
Me.Edit.Enabled = True
frmPosition.Text1.Text = ""
frmPosition.Text2.Text = ""
frmPosition.Text3.Text = ""

End If
End Sub

Private Sub del_Click()
Set rstpos = New ADODB.Recordset

Dim CON1 As ADODB.Connection
Set CON1 = New ADODB.Connection


conStr = "Provider=MSDASQL;DSN=payrollcairo;UID=; PWD=;"

CON1.Open conStr

    rstpos.Open "Select * from position_old", CON1, adOpenDynamic, adLockOptimistic

If frmPosition.ListView1.ListItems.Count = 0 Then
    Me.del.Enabled = False
    Exit Sub
  End If
varindex = frmPosition.ListView1.SelectedItem.Index
varitem = frmPosition.ListView1.SelectedItem.Text
'varsubit0 = frmPosition.
varsubitem = frmPosition.ListView1.SelectedItem.SubItems(1)
 xmsg = MsgBox("Are you Sure Deleting Position?    " & varsubitem & "", vbQuestion + vbYesNo, " Delete...")

  If xmsg = vbYes Then
     'frmPosition.ListView1.ListItems.Remove varindex
     xvar = Trim(varitem)
If rstpos.EOF = False Then
rstpos.MoveFirst
End If

      While rstpos.EOF = False
        If Trim(rstpos!Code) = Trim(xvar) Then
            rstpos.Delete
            rstpos.Update
            MsgBox "Records Deleted Successfully", vbInformation, "Confirmation"
        End If
     rstpos.MoveNext
        Wend
     
End If
End Sub

Private Sub edit_Click()

Set rstpos = New ADODB.Recordset

Dim CON1 As ADODB.Connection
Set CON1 = New ADODB.Connection


conStr = "Provider=MSDASQL;DSN=payrollcairo;UID=; PWD=;"

CON1.Open conStr

    rstpos.Open "Select * from position_old", CON1, adOpenDynamic, adLockOptimistic

If frmPosition.ListView1.ListItems.Count = 0 Then
    Me.Edit.Enabled = False
    Exit Sub
  End If
varindex = frmPosition.ListView1.SelectedItem.Index
varitem = frmPosition.ListView1.SelectedItem.Text
varsubitem = frmPosition.ListView1.SelectedItem.SubItems(1)
 'xmsg = MsgBox("Are you Sure Deleting Position?    " & varsubitem & "", vbQuestion + vbYesNo, " Delete...")

  If frmPosition.cmdedit.caption = Trim("&Edit") Then
     'frmPosition.ListView1.ListItems.Remove varindex
     xvar = Trim(varitem)
     If rstpos.EOF = False Then
rstpos.MoveFirst
End If

      While rstpos.EOF = False
        If Trim(rstpos!Code) = Trim(xvar) Then
        frmPosition.Text1.Text = IIf(IsNull(Trim(rstpos!Code)), "", Trim(rstpos!Code))
        frmPosition.Text2.Text = IIf(IsNull(Trim(rstpos!Name)), "", Trim(rstpos!Name))
        frmPosition.Text3.Text = IIf(IsNull(Trim(rstpos!arab)), "", Trim(rstpos!arab))
        frmPosition.cmdsave.Enabled = False
        frmPosition.Command4.Enabled = False
        frmPosition.cmdedit.caption = "Update"
        frmPosition.Command5.caption = "&Cancel"
        Me.Clear.Enabled = True

        End If
     rstpos.MoveNext
        Wend
     
End If
Me.Edit.Enabled = False

End Sub

Private Sub sAdd_Click()
frmPayCat.Text1.SetFocus
End Sub

Private Sub sclear_Click()

If frmPayCat.cmdedit.caption <> "&Edit" Then
frmPosition.cmdedit.caption = "&Edit"
frmPosition.Command4.Enabled = True
frmPosition.cmdsave.Enabled = True
frmPosition.Command5.caption = "E&xit"
Me.Edit.Enabled = True
frmPosition.Text1.Text = ""
frmPosition.Text2.Text = ""
frmPosition.Text3.Text = ""
End If
End Sub

Private Sub sDel_Click()
Set RstStat = New ADODB.Recordset
Dim CON1 As ADODB.Connection
Set CON1 = New ADODB.Connection

conStr = "Provider=MSDASQL;DSN=Finance;UID=sa; PWD=;"

CON1.Open conStr

    RstStat.Open "Select * from newPAymentfor", CON1, adOpenDynamic, adLockOptimistic

If frmPayCat.ListView1.ListItems.Count = 0 Then
    Me.del.Enabled = False
    Exit Sub
  End If
varindex = frmPayCat.ListView1.SelectedItem.Index
varitem = frmPayCat.ListView1.SelectedItem.Text
'varsubit0 = frmPayCat.
varsubitem = frmPayCat.ListView1.SelectedItem.SubItems(1)
 xmsg = MsgBox("Are you Sure Deleting Status?    " & varsubitem & "", vbQuestion + vbYesNo, " Delete...")

  If xmsg = vbYes Then
     'frmPayCat.ListView1.ListItems.Remove varindex
     xvar = Trim(varitem)

If RstStat.EOF = False Then
RstStat.MoveFirst
End If
     
      While RstStat.EOF = False
        If Trim(RstStat!Code) = Trim(xvar) Then
            RstStat.Delete
            RstStat.Update
            MsgBox "Records Deleted Successfully", vbInformation, "Confirmation"
        End If
     RstStat.MoveNext
        Wend
     
End If
End Sub

Private Sub sEdit_Click()

Set RstStat = New ADODB.Recordset

Dim CON1 As ADODB.Connection
Set CON1 = New ADODB.Connection
conStr = "Provider=MSDASQL;DSN=Finance;UID=sa; PWD=;"

CON1.Open conStr

    RstStat.Open "Select * from newPAymentfor", CON1, adOpenDynamic, adLockOptimistic

If frmPayCat.ListView1.ListItems.Count = 0 Then
    Me.sEdit.Enabled = False
    Exit Sub
  End If
varindex = frmPayCat.ListView1.SelectedItem.Index
varitem = frmPayCat.ListView1.SelectedItem.Text
varsubitem = frmPayCat.ListView1.SelectedItem.SubItems(1)
 'xmsg = MsgBox("Are you Sure Deleting Position?    " & varsubitem & "", vbQuestion + vbYesNo, " Delete...")

  If frmPayCat.cmdedit.caption = Trim("&Edit") Then
     'frmPayCat.ListView1.ListItems.Remove varindex
     xvar = Trim(varitem)
     
   If RstStat.EOF = False Then
RstStat.MoveFirst
End If
  
     
      While RstStat.EOF = False
        If Trim(RstStat!Code) = Trim(xvar) Then
        frmPayCat.Text1.Text = IIf(IsNull(Trim(RstStat!Code)), "", Trim(RstStat!Code))
        frmPayCat.Text2.Text = IIf(IsNull(Trim(RstStat!Name)), "", Trim(RstStat!Name))
        frmPayCat.Text3.Text = IIf(IsNull(Trim(RstStat!arab)), "", Trim(RstStat!arab))
        frmPayCat.cmdsave.Enabled = False
        frmPayCat.Command4.Enabled = False
        frmPayCat.cmdedit.caption = "Update"
        frmPayCat.Command5.caption = "&Cancel"
        Me.Clear.Enabled = True

        End If
     RstStat.MoveNext
        Wend
     
End If
Me.sEdit.Enabled = False
End Sub

Private Sub ShAdd_Click()
frmShift.Text1.SetFocus
End Sub

Private Sub Shdel_Click()
Set rstShift = New ADODB.Recordset
Dim CON1 As ADODB.Connection
Set CON1 = New ADODB.Connection

conStr = "Provider=MSDASQL;DSN=payrollcairo;UID=; PWD=;"

CON1.Open conStri

    rstShift.Open "Select * from Employeeshiftcode", CON1, adOpenDynamic, adLockOptimistic

If frmShift.ListView1.ListItems.Count = 0 Then
    Me.Shdel.Enabled = False
    Exit Sub
  End If
varindex = frmShift.ListView1.SelectedItem.Index
varitem = frmShift.ListView1.SelectedItem.Text
'varsubit0 = frmShift.
varsubitem = frmShift.ListView1.SelectedItem.SubItems(1)
 xmsg = MsgBox("You are About to Delete Shift, Delete?    " & varsubitem & "", vbQuestion + vbYesNo, " Delete...")

  If xmsg = vbYes Then
     'frmShift.ListView1.ListItems.Remove varindex
     xvar = Trim(varitem)
     
If rstShift.EOF = False Then
rstShift.MoveFirst
End If

      While rstShift.EOF = False
        If Trim(rstShift!shfcode) = Trim(xvar) Then
            rstShift.Delete
            rstShift.Update
            MsgBox "Records Deleted Successfully", vbInformation, "Confirmation"
        End If
     rstShift.MoveNext
        Wend
     
End If

End Sub

Private Sub shedit_Click()

Set rstShift = New ADODB.Recordset
Dim CON1 As ADODB.Connection
Set CON1 = New ADODB.Connection
conStr = "Provider=MSDASQL;DSN=payrollcairo;UID=; PWD=;"

CON1.Open conStr

    rstShift.Open "Select * from employeeshiftcode", CON1, adOpenDynamic, adLockOptimistic

If frmShift.ListView1.ListItems.Count = 0 Then
    Me.sEdit.Enabled = False
    Exit Sub
  End If
varindex = frmShift.ListView1.SelectedItem.Index
varitem = frmShift.ListView1.SelectedItem.Text
varsubitem = frmShift.ListView1.SelectedItem.SubItems(1)
 'xmsg = MsgBox("Are you Sure Deleting Position?    " & varsubitem & "", vbQuestion + vbYesNo, " Delete...")

  If frmShift.cmdedit.caption = Trim("&Edit") Then
     'frmShift.ListView1.ListItems.Remove varindex
     xvar = Trim(varitem)
 If rstShift.EOF = False Then
rstShift.MoveFirst
End If
    
     
      While rstShift.EOF = False
        If Trim(rstShift!shfcode) = Trim(xvar) Then
        frmShift.Text1.Text = IIf(IsNull(Trim(rstShift!shfcode)), "", Trim(rstShift!shfcode))
        frmShift.Text2.Text = IIf(IsNull(Trim(rstShift!shfnameeng)), "", Trim(rstShift!shfnameeng))
        frmShift.Text3.Text = IIf(IsNull(Trim(rstShift!shfnameara)), "", Trim(rstShift!shfnameara))
        frmShift.cmdsave.Enabled = False
        frmShift.Command4.Enabled = False
        frmShift.cmdedit.caption = "Update"
        frmShift.Command5.caption = "&Cancel"
        Me.Clear.Enabled = True

        End If
     rstShift.MoveNext
        Wend
     
End If
Me.sEdit.Enabled = False
End Sub

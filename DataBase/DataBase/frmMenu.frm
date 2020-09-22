VERSION 5.00
Begin VB.Form frmMenu 
   Caption         =   "My Menu"
   ClientHeight    =   1095
   ClientLeft      =   165
   ClientTop       =   1155
   ClientWidth     =   8295
   LinkTopic       =   "Form2"
   ScaleHeight     =   1095
   ScaleWidth      =   8295
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu Terms 
      Caption         =   "Terms"
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
   Begin VB.Menu rate 
      Caption         =   "Rate"
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
   Begin VB.Menu Rec 
      Caption         =   "Reciept"
      Begin VB.Menu rAdd 
         Caption         =   "Add"
      End
      Begin VB.Menu REdit 
         Caption         =   "Edit"
      End
      Begin VB.Menu Rdel 
         Caption         =   "Delete"
      End
      Begin VB.Menu RClear 
         Caption         =   "Clear"
      End
   End
   Begin VB.Menu ListV 
      Caption         =   "ListV"
      Begin VB.Menu PayCancel 
         Caption         =   "Payable Cancellation"
      End
      Begin VB.Menu kjk 
         Caption         =   "-"
      End
      Begin VB.Menu confirm 
         Caption         =   "Confirm"
      End
      Begin VB.Menu yry 
         Caption         =   "-"
      End
      Begin VB.Menu DoJou 
         Caption         =   " Journalise"
      End
      Begin VB.Menu efcvv 
         Caption         =   "-"
      End
      Begin VB.Menu search 
         Caption         =   "Search"
      End
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
      Begin VB.Menu X 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu xPost 
      Caption         =   "Post"
      Begin VB.Menu xPPV 
         Caption         =   "Print Paid Vouchers"
      End
      Begin VB.Menu PaidDueDate 
         Caption         =   "Paid Voucher List (By Due Date)"
      End
   End
   Begin VB.Menu Conf 
      Caption         =   "Conf"
      Begin VB.Menu Unpaidsss 
         Caption         =   "Print Unpaid Vaouchers"
      End
      Begin VB.Menu conPCL 
         Caption         =   "Print Cancelled List"
      End
      Begin VB.Menu deleateList 
         Caption         =   "Delete from the List"
      End
   End
   Begin VB.Menu PurcSetup 
      Caption         =   "PurcSetup"
      Begin VB.Menu PurchAdd 
         Caption         =   "Add"
      End
      Begin VB.Menu PurchEdit 
         Caption         =   "Edit"
      End
      Begin VB.Menu PurchDel 
         Caption         =   "Delete"
      End
      Begin VB.Menu PurchClear 
         Caption         =   "Clear"
      End
   End
   Begin VB.Menu xItem 
      Caption         =   "Item"
      Begin VB.Menu IAdd 
         Caption         =   "Add"
      End
      Begin VB.Menu IEdit 
         Caption         =   "Edit"
      End
      Begin VB.Menu IDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu IClear 
         Caption         =   "Clear"
      End
   End
   Begin VB.Menu Journal 
      Caption         =   "Journal"
      Begin VB.Menu PtoLedger 
         Caption         =   "Post to Ledger"
      End
      Begin VB.Menu JPr 
         Caption         =   "Print Unposted transactions"
      End
      Begin VB.Menu PrintPoteds 
         Caption         =   "Print Posted Transactions"
      End
      Begin VB.Menu dfsfsfr 
         Caption         =   "-"
      End
      Begin VB.Menu JENT 
         Caption         =   "Print Journal Entry(without Details)"
      End
      Begin VB.Menu PrtPayournalGroup 
         Caption         =   "Print Journal Entry"
      End
      Begin VB.Menu BN 
         Caption         =   "-"
      End
      Begin VB.Menu pmj 
         Caption         =   "Cancellation"
      End
      Begin VB.Menu fsdfsdfsdfs 
         Caption         =   "-"
      End
      Begin VB.Menu UPS 
         Caption         =   "Unpaid Voucher Report(By Due Date)"
      End
   End
   Begin VB.Menu Jounal1 
      Caption         =   "Jounal1"
      Begin VB.Menu JouSave 
         Caption         =   "Save"
      End
   End
   Begin VB.Menu List7 
      Caption         =   "List7"
      Begin VB.Menu PVList 
         Caption         =   "Print Voucher List"
      End
   End
   Begin VB.Menu Invo 
      Caption         =   "Invoic"
      Begin VB.Menu InvEdit 
         Caption         =   "Edit"
      End
   End
   Begin VB.Menu PetCash 
      Caption         =   "PettCash"
      Begin VB.Menu PetCEdit 
         Caption         =   "Edit"
      End
      Begin VB.Menu PettCancel 
         Caption         =   "Delete"
      End
   End
   Begin VB.Menu PettCashUnCon 
      Caption         =   "PettCashUnCon"
      Begin VB.Menu PCuconEdit 
         Caption         =   "Edit"
      End
      Begin VB.Menu PCuconCancell 
         Caption         =   "Cancell"
      End
      Begin VB.Menu PCuconPrint 
         Caption         =   "Print Request (Petty Cash)"
      End
      Begin VB.Menu PCuconDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu bkfgf 
         Caption         =   "-"
      End
      Begin VB.Menu Confir 
         Caption         =   "Confirm"
      End
   End
   Begin VB.Menu PettCashCon 
      Caption         =   "PettCashCon"
      Begin VB.Menu PCconEdit 
         Caption         =   "Edit"
      End
      Begin VB.Menu PCconCancel 
         Caption         =   "Cancel"
      End
      Begin VB.Menu PCconPrint 
         Caption         =   "Print"
      End
      Begin VB.Menu PCconDelete 
         Caption         =   "Delete"
      End
   End
   Begin VB.Menu PettyJournal 
      Caption         =   "PettyJournal"
      Begin VB.Menu PettyPost 
         Caption         =   "Post to the Ledger"
      End
      Begin VB.Menu PettyUnPostReport 
         Caption         =   "Print Unposted"
      End
      Begin VB.Menu PettyPostCancel 
         Caption         =   "Cancellation"
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstterm As ADODB.Recordset
Dim RstStat As ADODB.Recordset
Dim rstShift As ADODB.Recordset
Dim rstList3 As ADODB.Recordset
Dim rstReceipt As ADODB.Recordset
Dim xvar As String
Dim conStr As String
Dim CON1 As ADODB.Connection
Dim rstPurch As ADODB.Recordset
Dim RstItem As ADODB.Recordset
Dim XPundai As String
Public finalbalance  As Currency
Public repPettCRAccNo
Public VarForReportLabel
Public VarForReportLabe2
Public VarForReportLabe3
Dim catName As String
Dim Prevcap As String
Dim acctNo As String
 Dim Categry
Public PettyReqNo
Public PettyDatex
'Public DeleteTerm As String

Private Sub ConfirAlone()
Dim TickNo As String

Dim xJou As ADODB.Recordset
Set xJou = New ADODB.Recordset

Dim Pay As ADODB.Recordset
Set Pay = New ADODB.Recordset

Dim KelithiMeen As ADODB.Recordset
Set KelithiMeen = New ADODB.Recordset
Dim jPOno, jPOdate, jInvoiceNo, jInvDate, jStorNo, jStorDate
Dim jPAy3 As New ADODB.Recordset
Dim PAy4 As New ADODB.Recordset


Dim xSaman
xSaman = FrmPayableSetup.ListView1.SelectedItem

xset = "SELECT * From Payablesetup  Where serialno =" & "'" & xSaman & "'" & " and confirmedmark = 0 order by serialno"

Pay.Open xset, constring, adOpenDynamic, adLockOptimistic, adCmdText

SuraMeen = "Select * from Payablesetup where cancelledmark = '0' and deletemark = '0' and confirmedmark = '0'  and Post = 'No' and Paidmark = '0'"
Kelithi = "Select * from Payablesetup where cancelledmark = '0' and deletemark = '0' and confirmedmark = '1'  and Post = 'No' and Paidmark = '0'"
KelithiMeen.Open SuraMeen, constring, adOpenDynamic, adLockOptimistic
PAy4.Open Kelithi, constring, adOpenDynamic, adLockOptimistic



If Pay.EOF = True Then
On Error Resume Next
End If

If Pay.EOF = False Then
Pay.MoveFirst
End If
            Dim GotitConfirm
            While Pay.EOF = False
            If Trim(xSaman) = Trim(Pay!SerialNo) Then
            Pay!ConfirmedMark = 1
            GotitConfirm = "YEs"
            End If
            Pay.MoveNext
            Wend



'REFRESH List 1

FrmPayableSetup.ListView1.ListItems.Clear
If KelithiMeen.EOF = False Then
KelithiMeen.MoveFirst
End If


  While KelithiMeen.EOF = False
     Set MItem = FrmPayableSetup.ListView1.ListItems.Add(, , Format(KelithiMeen!SerialNo))
     MItem.SubItems(1) = Format(KelithiMeen!Xdate, "dd/mm/yyyy")
     MItem.SubItems(2) = Format(KelithiMeen!Requester)
     MItem.SubItems(3) = Format(KelithiMeen!DateDue, "dd/mm/yyyy")
     MItem.SubItems(4) = Format(KelithiMeen!RefNo)
     MItem.SubItems(5) = Format(KelithiMeen!printmark)
     MItem.SubItems(6) = Format(KelithiMeen!journaledmark)
     MItem.SubItems(7) = Format(KelithiMeen!amtreqested, "#############.#0")
       
       Totlist1 = Val(Totlist1) + Val(Trim(KelithiMeen!amtreqested)) 'This is for the Total of the List
     KelithiMeen.MoveNext
     Wend
FrmPayableSetup.txtTotList1.Text = Totlist1

'-----


        'this is to add Listview7
            FrmPayableSetup.ListView7.ListItems.Clear
            PAy4.MoveFirst
            While PAy4.EOF = False
            Dim Yaas
            Yaas = "Yes"
            Set MItem = FrmPayableSetup.ListView7.ListItems.Add(, , Format(PAy4!SerialNo))
            MItem.SubItems(1) = Format(PAy4!Xdate, "dd/mm/yyyy")
            MItem.SubItems(2) = Format(PAy4!Payee)
            MItem.SubItems(3) = Format(PAy4!DateDue)
            MItem.SubItems(4) = Format(PAy4!RefNo)
            MItem.SubItems(5) = Format(PAy4!amtreqested, "#############.#0")
            
            TotList7z = Val(TotList7z) + Val(Trim(PAy4!amtreqested)) 'This is for the Total of the List
            PAy4.MoveNext
            Wend
            FrmPayableSetup.txtTotList7.Text = Trim(TotList7z)

    If GotitConfirm = "YEs" Then
    MsgBox "Datas Confirmed Succusfully", vbInformation, "Confrimation"
    End If

GotitConfirm = ""


End Sub


Private Sub Add_Click()
FrmPayableSetup.txtTMDays.Visible = True
FrmPayableSetup.txtTMDes.Visible = True
FrmPayableSetup.txtTMlevel.Visible = True
FrmPayableSetup.txtTMmode.Visible = True
FrmPayableSetup.txtTmRate.Visible = True

FrmPayableSetup.txtTmRate.SetFocus
End Sub

Private Sub clear_Click()
'If FrmPayableSetup.cmdNew.Caption = "&Update" Then
'MsgBox "You Cannot Clear The List View Now", vbCritical, "Clear"
'FrmPayableSetup.txtTMDays.Text = ""
'FrmPayableSetup.txtTmRate.Text = ""
'FrmPayableSetup.txtTMDes.Text = ""
'FrmPayableSetup.txtTMmode.Text = ""
'FrmPayableSetup.txtTMlevel.Text = ""
'
'Exit Sub
'End If
'
'FrmPayableSetup.txtTMDays.Visible = False
'FrmPayableSetup.txtTMDes.Visible = False
'FrmPayableSetup.txtTMlevel.Visible = False
'FrmPayableSetup.txtTMmode.Visible = False
'FrmPayableSetup.txtTmRate.Visible = False
'Me.edit.Caption = "Edit"
'FrmPayableSetup.txtTmp1.Text = ""
'FrmPayableSetup.txtTmp2.Text = ""
'FrmPayableSetup.ListView2.ListItems.clear
Me.Edit.caption = "Edit" 'This will chang the Caption So we can Edit Again Inside the Lsit


End Sub



Private Sub Confir_Click() 'This si for the Pettycash
If frmPettyCash.LvwPCashUnConf.ListItems.Count = 0 Then
MsgBox " Sorry, There is no items to Confirm", vbInformation, "ListView is Empty"
Exit Sub
End If


Dim VarForSErNom
VarForSErNom = frmPettyCash.LvwPCashUnConf.SelectedItem
Dim rstMarkJrlPCashm As New ADODB.Recordset
rstMarkJrlPCashm.Open "Select * from pettycashheld where serialno = " & "'" & VarForSErNom & "'" & " and journalmark = '1'", constring, adOpenDynamic, adLockOptimistic

'If rstMarkJrlPCashm.EOF = True Then
'AVM = MsgBox("There is no record for this No,First you have to Journalise, Do you want to Journalise it", vbInformation + vbYesNo, "Please Select")
'If AVM = vbYes Then
'PCunPosJournal_Click
'Else
'Exit Sub
'End If
'End If
'



GHU = MsgBox("Mr." & cLogUser & ",Do you want to Confime the Selected item", vbYesNo + vbInformation, "Confirmation")
If GHU = vbNo Then
Exit Sub
End If


Call PCunPosJournal2



Dim VarSelItemLVWuc
'VarSelItemLVWuc = frmPettyCash.LvwPCashUnConf.SelectedItem

Dim rstConfPCLstLVWuc As New ADODB.Recordset

rstConfPCLstLVWuc.Open "Select * from pettycashheld where SerialNo = " & "'" & Trim(VarForSErNom) & "'" & "", constring, adOpenDynamic, adLockOptimistic

If rstConfPCLstLVWuc.EOF = False Then
rstConfPCLstLVWuc.MoveFirst
End If

On Error GoTo 0

While rstConfPCLstLVWuc.EOF = False
rstConfPCLstLVWuc!confMark = 1

rstConfPCLstLVWuc.Update

rstConfPCLstLVWuc.MoveNext
Wend

MsgBox "Record has been Confirmed Successfully", vbInformation, "Confirmation"

'Refresh the ListView LVWUc
'Call ProcLvwPCashUnConf

''Refresh the ListView LVWconf
'Call ProcLvwPCashConfirmed

 'This is to call the Class for the Listview confirmed
Dim xcls2 As New HabitatClass
Dim rs2 As New ADODB.Recordset
''xcls2.ProcLvwPCashConfirmed rs2

Dim rsRef As New ADODB.Recordset
'xcls2.ProcLvwPCashUnCon rsRef
End Sub
Private Sub ProcLvwPCashUnConf() 'this is the Procedere used for REFRESH the Listview UNconf
Dim rsConfLVWPCuc As New ADODB.Recordset


rsConfLVWPCuc.Open "Select Distinct serialno,DateX,AccountName,printmark,motham from pettycashheld where confMark is null and DeleteMark is null", constring, adOpenDynamic, adLockOptimistic

frmPettyCash.LvwPCashUnConf.ListItems.Clear
If rsConfLVWPCuc.EOF = False Then
rsConfLVWPCuc.MoveFirst
End If

While rsConfLVWPCuc.EOF = False

Set MItem = frmPettyCash.LvwPCashUnConf.ListItems.Add(, , Trim(rsConfLVWPCuc!SerialNo))
MItem.SubItems(1) = Format(rsConfLVWPCuc!Datex, "dd/mm/yyyy")
MItem.SubItems(2) = Trim(rsConfLVWPCuc!accountname)
MItem.SubItems(3) = IIf(IsNull(rsConfLVWPCuc!printmark), "", (rsConfLVWPCuc!printmark))
MItem.SubItems(4) = IIf(IsNull(rsConfLVWPCuc!motham), "", (rsConfLVWPCuc!motham))
frmPettyCash.txttotUc.Text = Val(frmPettyCash.txttotUc.Text) + Val(rsConfLVWPCuc!motham)

rsConfLVWPCuc.MoveNext
Wend
End Sub

Private Sub confirm_Click() 'This is for Payable SEtup

If FrmPayableSetup.ListView1.ListItems.Count = 0 Then
MsgBox "No item to Confirm", vbCritical, "List is Empty"
Exit Sub
End If


 If FrmPayableSetup.ListView1.SelectedItem.SubItems(5) = "Yes" Then
 Call ConfirAlone
 Exit Sub
 End If

Dim TickNo As String

Dim rstConf1 As ADODB.Recordset
Set rstConf1 = New ADODB.Recordset

Dim rstConf2 As ADODB.Recordset
Set rstConf2 = New ADODB.Recordset

Dim xJou As ADODB.Recordset
Set xJou = New ADODB.Recordset

Dim Pay As ADODB.Recordset
Set Pay = New ADODB.Recordset

Dim KelithiMeen As ADODB.Recordset
Set KelithiMeen = New ADODB.Recordset
Dim jPOno, jPOdate, jInvoiceNo, jInvDate, jStorNo, jStorDate, Explanations
Dim jPAy3 As New ADODB.Recordset
Dim PAy4 As New ADODB.Recordset

Dim AdditonDeta
If FrmPayableSetup.ListView1.ListItems.Count = 0 Then
MsgBox "no item selected"
Exit Sub
Else
End If


ms = MsgBox("Mr." & cLogUser & ", Do you want to Confirme It", vbYesNo, " Confirmed & Journalyse")

If ms = vbNo Then
Exit Sub
Else

Dim xSaman
xSaman = FrmPayableSetup.ListView1.SelectedItem

xpay = "SELECT * From Xpayment  Where serialno =" & "'" & xSaman & "'" & " and confirmedmark = 'uc' and Postmark = 'No' order by serialno"
xrec = "SELECT * From Xreceipt  Where serialno =" & "'" & xSaman & "'" & " and confirmedmark = 'uc' and Postmark = 'No' order by serialno"

xset = "SELECT * From Payablesetup  Where serialno =" & "'" & xSaman & "'" & " and confirmedmark = 0 order by serialno"
'xset2 = "SELECT * From Payablesetup  Where  confirmedmark = 0 order by serialno"
rstConf1.Open xpay, constring, adOpenDynamic, adLockOptimistic, adCmdText
rstConf2.Open xrec, constring, adOpenDynamic, adLockOptimistic, adCmdText
Pay.Open xset, constring, adOpenDynamic, adLockOptimistic, adCmdText

 SuraMeen = "Select * from Payablesetup where cancelledmark = '0' and deletemark = '0' and confirmedmark = '0'  and Post = 'No' and Paidmark = '0'"
 Kelithi = "Select * from Payablesetup where cancelledmark = '0' and deletemark = '0' and confirmedmark = '1'  and Post = 'No' and Paidmark = '0'"
KelithiMeen.Open SuraMeen, constring, adOpenDynamic, adLockOptimistic
PAy4.Open Kelithi, constring, adOpenDynamic, adLockOptimistic

xJou.Open "select * from Payjournal ", constring, adOpenDynamic, adLockOptimistic
If Pay.EOF = True Then
On Error Resume Next
End If


'This is to check the xPayment and XReceipt whether they did the Payment Analysis
'before confirming it

'laka = MsgBox("Mr." & cLogUser & ", You are Trying to Cancell the Journalised List,Are you sure to Cancell it", vbExclamation + vbYesNo, "Please Confirm it")

If rstConf2.EOF = True Then
 MsgBox "Mr." & cLogUser & ",You Have to Do the Payment Analysis to Confirm", vbInformation, "Message"
Exit Sub
End If



'0-0-0-0-0-0 Here another coding  for Listview 1

If Pay.EOF = False Then
Pay.MoveFirst
End If

  While Pay.EOF = False
  If Trim(xSaman) = Trim(Pay!SerialNo) Then
    Pay!ConfirmedMark = 1
     Set MItem = FrmPayableSetup.ListView7.ListItems.Add(, , Format(Pay!SerialNo))
     MItem.SubItems(1) = Format(Pay!Xdate, "dd/mm/yyyy")
     MItem.SubItems(2) = Format(Pay!payto)
     MItem.SubItems(3) = Format(Pay!DocNo)
     MItem.SubItems(4) = Format(Pay!TotCrAmt)
        On Error Resume Next
     Totlist1x = Val(Totlist1x) + Val(Trim(Pay!TotCrAmt)) 'This is for the Total of the List
  On Error GoTo 0
  
  End If
     Pay.MoveNext
     Wend
     FrmPayableSetup.txtTotList1.Text = Trim(Totlist1x)
'0-0-0-0-0-0-
TickNo = 1

If rstConf1.EOF = True Then
On Error Resume Next
End If

'this is to add To the     to Listview 8

FrmPayableSetup.ListView8.ListItems.Clear

If rstConf1.EOF = False Then
rstConf1.MoveFirst
End If

  While rstConf1.EOF = False
     Set MItem = FrmPayableSetup.ListView8.ListItems.Add(, , Format(rstConf1!SerialNo))
     MItem.SubItems(1) = Trim(TickNo)
     MItem.SubItems(2) = Date
     MItem.SubItems(3) = Format(rstConf1!AccNo)
     MItem.SubItems(4) = Format(rstConf1!AccName)
     MItem.SubItems(5) = Format(rstConf1!amount)
     'MItem.SubItems(5) = Format(rstConf1!amount)
TickNo = TickNo + 1
       TotList8Dbx = Val(TotList8Dbx) + Val(Trim(rstConf1!amount)) 'This is for the Total of the List

   rstConf1.MoveNext
   Wend
   

If rstConf2.EOF = True Then
On Error Resume Next
End If

   
If rstConf2.EOF = False Then
rstConf2.MoveFirst
End If

   Dim VarForRecordedDate As Date
   
'this is for the Receipt
  While rstConf2.EOF = False
     Set MItem = FrmPayableSetup.ListView8.ListItems.Add(, , Format(rstConf2!SerialNo))
     MItem.SubItems(1) = Trim(TickNo)
     MItem.SubItems(2) = Date
     MItem.SubItems(3) = Format(rstConf2!AccNo)
     MItem.SubItems(4) = Format(rstConf2!AccName)
     MItem.SubItems(6) = Format(rstConf2!amount)
   TickNo = TickNo + 1
   
      TotList8crx = Val(TotList8crx) + Val(Trim(rstConf2!amount)) 'This is for the Total of the List
 VarForRecordedDate = Format(rstConf2!recordedDate, "dd/mm/yyyy")
   rstConf2.MoveNext
   Wend


FrmPayableSetup.txtJdb.Text = Trim(TotList8Dbx)
FrmPayableSetup.txtJCr.Text = Trim(TotList8crx)


'Add new from Listview to the Table Payjournal

Dim JOurnalNo As New ADODB.Recordset
JOurnalNo.Open "Setup", constring, adOpenKeyset, adLockPessimistic, adCmdTable


If Val(Left(JOurnalNo!CurrentMoYr, 2) <> Format(Date, "mm")) Then
   JOurnalNo!CurrentMoYr = Format(Date, "mm") & Trim(Right(Year(Date), 2))
   JOurnalNo!nextjn = "00001"
   JOurnalNo.Update
Else
   Jn = "PYB" & "-" & Right(Year(Date), 2) & Trim(Format(Date, "mm")) & "-" & JOurnalNo!nextjn
   nextjn = Val(JOurnalNo!nextjn)
   If (Val(nextjn)) = 1 Then
    Zeros = "0000"
    ElseIf Len(nextjn) = 2 Then
    Zeros = "000"
    ElseIf Len(nextjn) = 3 Then
    Zeros = "00"
    ElseIf Len(nextjn) = 4 Then
    Zeros = "0"
'    ElseIf Len(nextjn) = 5 Then
'    Zeros = "0"
    ElseIf Len(nextjn) = 5 Then
    Zeros = ""
   End If
   JOurnalNo!nextjn = Zeros & Trim(Val(nextjn) + 1)
   JOurnalNo.Update
   JOurnalNo.Close
End If
                     

             'save entries only to GenJOuranlTrans table
             ' rstBA.Open "PayableJournalTrans", conString, adOpenDynamic, adLockOptimistic, adCmdTable
              i = 0
              For i = 1 To FrmPayableSetup.ListView8.ListItems.Count
                  SENo = FrmPayableSetup.ListView8.ListItems.Item(i)
                  TN = FrmPayableSetup.ListView8.ListItems.Item(i).SubItems(1)
                  AccNo = FrmPayableSetup.ListView8.ListItems.Item(i).SubItems(3)
                  Accna = FrmPayableSetup.ListView8.ListItems.Item(i).SubItems(4)
                  dbamo = FrmPayableSetup.ListView8.ListItems.Item(i).SubItems(5)
                  cramo = FrmPayableSetup.ListView8.ListItems.Item(i).SubItems(6)
                  
                  'condate = FrmPayableSetup.txtToday.Text
                  condate = Date
                         
  'This is to get the Additional Details For the Reports
  AdditonDeta = "Select * from PAyablesetup where SerialNo = " & "'" & SENo & "'" & " "
  jPAy3.Open AdditonDeta, CON1, adOpenDynamic, adLockOptimistic
                           
  Do Until jPAy3.EOF = True
  jPOno = jPAy3!PoNumber
  jPOdate = jPAy3!PODate
  jInvoiceNo = jPAy3!invoiceno
  jInvDate = jPAy3!invoicedate
  jStorNo = jPAy3!StoreEntNo
  jStorDate = jPAy3!StoreEntDate
  Explanations = jPAy3!Explanation
  
  
  jPAy3.MoveNext
  Loop
    jPAy3.Close

Dim varVoucherFromAnu, varVoucherDateAnu
Dim AnuData As New ADODB.Recordset
AnuData.Open "Select * from vouchers where Payopt = '005     Payables' and optRef = " & "'" & SENo & "'" & "", constring, adOpenDynamic, adLockOptimistic

If AnuData.EOF = False Then
varVoucherFromAnu = AnuData!receiptno
varVoucherDateAnu = AnuData!receiptdate
End If

AnuData.Close

 'This is the Place I hav to giv the Journal Number
Dim y
                   
                         
                         
                         
                            With xJou
                           .AddNew
                           !serno = SENo
                           !SerialNo = Jn
                           !ticket = TN
                           !AccNo = AccNo
                           
                            
                            'Here i have to call the proceduer to got the Father Category
                            acctNo = Trim(AccNo)
                            Prevcap = Trim(Me.caption)
                            Call DisplayCats(Prevcap, acctNo, catName)
                            Categry = catName

                                                    
                           
                           !AccName = Accna
                           !recordedDate = VarForRecordedDate
                           If condate = "" Then
                           Else
                           !confirmeddate = condate
                           End If
                           
                            !PoNumber = jPOno
                            !PODate = jPOdate
                            !InvNo = jInvoiceNo
                            !InvDate = jInvDate
                            !SENumber = jStorNo
                            !SEDate = jStorDate
 
                           
                           !DBamount = Val(dbamo)
                           !CRamount = Val(cramo)
                           '!transdate = Me.Combo2 'Format(Date, "dd/mm/yyyy")
                           !Status = "Unposted"
                           !Prepby = cLogUser 'FrmPayableSetup.CmbPrepBy
                           !NotedBy = FrmPayableSetup.CmbNotedBy
                           !AppBy = FrmPayableSetup.CmbApprovedBy
                           !Classification = Categry
                           !cancelledmark = 0
                           !particulars = "Voucher #:" & SENo & ">>>" & "DV No/Date:" & varVoucherFromAnu & ">>>" & varVoucherDateAnu
                           !Description = Explanations
                           .Update
                      End With
                   
                   
                   
  jPOno = ""
  jPOdate = ""
  jInvoiceNo = ""
  jInvDate = ""
  jStorNo = ""
  jStorDate = ""

               Next
 '----
 xJou.Close
'REFRESH List 1

FrmPayableSetup.ListView1.ListItems.Clear
If KelithiMeen.EOF = False Then
KelithiMeen.MoveFirst
End If



  While KelithiMeen.EOF = False
     Set MItem = FrmPayableSetup.ListView1.ListItems.Add(, , Format(KelithiMeen!SerialNo))
     MItem.SubItems(1) = Format(KelithiMeen!Xdate, "dd/mm/yyyy")
     MItem.SubItems(2) = Format(KelithiMeen!Requester)
     MItem.SubItems(3) = Format(KelithiMeen!DateDue, "dd/mm/yyyy")
     MItem.SubItems(4) = Format(KelithiMeen!RefNo)
     MItem.SubItems(5) = Format(KelithiMeen!printmark)
     MItem.SubItems(6) = Format(KelithiMeen!journaledmark)
     MItem.SubItems(7) = Format(KelithiMeen!amtreqested, "#############.#0")
       
       Totlist1 = Val(Totlist1) + Val(Trim(KelithiMeen!amtreqested)) 'This is for the Total of the List
     KelithiMeen.MoveNext
     Wend
FrmPayableSetup.txtTotList1.Text = Totlist1

'-----


'this is to add Listview7
 FrmPayableSetup.ListView7.ListItems.Clear
 PAy4.MoveFirst
 While PAy4.EOF = False
 Dim Yaas
 Yaas = "Yes"
     Set MItem = FrmPayableSetup.ListView7.ListItems.Add(, , Format(PAy4!SerialNo))
     MItem.SubItems(1) = Format(PAy4!Xdate, "dd/mm/yyyy")
     MItem.SubItems(2) = Format(PAy4!Payee)
     MItem.SubItems(3) = Format(PAy4!DateDue)
     MItem.SubItems(4) = Format(PAy4!RefNo)
     MItem.SubItems(5) = Format(PAy4!amtreqested, "#############.#0")

       TotList7z = Val(TotList7z) + Val(Trim(PAy4!amtreqested)) 'This is for the Total of the List
     PAy4.MoveNext
     Wend
FrmPayableSetup.txtTotList7.Text = Trim(TotList7z)



Dim XJou2 As New ADODB.Recordset
XJou2.Open "select * from Payjournal where status = 'Unposted'", constring, adOpenDynamic, adLockOptimistic


'This is to Refresh ListView 8
FrmPayableSetup.ListView8.ListItems.Clear

If XJou2.EOF = False Then
XJou2.MoveFirst
End If

  While XJou2.EOF = False
    If XJou2!cancelledmark = "0" Then

  
     Set MItem = FrmPayableSetup.ListView8.ListItems.Add(, , Format(XJou2!SerialNo))
     MItem.SubItems(1) = Format(XJou2!ticket)
     MItem.SubItems(2) = Format(XJou2!confirmeddate, "dd/mm/yyyy")
     MItem.SubItems(3) = Format(XJou2!AccNo)
     MItem.SubItems(4) = Format(XJou2!AccName)
     MItem.SubItems(5) = Format(XJou2!DBamount)
     MItem.SubItems(6) = Format(XJou2!CRamount)
                                                        
       TotList8Db = Val(TotList8Db) + Val(Trim(XJou2!DBamount)) 'This is for the Total of the List
       totlist8cr = Val(totlist8cr) + Val(Trim(XJou2!CRamount)) 'This is for the Total of the List
     
   End If
     XJou2.MoveNext
     Wend
     
FrmPayableSetup.txtJdb.Text = Trim(TotList8Db)
FrmPayableSetup.txtJCr.Text = Trim(totlist8cr)


  MsgBox "Datas Confirmed Successfullly"
  FrmPayableSetup.SSTab1.SetFocus
  SendKeys "{left}"
  SendKeys "{left}"

End If 'for msgbox


End Sub

Private Sub conPCL_Click()
DRPcancelled.Show 1
End Sub

Private Sub del_Click()
If FrmPayableSetup.ListView2.ListItems.Count = 0 Then
MsgBox "No items selected"
Exit Sub
End If

'Delete items in the ListV (Terms)
   xindex = FrmPayableSetup.ListView2.SelectedItem.Index
   xrate = FrmPayableSetup.ListView2.SelectedItem
   xdes = FrmPayableSetup.ListView2.SelectedItem.SubItems(1)
   xday = FrmPayableSetup.ListView2.SelectedItem.SubItems(2)
   xmode = FrmPayableSetup.ListView2.SelectedItem.SubItems(3)
   xlevel = FrmPayableSetup.ListView2.SelectedItem.SubItems(4)
        FrmPayableSetup.ListView2.ListItems.Remove xindex

'Delete From the File Permanently (Terms)
        If rstterm.EOF = False Then
        rstterm.MoveFirst
        End If
        
        While rstterm.EOF = False
        
        If FrmPayableSetup.txtSerialNo.Text = (rstterm!SerialNo) And xrate = (rstterm!Rate) And xdes = (rstterm!descr) And xmode = (rstterm!Mode) Then
        rstterm.Delete
        
        MsgBox "Records Successfully Deleted"
        
        End If
        
  rstterm.MoveNext
  Wend
End Sub

Private Sub delet_Click()

End Sub

Private Sub deleateList_Click()




If FrmPayableSetup.ListView4.ListItems.Count = 0 Then
MsgBox "No items To Select"
Exit Sub
End If

Dim CON1 As New ADODB.Connection
conStr = "Provider=MSDASQL;DSN=Finance;UID=sa; PWD=;"
CON1.Open conStr

xyes = MsgBox("Mr." & cLogUser & ",Are You Sure You Want to Delete", vbYesNo + vbQuestion, "Deleting record")
If xyes = vbYes Then

'Delete items in the ListV
        If FrmPayableSetup.ListView4.ListItems.Count = 0 Then
        frmMenu.deleateList.Enabled = False
        Exit Sub
        End If
  vindex = FrmPayableSetup.ListView4.SelectedItem.Index
   vItem = FrmPayableSetup.ListView4.SelectedItem
     
        FrmPayableSetup.ListView4.ListItems.Remove vindex

'Delete From the File Permanently(Put the Deletemark LATER)
    Dim rstPay22 As New ADODB.Recordset
    rstPay22.Open "Select * from Payablesetup", CON1, adOpenDynamic, adLockOptimistic

If rstPay22.EOF = False Then
rstPay22.MoveFirst
End If

        
        While rstPay22.EOF = False
        
        If vItem = (rstPay22!SerialNo) Then
        'rstpay22.Delete
        rstPay22!deletemark = 1
        rstPay22!DeleUser = FrmPayableSetup.CmbPrepBy
        MsgBox "Records deleted"
        Exit Sub
        End If
        
  rstPay22.MoveNext
  Wend
 Unload Me
'End If
Exit Sub
End If 'This is the end for Deletion

End Sub

Private Sub DoJou_Click()

'---------------------------------------------
'This is for Journal Only
'---------------------------------------------

Dim TickNo As String

Dim jrstConf1 As ADODB.Recordset
Set jrstConf1 = New ADODB.Recordset

Dim jrstConf2 As ADODB.Recordset
Set jrstConf2 = New ADODB.Recordset

Dim jxJou As ADODB.Recordset
Set jxJou = New ADODB.Recordset

Dim jPay As ADODB.Recordset
Set jPay = New ADODB.Recordset

Dim jKelithiMeen As ADODB.Recordset
Set jKelithiMeen = New ADODB.Recordset
Dim jjPOno, jjPOdate, jjInvoiceNo, jjInvDate, jjStorNo, jjStorDate
Dim jPAy3 As New ADODB.Recordset
Dim jPAy4 As New ADODB.Recordset

Dim jAdditonDeta
If FrmPayableSetup.ListView1.ListItems.Count = 0 Then
MsgBox "no item selected"
Exit Sub
Else
End If

Dim jxSaman
jxSaman = FrmPayableSetup.ListView1.SelectedItem


'H E R E    I  H A V E   T 0  D O  C H E C K WHETHER ALREADY JOURNALMARK = "yES" IN THE TABLE AND MSGBOX
If FrmPayableSetup.ListView1.SelectedItem.SubItems(6) = "Yes" Then
MsgBox "This is Already Journalysed"
Exit Sub
End If

Dim Chek As New ADODB.Recordset
Chek.Open "SELECT * From Payablesetup  Where serialno =" & "'" & jxSaman & "'" & " and JournaledMark ='Yes'", constring, adOpenDynamic, adLockOptimistic
If Chek.EOF = False Then
MsgBox "Already Journalised"
Exit Sub
End If




ms = MsgBox("Mr." & cLogUser & ", Do you want to Make the Journal", vbYesNo, " Journalyse")

If ms = vbNo Then
Exit Sub
Else


jxpay = "SELECT * From Xpayment  Where serialno =" & "'" & jxSaman & "'" & " and confirmedmark = 'uc' and Postmark = 'No'  order by serialno"
jxrec = "SELECT * From Xreceipt  Where serialno =" & "'" & jxSaman & "'" & " and confirmedmark = 'uc' and Postmark = 'No'  order by serialno"

jxset = "SELECT * From Payablesetup  Where serialno =" & "'" & jxSaman & "'" & " and confirmedmark = 0 and Deletemark = 0 order by serialno"
'xset2 = "SELECT * From Payablesetup  Where  confirmedmark = 0 order by serialno"
jrstConf1.Open jxpay, constring, adOpenDynamic, adLockOptimistic, adCmdText
jrstConf2.Open jxrec, constring, adOpenDynamic, adLockOptimistic, adCmdText
jPay.Open jxset, constring, adOpenDynamic, adLockOptimistic, adCmdText

SuraMeen = "Select * from Payablesetup where cancelledmark = '0' and deletemark = '0' and confirmedmark = '0'  and Post = 'No' and Paidmark = '0'"
jKelithi = "Select * from Payablesetup where cancelledmark = '0' and deletemark = '0' and confirmedmark = '1'  and Post = 'No' and Paidmark = '0'"
jKelithiMeen.Open SuraMeen, constring, adOpenDynamic, adLockOptimistic
jPAy4.Open jKelithi, constring, adOpenDynamic, adLockOptimistic

jxJou.Open "select * from Payjournal ", constring, adOpenDynamic, adLockOptimistic
If jPay.EOF = True Then
On Error Resume Next
End If


'This is to check the xPayment and XReceipt whether they did the Payment Analysis
'before confirming it


If jrstConf2.EOF = True Then
Hel = MsgBox("Mr." & cLogUser & ",You Have to Do the Payment Analysis to Jornalise", vbInformation, "Message")
Exit Sub
End If

TickNo = 1

If jrstConf1.EOF = True Then
On Error Resume Next
End If

'this is to add To the     to Listview 8

FrmPayableSetup.ListView8.ListItems.Clear

If jrstConf1.EOF = False Then
jrstConf1.MoveFirst
End If

  While jrstConf1.EOF = False
     Set MItem = FrmPayableSetup.ListView8.ListItems.Add(, , Format(jrstConf1!SerialNo))
     MItem.SubItems(1) = Trim(TickNo)
     MItem.SubItems(2) = Date
     MItem.SubItems(3) = Format(jrstConf1!AccNo)
     MItem.SubItems(4) = Format(jrstConf1!AccName)
     MItem.SubItems(5) = Format(jrstConf1!amount)
     'MItem.SubItems(5) = Format(jrstConf1!amount)
TickNo = TickNo + 1
       TotList8Dbx = Val(TotList8Dbx) + Val(Trim(jrstConf1!amount)) 'This is for the Total of the List

   jrstConf1.MoveNext
   Wend
   

If jrstConf2.EOF = True Then
On Error Resume Next
End If

   
If jrstConf2.EOF = False Then
jrstConf2.MoveFirst
End If

   Dim VarForRecordedDate As Date
   
'this is for the Receipt
  While jrstConf2.EOF = False
     Set MItem = FrmPayableSetup.ListView8.ListItems.Add(, , Format(jrstConf2!SerialNo))
     MItem.SubItems(1) = Trim(TickNo)
     MItem.SubItems(2) = Date
     MItem.SubItems(3) = Format(jrstConf2!AccNo)
     MItem.SubItems(4) = Format(jrstConf2!AccName)
     MItem.SubItems(6) = Format(jrstConf2!amount)
   TickNo = TickNo + 1
   
      TotList8crx = Val(TotList8crx) + Val(Trim(jrstConf2!amount)) 'This is for the Total of the List
 VarForRecordedDate = Format(jrstConf2!recordedDate, "dd/mm/yyyy")
   jrstConf2.MoveNext
   Wend


FrmPayableSetup.txtJdb.Text = Trim(TotList8Dbx)
FrmPayableSetup.txtJCr.Text = Trim(TotList8crx)


'Add new from Listview to the Table Payjournal

Dim JOurnalNo As New ADODB.Recordset
JOurnalNo.Open "Setup", constring, adOpenKeyset, adLockPessimistic, adCmdTable


If Val(Left(JOurnalNo!CurrentMoYr, 2) <> Format(Date, "mm")) Then
   JOurnalNo!CurrentMoYr = Format(Date, "mm") & Trim(Right(Year(Date), 2))
   JOurnalNo!nextjn = "00001"
   JOurnalNo.Update
Else
   Jn = "PYB" & "-" & Right(Year(Date), 2) & Trim(Format(Date, "mm")) & "-" & JOurnalNo!nextjn
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
   JOurnalNo.Update
   JOurnalNo.Close
End If
                      

             'save entries only to GenJOuranlTrans table
             ' rstBA.Open "PayableJournalTrans", conString, adOpenDynamic, adLockOptimistic, adCmdTable
              i = 0
              For i = 1 To FrmPayableSetup.ListView8.ListItems.Count
                  SENo = FrmPayableSetup.ListView8.ListItems.Item(i)
                  TN = FrmPayableSetup.ListView8.ListItems.Item(i).SubItems(1)
                  AccNo = FrmPayableSetup.ListView8.ListItems.Item(i).SubItems(3)
                  Accna = FrmPayableSetup.ListView8.ListItems.Item(i).SubItems(4)
                  dbamo = FrmPayableSetup.ListView8.ListItems.Item(i).SubItems(5)
                  cramo = FrmPayableSetup.ListView8.ListItems.Item(i).SubItems(6)
                  
                  'condate = FrmPayableSetup.txtToday.Text
                         condate = Date
                         
  'This is to get the Additional Details For the Reports
  jAdditonDeta = "Select * from PAyablesetup where SerialNo = " & "'" & SENo & "'" & " "
  jPAy3.Open jAdditonDeta, CON1, adOpenDynamic, adLockOptimistic
                           
  Do Until jPAy3.EOF = True
  jjPOno = jPAy3!PoNumber
  jjPOdate = jPAy3!PODate
  jjInvoiceNo = jPAy3!invoiceno
  jjInvDate = jPAy3!invoicedate
  jjStorNo = jPAy3!StoreEntNo
  jjStorDate = jPAy3!StoreEntDate
   Explanations = jPAy3!Explanation

  jPAy3.MoveNext
  Loop
    jPAy3.Close



Dim varVoucherFromAnu2, varVoucherDateAnu2
Dim AnuData2 As New ADODB.Recordset
AnuData2.Open "Select * from vouchers where Payopt = '005     Payables' and optRef = " & "'" & SENo & "'" & "", constring, adOpenDynamic, adLockOptimistic

If AnuData2.EOF = False Then
varVoucherFromAnu2 = AnuData2!receiptno
varVoucherDateAnu2 = AnuData2!receiptdate
End If

AnuData2.Close




 'This is the Place I hav to giv the Journal Number
                         
                   
                         
                         
                         
                            With jxJou
                           .AddNew
                           !serno = SENo
                           !SerialNo = Jn
                           !ticket = TN
                           !AccNo = AccNo
                           
                           
                                                    
                            'Here i have to call the proceduer to got the Father Category
                            acctNo = Trim(AccNo)
                            Prevcap = Trim(Me.caption)
                            Call DisplayCats(Prevcap, acctNo, catName)
                            Categry = catName
   
                           
                           !AccName = Accna
                           !recordedDate = VarForRecordedDate
                           If condate = "" Then
                           Else
                           !confirmeddate = condate
                           End If
                           
                            !PoNumber = jjPOno
                            !PODate = jjPOdate
                            !InvNo = jjInvoiceNo
                            !InvDate = jjInvDate
                            !SENumber = jjStorNo
                            !SEDate = jjStorDate
 
                           
                           !DBamount = Val(dbamo)
                           !CRamount = Val(cramo)
                           '!transdate = Me.Combo2 'Format(Date, "dd/mm/yyyy")
                           !Status = "Unposted"
                           !Prepby = cLogUser 'FrmPayableSetup.CmbPrepBy
                           !NotedBy = FrmPayableSetup.CmbNotedBy
                           !Classification = Categry
                           !AppBy = FrmPayableSetup.CmbApprovedBy
                           !particulars = "Voucher #:" & SENo & ">>>" & "DV No/Date:" & varVoucherFromAnu2 & ">>>" & varVoucherDateAnu2
                           !Description = Explanations

                           .Update
                      End With
                   
                   
                   
  jjPOno = ""
  jjPOdate = ""
  jjInvoiceNo = ""
  jjInvDate = ""
  jjStorNo = ""
  jjStorDate = ""

               Next
 '----
 jxJou.Close

'REFRESH List 1

Dim RstPayforJournalMark As New ADODB.Recordset
Dim VarMark
VarMark = "SELECT * From Payablesetup  Where serialno =" & "'" & jxSaman & "'" & " and confirmedmark = 0 order by serialno"

RstPayforJournalMark.Open VarMark, CON1, adOpenDynamic, adLockOptimistic

If RstPayforJournalMark.EOF = False Then
RstPayforJournalMark.MoveFirst
End If

While RstPayforJournalMark.EOF = False

RstPayforJournalMark!journaledmark = "Yes"
RstPayforJournalMark.MoveNext
Wend




FrmPayableSetup.ListView1.ListItems.Clear
If jKelithiMeen.EOF = False Then
jKelithiMeen.MoveFirst
End If



  While jKelithiMeen.EOF = False
     Set MItem = FrmPayableSetup.ListView1.ListItems.Add(, , Format(jKelithiMeen!SerialNo))
     MItem.SubItems(1) = Format(jKelithiMeen!Xdate, "dd/mm/yyyy")
     MItem.SubItems(2) = Format(jKelithiMeen!Requester)
     MItem.SubItems(3) = Format(jKelithiMeen!DateDue, "dd/mm/yyyy")
     MItem.SubItems(4) = Format(jKelithiMeen!RefNo)
     MItem.SubItems(5) = Format(jKelithiMeen!printmark)
     MItem.SubItems(6) = Format(jKelithiMeen!journaledmark)
     MItem.SubItems(7) = Format(jKelithiMeen!amtreqested, "#############.#0")

       Totlist1 = Val(Totlist1) + Val(Trim(jKelithiMeen!amtreqested)) 'This is for the Total of the List
     jKelithiMeen.MoveNext
     Wend
FrmPayableSetup.txtTotList1.Text = Totlist1

'-----


'this is to add Listview7
 FrmPayableSetup.ListView7.ListItems.Clear
 jPAy4.MoveFirst
 While jPAy4.EOF = False
 Dim Yaas
 Yaas = "Yes"
     Set MItem = FrmPayableSetup.ListView7.ListItems.Add(, , Format(jPAy4!SerialNo))
     MItem.SubItems(1) = Format(jPAy4!Xdate, "dd/mm/yyyy")
     MItem.SubItems(2) = Format(jPAy4!Payee)
     MItem.SubItems(3) = Format(jPAy4!DateDue)
     MItem.SubItems(4) = Format(jPAy4!RefNo)
     MItem.SubItems(5) = Format(jPAy4!amtreqested, "#############.#0")

       TotList7z = Val(TotList7z) + Val(Trim(jPAy4!amtreqested)) 'This is for the Total of the List
     jPAy4.MoveNext
     Wend
FrmPayableSetup.txtTotList7.Text = Trim(TotList7z)



Dim XJou2 As New ADODB.Recordset
XJou2.Open "select * from Payjournal where status = 'Unposted'", constring, adOpenDynamic, adLockOptimistic


'This is to Refresh ListView 8
FrmPayableSetup.ListView8.ListItems.Clear

If XJou2.EOF = False Then
XJou2.MoveFirst
End If

  While XJou2.EOF = False
  
     Set MItem = FrmPayableSetup.ListView8.ListItems.Add(, , Format(XJou2!SerialNo))
     MItem.SubItems(1) = Format(XJou2!ticket)
     MItem.SubItems(2) = Format(XJou2!confirmeddate, "dd/mm/yyyy")
     MItem.SubItems(3) = Format(XJou2!AccNo)
     MItem.SubItems(4) = Format(XJou2!AccName)
     MItem.SubItems(5) = Format(XJou2!DBamount)
     MItem.SubItems(6) = Format(XJou2!CRamount)
                                                        
       TotList8Db = Val(TotList8Db) + Val(Trim(XJou2!DBamount)) 'This is for the Total of the List
       totlist8cr = Val(totlist8cr) + Val(Trim(XJou2!CRamount)) 'This is for the Total of the List
     XJou2.MoveNext
     Wend
     
FrmPayableSetup.txtJdb.Text = Trim(TotList8Db)
FrmPayableSetup.txtJCr.Text = Trim(totlist8cr)


  MsgBox "Datas Confirmed Successfullly"
  FrmPayableSetup.SSTab1.SetFocus
  SendKeys "{left}"
  SendKeys "{left}"

End If 'for msgbox



End Sub

Private Sub edit_Click()
 
 FrmPayableSetup.TxtKallapundai.Text = FrmPayableSetup.ListView2.SelectedItem.Index

On Error Resume Next
If FrmPayableSetup.ListView2.SelectedItem.Text = "" Then
MsgBox "listview is Empty or No item Selected", vbInformation, "Edit"
Exit Sub
End If
        
  If Me.Edit.caption = "Edit" Then  'This is to Edit So it will bring all the ListView Datas to the Particular TextBoxes
        FrmPayableSetup.txtTmRate.Text = FrmPayableSetup.ListView2.SelectedItem.Text
        FrmPayableSetup.txtTMDes.Text = FrmPayableSetup.ListView2.SelectedItem.SubItems(1)
        FrmPayableSetup.txtTMDays.Text = FrmPayableSetup.ListView2.SelectedItem.SubItems(2)
        FrmPayableSetup.txtTMmode.Text = FrmPayableSetup.ListView2.SelectedItem.SubItems(3)
        FrmPayableSetup.txtTMlevel.Text = FrmPayableSetup.ListView2.SelectedItem.SubItems(4)
        
        FrmPayableSetup.txtTmp1.Text = FrmPayableSetup.txtTmRate.Text
        FrmPayableSetup.txtTmp2.Text = FrmPayableSetup.txtTMDays.Text
      
      Me.Edit.caption = "Update"
      Me.Clear.caption = "Cancel" 'This will Enable Internal Edit Again"

FrmPayableSetup.txtTMDays.Visible = True
FrmPayableSetup.txtTMDes.Visible = True
FrmPayableSetup.txtTMlevel.Visible = True
FrmPayableSetup.txtTMmode.Visible = True
FrmPayableSetup.txtTmRate.Visible = True

Else ' this is to update the textBox's Datas to the Database
   
 If FrmPayableSetup.CmbPrepBy.Text = "" Then
MsgBox "fill the Combo 'Prepared By' "

FrmPayableSetup.CmbPrepBy.SetFocus
Exit Sub
End If

If rstterm.EOF = False Then
rstterm.MoveFirst
End If

   
      While rstterm.EOF = False
       Dim xvar As String
       xvar = Trim(rstterm!Rate)
         If Trim(rstterm!SerialNo) = Trim(FrmPayableSetup.txtSerialNo.Text) And ((xvar) = FrmPayableSetup.txtTmp1.Text) And ((rstterm!days) = (FrmPayableSetup.txtTmp2.Text)) Then
        rstterm!SerialNo = FrmPayableSetup.txtSerialNo.Text
        rstterm!Rate = FrmPayableSetup.txtTmRate.Text
        rstterm!descr = FrmPayableSetup.txtTMDes.Text
        rstterm!days = FrmPayableSetup.txtTMDays.Text
        rstterm!Mode = FrmPayableSetup.txtTMmode.Text
        rstterm!xlevel = FrmPayableSetup.txtTMlevel.Text
        
        Me.Edit.caption = "&Edit"
      End If
   rstterm.MoveNext
   Wend



'..This is to REFRESH the ListView, after Adding New
   FrmPayableSetup.ListView2.ListItems.Clear
   
If rstterm.EOF = False Then
rstterm.MoveFirst
End If

   While rstterm.EOF = False
       
            If Trim(rstterm!SerialNo) = Trim(FrmPayableSetup.txtSerialNo.Text) Then
        
     Set MItem = FrmPayableSetup.ListView2.ListItems.Add(, , Trim(rstterm!Rate))
     MItem.SubItems(1) = Trim(rstterm!descr)
     MItem.SubItems(2) = Trim(rstterm!days)
     MItem.SubItems(3) = Trim(rstterm!Mode)
     MItem.SubItems(4) = Trim(rstterm!xlevel)
        End If
     rstterm.MoveNext
     Wend
'........................................



End If
End Sub

Private Sub Form_Load()
Set CON1 = New ADODB.Connection
Set rstList3 = New ADODB.Recordset
Set RstItem = New ADODB.Recordset
Set rstPurch = New ADODB.Recordset
Set rstReceipt = New ADODB.Recordset

conStr = "Provider=MSDASQL;DSN=Finance;UID=sa; PWD=;"
CON1.Open conStr
Set rstterm = New ADODB.Recordset
rstterm.Open "Select * from term", CON1, adOpenDynamic, adLockOptimistic
rstList3.Open "Select * from xpayment", CON1, adOpenDynamic, adLockOptimistic
RstItem.Open "Select * from Purchaseitem", CON1, adOpenDynamic, adLockOptimistic
rstPurch.Open "Select * from PurchaseSetup", CON1, adOpenDynamic, adLockOptimistic
rstReceipt.Open "Select * from xReceipt", CON1, adOpenDynamic, adLockOptimistic
End Sub

Private Sub IEdit_Click()
On Error Resume Next

If frmPurchaseSetup.ListView2.SelectedItem.Text = "" Then
MsgBox "listview is Empty or No item Selected", vbInformation, "Edit"
End If



'this is to EDIT
If Me.IEdit.caption = "Edit" Then
        frmPurchaseSetup.txtItemCode.Text = frmPurchaseSetup.ListView2.SelectedItem.Text
        frmPurchaseSetup.TxtItemDes.Text = frmPurchaseSetup.ListView2.SelectedItem.SubItems(1)
        frmPurchaseSetup.txtItemModelNo.Text = frmPurchaseSetup.ListView2.SelectedItem.SubItems(2)
        frmPurchaseSetup.txtItemDiam.Text = frmPurchaseSetup.ListView2.SelectedItem.SubItems(3)
        frmPurchaseSetup.txtItemQty.Text = frmPurchaseSetup.ListView2.SelectedItem.SubItems(4)
        frmPurchaseSetup.txtItemPrize.Text = frmPurchaseSetup.ListView2.SelectedItem.SubItems(5)
        frmPurchaseSetup.txtSurTax.Text = frmPurchaseSetup.ListView2.SelectedItem.SubItems(6)
        frmPurchaseSetup.txtVAt.Text = frmPurchaseSetup.ListView2.SelectedItem.SubItems(7)
        frmPurchaseSetup.txtTaxCredit.Text = frmPurchaseSetup.ListView2.SelectedItem.SubItems(8)
        frmPurchaseSetup.txtItemInvInn.Text = frmPurchaseSetup.ListView2.SelectedItem.SubItems(9)

        
        Me.IEdit.caption = "Update"
        Me.IClear.caption = "Cancel" 'This will Enable Internal Edit Again"

frmPurchaseSetup.txtItemTemp1List1.Text = frmPurchaseSetup.txtItemCode.Text
'frmPurchaseSetup.txtTemp2List3.Text = frmPurchaseSetup.txtPartic.Text
'
    
    
'This is  "Updating"
ElseIf Me.IEdit.caption = "Update" Then


If frmPurchaseSetup.CmbPrepBy.Text = "" Then
MsgBox "fill the Combo 'Prepared By' "

FrmPayableSetup.CmbPrepBy.SetFocus
Exit Sub
End If

'This is to identify that the Updation is for List3 from frmPassword
frmPassword.txtBuffer.Text = "UpdateItemList2"
frmPassword.txtUserId.Text = frmPurchaseSetup.CmbPrepBy.Text
frmPassword.caption = "Enter Password to Update"
frmPassword.txtPrepBy = frmPurchaseSetup.CmbPrepBy.Text

frmPassword.Show 1
End If

End Sub

Private Sub JENT_Click()
If FrmPayableSetup.ListView8.ListItems.Count = 0 Then
MsgBox "No items"
Exit Sub
End If

Dim VarList
VarList = FrmPayableSetup.ListView8.SelectedItem

On Error Resume Next
DataEnvironment1.rsPayJournalLast.Close
DataEnvironment1.PayJournalLast VarList

PayJournalLast2.Show


End Sub
Private Sub ProcedByDate(a As RptLabel)
a.caption = mskDueDate.Text
End Sub
Private Sub ProcedByDate2(X As RptLabel)
X.caption = cmbPayee.Text
End Sub

Private Sub ProcedPrepBy(b As RptLabel)
b.caption = cLogUser
End Sub
Private Sub ProcedPReqNo(rNo As RptLabel)
rNo.caption = PettyReqNo
End Sub
Private Sub ProcedDatex(Dtx As RptLabel)
Dtx.caption = PettyDatex
End Sub
Private Sub procePrintBy(PrBy As RptLabel)
PrBy.caption = cLogUser
End Sub


Private Sub JPr_Click()
'RepUnposted.Show 1
MsgBox "Please Goto the Main Menu, Click the PYB-Payable setup , Right click the Date and Print the Report", vbInformation, "HELP"

End Sub

Private Sub PaidDueDate_Click()
        On Error Resume Next
       ' DataEnvironment1.rsUnpaidByDueDate.Close
       ' DataEnvironment1.UnpaidByDueDate VarList
        
         ' ProcedByDate UnpaidByDateDue.Sections(1).Controls("label3")
          ProcedPrepBy paidByDateDue.Sections(2).Controls("lblPrepby")
        
        paidByDateDue.Show 1

End Sub

Private Sub PayCancel_Click()
If FrmPayableSetup.ListView1.SelectedItem.Text = "" Then
MsgBox "listview is Empty or No item Selected", vbInformation, "Edit"

End If

If FrmPayableSetup.CmbPrepBy.Text = "" Then
MsgBox "Fill the Combo Prepared by", vbInformation, "Try Again"
End If

xconf = MsgBox("You are about Cancelling the Payment, Are You Sure", vbQuestion + vbOKCancel, "Confirmation")
 If xconf = vbCancel Then
 Exit Sub
 Else 'Vb OK


'This is to identify that the Updation is for List3 from frmPassword
frmPassword.txtBuffer.Text = "Payable cancelation"
frmPassword.caption = "Enter Password to Cancel the Payemnt"
frmPassword.txtPrepBy = FrmPayableSetup.CmbPrepBy.Text

frmPassword.Show 1
End If
End Sub

Private Sub PCconCancel_Click()
If frmPettyCash.LvwPCashConf.ListItems.Count = 0 Then
MsgBox "List is Empty", vbInformation
Exit Sub
End If

Dim VarCanPCLstc, varLovec
Dim CanPCLstc As New ADODB.Recordset
varLovecc = frmPettyCash.LvwPCashConf.SelectedItem


Mymsgx = MsgBox("Mr." & cLogUser & ",Do you want to Cancell the Record  " & varLovecc & "", vbInformation + vbYesNo, "Please Confirm")
If Mymsgx = vbNo Then
Exit Sub
End If


CanPCLstc.Open "SElect * from pettycashheld where serialno = " & "'" & varLovecc & "'" & "", constring, adOpenDynamic, adLockOptimistic

While CanPCLstc.EOF = False
CanPCLstc!cancelledmark = 1
CanPCLstc.MoveNext
Wend

 'This is to call the Class for the Listview confirmed
Dim xclsr As New HabitatClass
Dim rs5 As New ADODB.Recordset
'xclsr.ProcLvwPCashConfirmed rs5

'This is to call the Class for the Listview UN confirm
Dim rs55 As New ADODB.Recordset
'xclsr.ProcLvwPCashUnCon rs55

'This is to call the Class for the Listview UN confirm
Dim rs555 As New ADODB.Recordset
'xclsr.ProcLvwPCJournal rs555

 MsgBox "Records of " & varLovecc & "  has been Cancelled Successfully", vbInformation, "Confirmation"

End Sub

Private Sub PCconDelete_Click()
If frmPettyCash.LvwPCashConf.ListItems.Count = 0 Then
MsgBox "List is Empty", vbInformation
Exit Sub
End If

Dim VarDelPCLstc, varLov
Dim DelPCLstc As New ADODB.Recordset
varLov = frmPettyCash.LvwPCashConf.SelectedItem


Mymsgx = MsgBox("Mr." & cLogUser & ",Do you want to Delete " & varLov & " from the Record", vbInformation + vbYesNo, "Please Confirm")
If Mymsgx = vbNo Then
Exit Sub
End If

DelPCLstc.Open "Delete from pettycashheld where serialno = " & "'" & varLov & "'" & "", constring, adOpenDynamic, adLockOptimistic



 'This is to call the Class for the Listview confirmed
Dim xclss As New HabitatClass
Dim rsets As New ADODB.Recordset
'xclss.ProcLvwPCashConfirmed rsets


'This is to call the Class for the Listview UN confirm
Dim rsetJ As New ADODB.Recordset
'xclss.ProcLvwPCJournal rsetJ

MsgBox "Records of " & varLov & " Deleted", vbInformation, "Confirmation"


End Sub

Private Sub PCuconCancell_Click()
If frmPettyCash.LvwPCashUnConf.ListItems.Count = 0 Then
MsgBox "List is Empty", vbInformation
Exit Sub
End If


Dim VarCanPCLstUc, varLovec
Dim CanPCLstUc As New ADODB.Recordset
varLovec = frmPettyCash.LvwPCashUnConf.SelectedItem



Mymsgx = MsgBox("Mr." & cLogUser & ",Do you want to Cancell the Record " & varLovec & "", vbInformation + vbYesNo, "Please Confirm")
If Mymsgx = vbNo Then
Exit Sub
End If

CanPCLstUc.Open "SElect * from pettycashheld where serialno = " & "'" & varLovec & "'" & "", constring, adOpenDynamic, adLockOptimistic

While CanPCLstUc.EOF = False
CanPCLstUc!cancelledmark = 1
CanPCLstUc.MoveNext
Wend

 'This is to call the Class for the Listview confirmed
Dim xclsr As New HabitatClass
Dim rs5 As New ADODB.Recordset
'xclsr.ProcLvwPCashConfirmed rs5

'This is to call the Class for the Listview UN confirm
Dim rs55 As New ADODB.Recordset
'xclsr.ProcLvwPCashUnCon rs55

'This is to call the Class for the Listview UN confirm
Dim rs555 As New ADODB.Recordset
'xclsr.ProcLvwPCJournal rs555
 
 MsgBox "Records of " & varLovec & "  has been Cancelled Successfully", vbInformation, "Confirmation"
 

End Sub

Private Sub PCuconDelete_Click()

If frmPettyCash.LvwPCashUnConf.ListItems.Count = 0 Then
MsgBox "List is Empty", vbInformation
Exit Sub
End If


Mymsgx = MsgBox("Mr." & cLogUser & ",Do you want to Delete from the Record", vbInformation + vbYesNo, "Please Confirm")
If Mymsgx = vbNo Then
Exit Sub
End If

Dim VarDelPCLstUc, varLove
Dim DelPCLstUc As New ADODB.Recordset
varLove = frmPettyCash.LvwPCashUnConf.SelectedItem
DelPCLstUc.Open "Delete from pettycashheld where serialno = " & "'" & varLove & "'" & "", constring, adOpenDynamic, adLockOptimistic



'This is to Remove from the ListViewUC
Dim LVWPCIndex3, TotL3
TotL3 = frmPettyCash.LvwPCashUnConf.SelectedItem.SubItems(4)

LVWPCIndex3 = frmPettyCash.LvwPCashUnConf.SelectedItem.Index
frmPettyCash.LvwPCashUnConf.ListItems.Remove (LVWPCIndex3)
frmPettyCash.txttotUc.Text = Val(frmPettyCash.txttotUc.Text) - Val(TotL3)



End Sub

Private Sub PCuconEdit_Click()

If frmPettyCash.LvwPCashUnConf.ListItems.Count = 0 Then
MsgBox "Sorry, List is Empty,You can not edit this", vbInformation
Exit Sub
End If

    frmPettyCash.txtPaidto.Enabled = True
    frmPettyCash.txtPartic.Enabled = True
    frmPettyCash.txtExpl.Enabled = True
    frmPettyCash.txtInvNo.Enabled = True
   ' frmPettyCash.txtPoNo.Enabled = True
   ' frmPettyCash.mskPODate.Enabled = True
    
    frmPettyCash.CMBAccName.Enabled = True
    frmPettyCash.CMBAccNo.Enabled = True
    frmPettyCash.cmbAccName2.Enabled = True
    frmPettyCash.cmbAccNo2.Enabled = True
    frmPettyCash.txtAmt.Enabled = True
    frmPettyCash.txtSerialNo.Enabled = True
    frmPettyCash.mskDate.Enabled = True
    frmPettyCash.cmdedit.caption = "&Update"


Dim VarUIx, VarrsTake2LIstx, totCRforLVWPCx
VarUIx = frmPettyCash.LvwPCashUnConf.SelectedItem


Dim rsTake2LIstx As New ADODB.Recordset
 
frmPettyCash.txtSerialNo.Text = VarUIx
VarrsTake2LIstx = "Select * from pettycashHeld where SErialno = " & "'" & VarUIx & "'" & ""

rsTake2LIstx.Open VarrsTake2LIstx, constring, adOpenDynamic, adLockOptimistic


If rsTake2LIstx.EOF = False Then
 rsTake2LIstx.MoveFirst
End If

frmPettyCash.CMBAccName.Text = rsTake2LIstx!accountname
frmPettyCash.CMBAccNo.Text = rsTake2LIstx!AccountNo
frmPettyCash.mskDate.Text = rsTake2LIstx!Datex
frmPettyCash.LVWPettyCash.ListItems.Clear
  While rsTake2LIstx.EOF = False
    ' Set mitem = frmPettycash.LVWPettyCash.ListItems.Add(, , Format(rsTake2LIstx!serialno))
     Set MItem = frmPettyCash.LVWPettyCash.ListItems.Add(, , Trim(rsTake2LIstx!PaidTo))
    ' mitem.SubItems(1) = Trim(rsTake2LIstx!PaidTo)
     MItem.SubItems(1) = Trim(rsTake2LIstx!Partic)
     MItem.SubItems(2) = Trim(rsTake2LIstx!Explan)
     MItem.SubItems(3) = IIf(IsNull(rsTake2LIstx!InvNo), "", (rsTake2LIstx!InvNo))
     MItem.SubItems(4) = IIf(IsNull(rsTake2LIstx!PoNumber), "", (rsTake2LIstx!PoNumber))
     
     MItem.SubItems(5) = IIf(IsNull(rsTake2LIstx!PODate), "", (rsTake2LIstx!PODate))
     MItem.SubItems(6) = IIf(IsNull(rsTake2LIstx!AccountNo2), "", (rsTake2LIstx!AccountNo2))
     MItem.SubItems(7) = Trim(rsTake2LIstx!accountname2)
     MItem.SubItems(8) = Format(IIf(IsNull(rsTake2LIstx!amount), "", (rsTake2LIstx!amount)), "############.#0")
     MItem.SubItems(9) = Format(IIf(IsNull(rsTake2LIstx!creditamount), "", (rsTake2LIstx!creditamount)), "############.#0")
     
     totCRforLVWPC = Val(totCRforLVWPC) + Val(IIf(IsNull(rsTake2LIstx!amount), 0, (rsTake2LIstx!amount)))
     totCRforLVWPC2 = Val(totCRforLVWPC2) + Val(IIf(IsNull(rsTake2LIstx!creditamount), 0, (rsTake2LIstx!creditamount)))
     
     rsTake2LIstx.MoveNext
     Wend
 frmPettyCash.txttotal.Text = totCRforLVWPC
  frmPettyCash.txtTotCr.Text = totCRforLVWPC2

  frmPettyCash.SSTab1.SetFocus
  SendKeys "{Left}"

End Sub

Private Sub PCuconPrint_Click()


'This is to Delete the Existing Datas

Dim rsPettPrDel As New ADODB.Recordset
rsPettPrDel.Open "Delete from pettyprintCr", constring, adOpenDynamic, adLockOptimistic

Dim rsPettPrDel2 As New ADODB.Recordset
rsPettPrDel2.Open "Delete from PettyPrint", constring, adOpenDynamic, adLockOptimistic

Dim SEro
SEro = Trim(frmPettyCash.LvwPCashUnConf.SelectedItem)

Dim RsNEwPettyForPrint As New ADODB.Recordset
RsNEwPettyForPrint.Open "Select * from pettyPrint ", constring, adOpenDynamic, adLockOptimistic

Dim rsRepPettCr As New ADODB.Recordset
Dim repPettCRSerNo, repPettCRAccName, repPettCRDate, repPettCRTotalAmt
rsRepPettCr.Open "Select distinct SerialNo,AccountNo,Accountname,datex,motham from Pettycashheld where  SerialNo = " & "'" & SEro & "'" & " and confMark is null and journalmark is null ", constring, adOpenDynamic, adLockOptimistic


While rsRepPettCr.EOF = False                      'This is for the Credit side
 repPettCRSerNo = rsRepPettCr!SerialNo
 repPettCRAccNo = rsRepPettCr!AccountNo
 repPettCRAccName = rsRepPettCr!accountname
 repPettCRDate = rsRepPettCr!Datex
 repPettCRTotalAmt = rsRepPettCr!motham
rsRepPettCr.MoveNext
Wend


'This is to get the Balance for the Credit Amount
Call prcgetdata
'---------------
  
  Dim rsRepPettyForPrintCr As New ADODB.Recordset
  rsRepPettyForPrintCr.Open "Select * from pettyprintcr where SerialNo = " & "'" & SEro & "'" & "", constring, adOpenDynamic, adLockOptimistic
   
   With rsRepPettyForPrintCr
   .AddNew
!SerialNo = repPettCRSerNo
!accountname = repPettCRAccName
!AccountCode = repPettCRAccNo
!Datex = repPettCRDate
!totalamt = repPettCRTotalAmt
!finalbalance = finalbalance
!AfterDeduc = Val(finalbalance) - Val(repPettCRTotalAmt)
'!TotUnposted = FirstA
 .Update
 finalbalance = 0
 End With
rsRepPettyForPrintCr.Close



     Dim rsRepPettDeb As New ADODB.Recordset
     rsRepPettDeb.Open "Select * from Pettycashheld where serialno = " & "'" & repPettCRSerNo & "'" & "", constring, adOpenDynamic, adLockOptimistic
     Dim repPettDbPaidto, repPettDbInv, repPettDBAccNo2, repPettDBAccName2, repPettDBPart, repPettDBExpl, repPettDBAmount, repPettCRAmount, repPettDbPod, repPettDbPOn
     
  Dim mNo
  mNo = 0
   While rsRepPettDeb.EOF = False
     repPettDbPaidto = rsRepPettDeb!PaidTo
     repPettDbInv = rsRepPettDeb!InvNo
     repPettDbPod = rsRepPettDeb!PODate
     repPettDbPOn = rsRepPettDeb!PoNumber
     
     repPettDBAccNo2 = rsRepPettDeb!AccountNo2
     repPettDBAccName2 = rsRepPettDeb!accountname2
     repPettDBPart = rsRepPettDeb!Partic
     repPettDBExpl = rsRepPettDeb!Explan
     repPettDBAmount = rsRepPettDeb!amount
     repPettCRAmount = rsRepPettDeb!creditamount
      repPettClass = rsRepPettDeb!Classification
        mNo = Val(mNo) + 1
        
            With RsNEwPettyForPrint
            .AddNew
                 !nox = mNo
                !SerialNo = repPettCRSerNo
                !PaidTo = repPettDbPaidto
'                !Datex = repPettCRDate
'                !Accountcode = repPettCRAccNo
'                !Accountname = repPettCRAccName
                !Accountcode2 = repPettDBAccNo2
                !accountname2 = repPettDBAccName2
                !Partic = repPettDBPart
                !Explan = repPettDBExpl
                !InvNo = repPettDbInv
                !PoNumber = repPettDbPOn
                !PODate = repPettDbPod
                 On Error Resume Next
                 
 '               !AmountCr = repPettCRTotalAmt
                !creditamount = repPettCRAmount
                !amount = repPettDBAmount
                !DebitLESScrAmt = Val(IIf(IsNull(repPettDBAmount), 0, (repPettDBAmount))) - Val(IIf(IsNull(repPettCRAmount), 0, (repPettCRAmount)))
                !Classification = repPettClass
                On Error GoTo 0
          .Update
          End With
        
                repPettDbPOn = ""
                repPettDbPod = ""
                repPettDbInv = ""
                repPettDbPaidto = ""
               ' repPettCRDate = ""
                'repPettCRAccNo = ""
                'repPettCRAccName = ""
                repPettDBAccNo2 = ""
                repPettDBAccName2 = ""
                repPettDBPart = ""
                repPettDBExpl = ""
                repPettCRTotalAmt = ""
                repPettDBAmount = ""
                creditamount = ""
                    
       rsRepPettDeb.MoveNext
       Wend
       rsRepPettDeb.Close
rsRepPettCr.Close

Dim AA2, AA1, FirstA, SecondB, ThirdC

'This is to Find out the Unposted but Confirmed Balances from PayJournal
Dim rsUnpostPayJor As New ADODB.Recordset
rsUnpostPayJor.Open "Select sum(CrAmount) as CrAmount,sum(DBamount) as DBAmount from PayJournal where Accno = " & "'" & repPettCRAccNo & "'" & " and status = 'Unposted' ", constring, adOpenDynamic, adLockOptimistic

With rsUnpostPayJor
AA1 = Val(IIf(IsNull(!DBamount), 0, (!DBamount))) - Val(IIf(IsNull(!CRamount), 0, (!CRamount)))
End With


'This is to Find out the Unposted but Confirmed Balances from PayJournal
Dim rsUnpostPettyJor As New ADODB.Recordset
rsUnpostPettyJor.Open "Select sum(cramount) as CrAmount,sum(DBamount) as DBAmount from PettyJournal where Accoutno = " & "'" & repPettCRAccNo & "'" & " and Postmark = 'Unposted'", constring, adOpenDynamic, adLockOptimistic

With rsUnpostPettyJor
AA2 = Val(IIf(IsNull(!DBamount), 0, (!DBamount))) - Val(IIf(IsNull(!CRamount), 0, (!CRamount)))
End With

FirstA = Val(AA1) + Val(AA2)

'-------------------------------------------


Dim PtyRepATotals As New ADODB.Recordset
PtyRepATotals.Open "select sum(DebitLESScrAmt) as DebitLESScrAmt,SUM(Amount) AS Amount, SUM(CreditAmount) AS CreditAmount from PettyPrint  where serialno  = " & "'" & repPettCRSerNo & "'" & "", constring, adOpenDynamic, adLockOptimistic

With PtyRepATotals
 'FirstA = PtyRepATotals!Amount
 ThirdC = PtyRepATotals!DebitLESScrAmt
End With

PtyRepATotals.Close



Dim ptyRepB As New ADODB.Recordset
ptyRepB.Open "select sum(Amount) as Amount from PettyCashHeld where AccountNo = " & "'" & repPettCRAccNo & "'" & " And serialno  <> " & "'" & repPettCRSerNo & "'" & " and confMark Is Null", constring, adOpenDynamic, adLockOptimistic

With ptyRepB
SecondB = IIf(IsNull(!amount), 0, (!amount))
End With
ptyRepB.Close



Dim ptyLast As New ADODB.Recordset
ptyLast.Open "select * from PettyPrintcr  where serialno  = " & "'" & repPettCRSerNo & "'" & "", constring, adOpenDynamic, adLockOptimistic

With ptyLast
 !BalaneLast = Val(Val(!finalbalance) + Val(FirstA)) - Val(IIf(IsNull(SecondB), 0, (SecondB)) + Val(IIf(IsNull(ThirdC), 0, (ThirdC))))
 !GLforPrintout = !finalbalance
 !finalbalance = Val(!finalbalance) + Val(FirstA)
 !PendingCredits = SecondB
 !TotUnposted = FirstA

.Update
End With



'While PtyRepTotals.EOF = False
' PtyRepTotals.MoveLast
' PtyRepTotals!DebitLESScrAmt = Val(PtyRepTotals!Amount) - Val(PtyRepTotals!CreditAmount)
'PtyRepTotals.MoveNext
'Wend

PettyReqNo = repPettCRSerNo
PettyDatex = repPettCRDate
ProcedPReqNo RepPettyReqest.Sections(2).Controls("lblRequestNo")
ProcedDatex RepPettyReqest.Sections(2).Controls("lblDatex")
procePrintBy RepPettyReqest.Sections(2).Controls("lblPrintedBy")

On Error Resume Next
DataEnvironment1.rsPettyPrint.Close
DataEnvironment1.rsPettyPrint.Requery
RepPettyReqest.Show 1
On Error GoTo 0


'-------------------------------------
'THIS IS FOR THE DERECT PRINT
uifo = MsgBox("Did you print it?", vbYesNo + vbQuestion, "Please confirm")
If uifo = vbYes Then
   
Dim rstconPr As New ADODB.Recordset
rstconPr.Open "Select * from pettycashheld where serialno = " & "'" & repPettCRSerNo & "'" & "", constring, adOpenDynamic, adLockOptimistic, adCmdText

  
rstconPr!printmark = "Printed"
rstconPr.Update
End If
'---------------------------------


'Dim 'xcls3 As New HabitatClass

'This is to call the Class for the Listview UN confirm
Dim rs55 As New ADODB.Recordset
''xcls3.ProcLvwPCashUnCon rs55


End Sub

Private Sub PCunPosJournal2()
'If frmPettyCash.LvwPCashUnConf.ListItems.Count = 0 Then
'MsgBox "List View is Empty", vbInformation
'Exit Sub
'End If
'Dim VarNOx

'This is to put the Journal mark in the pettycashheld table
Dim VarForSErNo
VarForSErNo = frmPettyCash.LvwPCashUnConf.SelectedItem
Dim rstMarkJrlPCash As New ADODB.Recordset
rstMarkJrlPCash.Open "Select * from pettycashheld where serialno = " & "'" & VarForSErNo & "'" & "", constring, adOpenDynamic, adLockOptimistic

While rstMarkJrlPCash.EOF = False
If rstMarkJrlPCash!journalmark = 1 Then
MsgBox "Sorry, This record is already journalised", vbInformation, "Message"
Exit Sub
End If

rstMarkJrlPCash!journalmark = 1
rstMarkJrlPCash.Update

rstMarkJrlPCash.MoveNext
Wend
rstMarkJrlPCash.Close
'This is to get the journal number for the new table and for the journalisation


Dim JOurnalNoforPetty As New ADODB.Recordset
JOurnalNoforPetty.Open "Setup", constring, adOpenKeyset, adLockPessimistic, adCmdTable


If Val(Left(JOurnalNoforPetty!CurrentMoYr, 2) <> Format(Date, "mm")) Then
   JOurnalNoforPetty!CurrentMoYr = Format(Date, "mm") & Trim(Right(Year(Date), 2))
   JOurnalNoforPetty!nextjn = "00001"
   JOurnalNoforPetty.Update
Else
   Jn = "PTC" & "-" & Right(Year(Date), 2) & Trim(Format(Date, "mm")) & "-" & JOurnalNoforPetty!nextjn
   nextjn = Val(JOurnalNoforPetty!nextjn)
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
   JOurnalNoforPetty!nextjn = Zeros & Trim(Val(nextjn) + 1)
   JOurnalNoforPetty.Update
   JOurnalNoforPetty.Close
End If
      frmPettyCash.LvwPCJournal.ListItems.Clear
      
      
      
      
                      
'This is to put the details in the Lsitview for the Journal

Dim rstPCjourn As New ADODB.Recordset
rstPCjourn.Open "SElect * from pettycashheld where serialno = " & "'" & VarForSErNo & "'" & " and journalmark = '1'", constring, adOpenDynamic, adLockOptimistic

VarNOx = 0
                'This part is to add the details of the Total Credit and it is only one time
Set MItem = frmPettyCash.LvwPCJournal.ListItems.Add(, , Jn) 'This is for the Credit balance
MItem.SubItems(1) = VarNOx
MItem.SubItems(2) = Trim(rstPCjourn!Datex)
'mitem.SubItems(3) = Trim(rstPCjourn!Partic)
'mitem.SubItems(4) = Trim(rstPCjourn!Explan)
MItem.SubItems(3) = "Total Credit"
MItem.SubItems(4) = "Total Credit "
MItem.SubItems(5) = Trim(rstPCjourn!AccountNo)
MItem.SubItems(6) = Trim(rstPCjourn!accountname)
MItem.SubItems(8) = IIf(IsNull(rstPCjourn!motham), 0, rstPCjourn!motham)
                 
                '--------------------------------------------------------------------------
                 



While rstPCjourn.EOF = False

VarNOx = Val(VarNOx) + 1

If (rstPCjourn!amount) <> "" Then
Set MItem = frmPettyCash.LvwPCJournal.ListItems.Add(, , Jn) 'This is for the Debit balance
MItem.SubItems(1) = VarNOx
MItem.SubItems(2) = Trim(rstPCjourn!Datex)
MItem.SubItems(3) = Trim(rstPCjourn!Partic)
MItem.SubItems(4) = Trim(rstPCjourn!Explan)
MItem.SubItems(5) = Trim(rstPCjourn!AccountNo2)
MItem.SubItems(6) = Trim(rstPCjourn!accountname2)
MItem.SubItems(7) = IIf(IsNull(rstPCjourn!amount), 0, (rstPCjourn!amount))

Else
Set MItem = frmPettyCash.LvwPCJournal.ListItems.Add(, , Jn) 'This is for the Debit balance
MItem.SubItems(1) = VarNOx
MItem.SubItems(2) = Trim(rstPCjourn!Datex)
'mitem.SubItems(3) = Trim(rstPCjourn!Partic)
'mitem.SubItems(4) = Trim(rstPCjourn!Explan)
MItem.SubItems(3) = "Other Credit"
MItem.SubItems(4) = "Other Credit "
MItem.SubItems(5) = Trim(rstPCjourn!AccountNo2)
MItem.SubItems(6) = Trim(rstPCjourn!accountname2)
MItem.SubItems(8) = IIf(IsNull(rstPCjourn!creditamount), 0, rstPCjourn!creditamount)

End If
'VarNOx = Val(VarNOx) + 1

'If Trim(rstPCjourn!Accountno) = "" Then
'Else
'mitem.SubItems(4) = Trim(rstPCjourn!Accountno)
'mitem.SubItems(5) = Trim(rstPCjourn!AccountName)
'mitem.SubItems(6) = Trim(rstPCjourn!Amount)
'End If

rstPCjourn.MoveNext
Wend
rstPCjourn.Close

'This is to get the details and put it into the new table (Journal)
Dim rsPettyJournal As New ADODB.Recordset

rsPettyJournal.Open "Select * from PettyJournal", constring, adOpenDynamic, adLockOptimistic

Dim i, JPCjourno, JPCDate, JPCpart, JPCexpl, JPCaccountno, JPCaccoutname, JPCDbamt, JPCCramt
i = 0
For i = 1 To frmPettyCash.LvwPCJournal.ListItems.Count
    Noms = frmPettyCash.LvwPCJournal.ListItems.Item(i).SubItems(1)
    JPCDate = frmPettyCash.LvwPCJournal.ListItems.Item(i).SubItems(2)
    JPCpart = frmPettyCash.LvwPCJournal.ListItems.Item(i).SubItems(3)
    JPCexpl = frmPettyCash.LvwPCJournal.ListItems.Item(i).SubItems(4)
    JPCaccountno = frmPettyCash.LvwPCJournal.ListItems.Item(i).SubItems(5)
    JPCaccoutname = frmPettyCash.LvwPCJournal.ListItems.Item(i).SubItems(6)
    JPCDbamt = frmPettyCash.LvwPCJournal.ListItems.Item(i).SubItems(7)
    JPCCramt = frmPettyCash.LvwPCJournal.ListItems.Item(i).SubItems(8)
    
    
    
                              'Here i have to call the proceduer to got the Father Category
                            acctNo = Trim(JPCaccountno)
                            Prevcap = Trim(Me.caption)
                            Call DisplayCats(Prevcap, acctNo, catName)
                            Categry = catName
  
    
    
    
With rsPettyJournal
.AddNew
!nox = Noms
!Journo = Jn
!SerialNo = VarForSErNo
!Datex = Format(JPCDate, "mm/dd/yyyy")
!confirmeddate = Format(Date, "mm/dd/yyyy")
!Partic = "Voucher # :" & VarForSErNo 'JPCpart
!Expla = JPCexpl & ">>>" & JPCpart
!AccoutNo = JPCaccountno
!accountname = JPCaccoutname
'On Error Resume Next
!DBamount = IIf(Trim(JPCDbamt) = "", 0, (JPCDbamt))
!CRamount = IIf(Trim(JPCCramt) = "", 0, (JPCCramt))
!category = Categry
!Description = VarForSErNo
!Prepby = cLogUser
'On Error GoTo 0
.Update
End With
Next

'MsgBox "Records Journalised Successfully", vbInformation, "Confirmation"

'refresh all the listviews
 'This is to call the Class for the Listview confirmed
'Dim 'xcls3 As New HabitatClass
Dim rs3 As New ADODB.Recordset
'xcls3.ProcLvwPCashConfirmed rs3

Dim rs4 As New ADODB.Recordset
'xcls3.ProcLvwPCashUnCon rs4

Dim rs5 As New ADODB.Recordset
'xcls3.ProcLvwPCJournal rs5

frmPettyCash.SSTab1.SetFocus
SendKeys "(Right)"
End Sub

Private Sub prcgetdata()

Dim constring As String
Dim myclass As New HabitatClass
Dim sqltable As Boolean
Dim xtable As String
Dim confin As New ADODB.Connection
Dim recfin As New ADODB.Recordset
Dim recglb As New ADODB.Recordset
Dim recope1 As New ADODB.Recordset
Dim recope As New ADODB.Recordset

sqltable = True
constring = "dsn=finance;uid=sa;pwd=;"
xtable = "Select * from FinanceMaster where active <> '0' and Accountcode = " & "'" & repPettCRAccNo & "'"
myclass.GetTables recfin, CON1, xtable, constring, sqltable

            Dim recglbdebitamount As Currency
            Dim recglbcreditamount As Currency
            Dim recopebeginningDebit As Currency
            Dim recopebeginningCredit As Currency
               
       'this is all for GLMaster table
        recglb.Open "SELECT AccountCode, SUM(CreditAmount) AS creditamount, SUM(DebitAmount) AS debitamount, CASE WHEN SUM(DebitAmount) >= SUM(creditamount) tHEN SUM(debitamount) - SUM(creditamount)else 0 END AS endingDebit, CASE WHEN SUM(creditamount) >= SUM(DebitAmount) THEN SUM(creditamount) - SUM(debitamount) Else 0 END AS endingCredit from GLMaster where accountcode = " & "'" & repPettCRAccNo & "'" & " GROUP BY AccountCode", constring, adOpenKeyset, adLockOptimistic
               
       'this is for the opening balance table
        recope.Open "select sum(beginningdebit) as beginningdebit,sum(beginningcredit) as beginningcredit from OpeningBalance where accountcode = " & "'" & repPettCRAccNo & "'", constring, adOpenKeyset, adLockOptimistic
        
'            !AccountCode = recfin!AccountCode
'            !AccountName = recfin!accountnameeng
'            !accountnameara = recfin!accountnamearab
            'take the begin balance from glmaster upto specific date
            'If recope1.BOF = False Then
            '    recopebeginningDebit11 = Val(IIf(IsNull(recope1!dEBITAmount), 0, recope1!dEBITAmount))
            '    recopebeginningCredit11 = Val(IIf(IsNull(recope1!creditamount), 0, recope1!creditamount))
            'Else
                recopebeginningDebit11 = 0
                recopebeginningCredit11 = 0
            'End If
            ' take the begin balance from opening balance table
            If recope.BOF = False Then
                recopebeginningDebit = Val(IIf(IsNull(recope!beginningdebit), 0, recope!beginningdebit)) + Val(recopebeginningDebit11)
                recopebeginningCredit = Val(IIf(IsNull(recope!beginningcredit), 0, recope!beginningcredit)) + Val(recopebeginningCredit11)
            Else
                recopebeginningDebit = 0 + Val(recopebeginningDebit11)
                recopebeginningCredit = 0 + Val(recopebeginningCredit11)
            End If
            'check for whether debit more or credit more
            If recopebeginningDebit >= recopebeginningCredit Then
                recopebeginningDebit = recopebeginningDebit - recopebeginningCredit
                recopebeginningCredit = 0
            Else
                recopebeginningCredit = recopebeginningCredit - recopebeginningDebit
                recopebeginningDebit = 0
            End If
                '!beginningdebit = recopebeginningDebit
                '!beginningcredit = recopebeginningCredit
            If recglb.BOF = False Then
                '!activitydebit = recglb!debitamount
                '!ActivityCredit = recglb!creditamount
                recglbdebitamount = IIf(IsNull(recglb!DebitAmount), 0, recglb!DebitAmount)
                recglbcreditamount = IIf(IsNull(recglb!creditamount), 0, recglb!creditamount)
            Else
                '!activitydebit = 0
                '!ActivityCredit = 0
                recglbdebitamount = 0
                recglbcreditamount = 0
            End If
            ed = Val(recglbdebitamount) + Val(recopebeginningDebit) ' Ending Debit
            ec = Val(recglbcreditamount) + Val(recopebeginningCredit) ' Ending Credit
'            If ed >= ec Then
'                ed = ed - ec
'                ec = 0
'            Else
'                ec = ec - ed
'                ed = 0
'            End If
               'final balance
               finalbalance = ed - ec
    recglb.Close
    recope.Close
End Sub

Private Sub PCunPosJournal_Click()

End Sub

Private Sub PetCEdit_Click()
If frmPettyCash.LVWPettyCash.ListItems.Count = 0 Then
MsgBox "List View is Empty ", vbInformation
Exit Sub
End If


If PetCEdit.caption = "Edit" Then
  Else
  '------------------------
  'THIS IS FOR UPDATING
  '------------------------
 Dim RSUpd As New ADODB.Recordset
 Dim varUpListTic, varUpListSEr
 
varUpListSEr = frmPettyCash.txtSerialNo.Text
'varUpListTic = frmPettyCash.txtFlag.Text
  RSUpd.Open "Select * from pettyCashHeld where ticketx = " & "'" & varUpListTic & "'" & " and serialno = " & "'" & varUpListSEr & "'" & "", conStr, adOpenDynamic, adLockOptimistic


End If

End Sub

Private Sub PettCancel_Click()

If frmPettyCash.LVWPettyCash.ListItems.Count = 0 Then
MsgBox "List is Empty", vbInformation
Exit Sub
End If

Mymsgx = MsgBox("Mr." & cLogUser & ",Do you want to Delete from the ListView", vbInformation + vbYesNo, "Please Confirm")
If Mymsgx = vbYes Then
Dim LVWPCIndex2, TotL
TotL = frmPettyCash.txtAmt

LVWPCIndex2 = frmPettyCash.LVWPettyCash.SelectedItem.Index
frmPettyCash.LVWPettyCash.ListItems.Remove (LVWPCIndex2)
frmPettyCash.txttotal.Text = Val(frmPettyCash.txttotal.Text) - Val(TotL)

End If
End Sub

Private Sub PettyPost_Click()
MsgBox "Please Goto the Main Menu, Click the PTC-Petty Cash , Right click the Date and Post it", vbInformation, "HELP  ->  Posting"

End Sub

Private Sub PettyPostCancel_Click()
If frmPettyCash.LvwPCJournal.ListItems.Count = 8 Then
MsgBox "List is Empty", vbInformation
Exit Sub
End If


Dim VarCanPCLstJ, varLovec
Dim CanPCLstJ As New ADODB.Recordset
varLoveccc = frmPettyCash.LvwPCJournal.SelectedItem



Mymsgx = MsgBox("Do you want to Cancell the Journals for " & varLoveccc & "", vbInformation + vbYesNo, "Please Confirm")
If Mymsgx = vbNo Then
Exit Sub
End If

CanPCLstJ.Open "SElect * from pettyjournal where Journo = " & "'" & varLoveccc & "'" & "", constring, adOpenDynamic, adLockOptimistic

While CanPCLstJ.EOF = False
CanPCLstJ!cancelledmark = 1
CanPCLstJ.MoveNext
Wend

Dim xclss As New HabitatClass

'This is to call the Class for the Listview UN confirm
Dim rsett As New ADODB.Recordset
'xclss.ProcLvwPCJournal rsett
 
 MsgBox "Records of " & varLoveccc & "  has been Cancelled Successfully", vbInformation, "Confirmation"

End Sub

Private Sub PettyUnPostReport_Click()

MsgBox "Please Goto the Main Menu, Click the PTC-Petty Cash , Right click the Date and Print the Report", vbInformation, "HELP"

'RepPettyUnposted.Show 1
End Sub

Private Sub pmj_Click()
If FrmPayableSetup.ListView8.ListItems.Count = 0 Then
MsgBox "No item to cancell", vbInformation
Exit Sub
End If
Dim rsCancelPayJ As New ADODB.Recordset
Dim rsCancelPayab As New ADODB.Recordset


laka = MsgBox("Mr. & cloguser & , You are Trying to Cancell the Journalised List,Are you sure to Cancell it", vbQuestion + vbYesNo, "Please Confirm it")
 If laka = vbNo Then
 Exit Sub
 End If


Dim CanVar, VarCanPayJ, VarCanPayab, RAms
CanVar = FrmPayableSetup.ListView8.SelectedItem

VarCanPayJ = "Select * from PayJournal where SerialNo =" & "'" & CanVar & "'" & ""
rsCancelPayJ.Open VarCanPayJ, constring, adOpenDynamic, adLockOptimistic

If rsCancelPayJ.EOF = False Then
rsCancelPayJ.MoveFirst
End If

CID = 0
While rsCancelPayJ.EOF = False
RAms = rsCancelPayJ!serno
rsCancelPayJ!cancelledmark = 1
rsCancelPayJ!CancelledBy = cLogUser
rsCancelPayJ.Update

rsCancelPayJ.MoveNext
CID = 1
Wend
rsCancelPayJ.Close

VarCanPayab = "Select * from PayableSetup where serialno = " & "'" & RAms & "'" & ""
rsCancelPayab.Open VarCanPayab, constring, adOpenDynamic, adLockOptimistic

If rsCancelPayab.EOF = False Then
rsCancelPayab.MoveFirst
End If


FBI = 0
While rsCancelPayab.EOF = False
rsCancelPayab!cancelledmark = 1
rsCancelPayab!CacelledUser = cLogUser
rsCancelPayab.Update
FBI = 1
rsCancelPayab.MoveNext
Wend


If CID = 1 And FBI = 1 Then
MsgBox "Cancellation Done Successfully", vbInformation, "Confirmation"
End If




'This is to refresh the listview 4
Dim Apple2
Apple2 = "Select * from PayableSetup where deletemark = '0'  and Post = 'No' and cancelledmark = '1'"
 Dim rstPaySetupCanc As New ADODB.Recordset
rstPaySetupCanc.Open Apple2, conStr, adOpenDynamic, adLockOptimistic

FrmPayableSetup.ListView4.ListItems.Clear

'This is  for ListView4
 rstPaySetupCanc.MoveFirst
  While rstPaySetupCanc.EOF = False
     Set MItem = FrmPayableSetup.ListView4.ListItems.Add(, , Format(rstPaySetupCanc!SerialNo))
     MItem.SubItems(1) = Format(rstPaySetupCanc!Xdate, "dd/mm/yyyy")
     MItem.SubItems(2) = Format(rstPaySetupCanc!Payee)
     MItem.SubItems(3) = Format(rstPaySetupCanc!DocNo)
     MItem.SubItems(4) = Format(rstPaySetupCanc!amtreqested, "#############.#0")
 On Error Resume Next
   Totlist4 = Val(Totlist4) + Val(Trim(rstPaySetupCanc!TotCrAmt)) 'This is for the Total of the List
     rstPaySetupCanc.MoveNext
     Wend
FrmPayableSetup.txtTotCanceleldBal = Trim(Totlist4)
On Error GoTo 0



'This is for the LIstview 8
Dim Xss
Dim rstJournalCan As New ADODB.Recordset

Xss = "SELECT * From PayJournal where status='unposted' order by serialno"
rstJournalCan.Open Xss, conStr, adOpenDynamic, adLockOptimistic, adCmdText

'this is to add Listview8
If rstJournalCan.EOF = False Then
 rstJournalCan.MoveFirst
End If
FrmPayableSetup.ListView8.ListItems.Clear
  While rstJournalCan.EOF = False
  If rstJournalCan!cancelledmark = "0" Then
     Set MItem = FrmPayableSetup.ListView8.ListItems.Add(, , Format(rstJournalCan!SerialNo))
     MItem.SubItems(1) = Format(rstJournalCan!ticket)
     MItem.SubItems(2) = Format(rstJournalCan!confirmeddate, "dd/mm/yyyy")
     MItem.SubItems(3) = Format(rstJournalCan!AccNo)
     MItem.SubItems(4) = Format(rstJournalCan!AccName)
     MItem.SubItems(5) = Format(rstJournalCan!DBamount, "#############.#0")
     MItem.SubItems(6) = Format(rstJournalCan!CRamount, "#############.#0")
       TotList8Db = Val(TotList8Db) + Val(Trim(rstJournalCan!DBamount)) 'This is for the Total of the List
       totlist8cr = Val(totlist8cr) + Val(Trim(rstJournalCan!CRamount)) 'This is for the Total of the List
 End If
     rstJournalCan.MoveNext
     Wend
FrmPayableSetup.txtJdb.Text = Trim(TotList8Db)
FrmPayableSetup.txtJCr.Text = Trim(totlist8cr)







End Sub

Private Sub PrintPoteds_Click()

MsgBox "Please Goto the Main Menu, Click the PYB-Payable setup , Right click the Date and Print the Report", vbInformation, "HELP"

'POSTEDpaysetupJNL.Show 1
End Sub

Private Sub PrtPayournalGroup_Click()
Dim GrandTotDB, GrandTotCR
'Dim VarList2
If FrmPayableSetup.ListView8.ListItems.Count = 0 Then
MsgBox "No items"
Exit Sub
End If

'
'Dim VarList2
'VarList2 = FrmPayableSetup.ListView8.SelectedItem
'
'On Error Resume Next
'DataEnvironment1.rsPayJournalGroup_Grouping.Close
'DataEnvironment1.PayJournalGroup_Grouping VarList2
'
'DataReport5.Show

Dim uSeriNo, uAccNo, uTN, uAccName, uDebit, uCredit, uClassification, uInvNo, uInvDt, uDDue

Dim VarList2
VarList2 = FrmPayableSetup.ListView8.SelectedItem


'Delete the Table
Dim DelPayUnpos As New ADODB.Recordset
DelPayUnpos.Open "Delete from printpayunposted", constring, adOpenDynamic, adLockOptimistic


Dim rspayMain As New ADODB.Recordset
rspayMain.Open "Select * from PayJournal where serialno = " & "'" & VarList2 & "'" & "", constring, adOpenDynamic, adLockOptimistic

While rspayMain.EOF = False
 uAccNo = rspayMain!AccNo
 uAccName = rspayMain!AccName
 uTN = rspayMain!ticket
 uClassification = rspayMain!Classification
 uDebit = rspayMain!DBamount
 uCredit = rspayMain!CRamount
 uSeriNo = rspayMain!serno
 VarForReportLabe2 = rspayMain!confirmeddate
GrandTotDB = Val(GrandTotDB) + Val(uDebit)
GrandTotCR = Val(GrandTotCR) + Val(uCredit)

Dim rsPrintPayUp As New ADODB.Recordset
rsPrintPayUp.Open "Select * from PrintPayUnposted", constring, adOpenDynamic, adLockOptimistic


  With rsPrintPayUp
  .AddNew
 !SerialNo = uSeriNo
!JOurnalNo = VarList2
!AccountNo = uAccNo
!accountname = uAccName
!ticket = uTN
!Classifcation = uClassification
!DebitAmount = uDebit
!creditamount = uCredit
'!InvoiceNo   =
'!InvoiceDate =
'!DueDate =
'!PoNumber    =
'!PODate  =
'!SENumber    =
'!SEDate  =
'!Amount  =
'!Explanations    =
.Update
End With
rsPrintPayUp.Close

VarForReportLabel = VarList2

rspayMain.MoveNext
Wend


'-------------    LINE     -----------------
Dim rsPrintPayUpLine0 As New ADODB.Recordset
rsPrintPayUpLine0.Open "Select * from PrintPayUnposted", constring, adOpenDynamic, adLockOptimistic


  With rsPrintPayUpLine0
  .AddNew
!SerialNo = uSeriNo
!xLabel8 = "Grand Total :"
!xLabel9 = Format(GrandTotDB, "###,###,###,###.#0")
!xLabel10 = Format(GrandTotCR, "###,###,###,###.#0")
.Update
End With
rsPrintPayUpLine0.Close






'-------------    LINE     -----------------
Dim rsPrintPayUpLine As New ADODB.Recordset
rsPrintPayUpLine.Open "Select * from PrintPayUnposted", constring, adOpenDynamic, adLockOptimistic


  With rsPrintPayUpLine
  .AddNew
!SerialNo = uSeriNo
!xLabel6 = "  nvoice No              Invoice Date               Date Due                 P.O No              P.O Date              S.E No              S.E Date                   Amount"
.Update
End With
rsPrintPayUpLine.Close





'-------------    INVOICE LABELS      -----------------

'Dim rsPrintPayUp4 As New ADODB.Recordset
'rsPrintPayUp4.Open "Select * from PrintPayUnposted", conString, adOpenDynamic, adLockOptimistic
'
'
'  With rsPrintPayUp4
'  .AddNew
'!serialno = uSeriNo
'!xLabel = "Invoice No"
'!xLabel1 = "Invoice Date"
'!xLabel2 = "Date Due"
'!xLabel3 = "P.O No"
'!xLabel4 = "P.O Date"
'!xLabel5 = "S.E No"
'!xLabel6 = "S.E No"
'!xLabel7 = "Amount"
''!DueDate =
''!PoNumber    =
''!PODate  =
''!SENumber    =
''!SEDate  =
''!Amount  =
''!Explanations    =
'.Update
'End With
'rsPrintPayUp4.Close
'
'
'-------------    INVOICE DETAILS      -----------------

Dim rspaySub As New ADODB.Recordset
rspaySub.Open "Select * from PayInvoiceDetails where serialno = " & "'" & uSeriNo & "'" & "", constring, adOpenDynamic, adLockOptimistic

While rspaySub.EOF = False
 uInvNo = rspaySub!InvNo
 uInvDate = rspaySub!InvDate
 uDDue = rspaySub!duedate
 uPOno = rspaySub!PoNumber
 uPODate = rspaySub!PODate
uSENo = rspaySub!SENumber
 uSEDate = rspaySub!SEDate
 uAmount = rspaySub!amount
Dim rsPrintPayUp2 As New ADODB.Recordset
rsPrintPayUp2.Open "Select * from PrintPayUnposted", constring, adOpenDynamic, adLockOptimistic



  With rsPrintPayUp2
  .AddNew
!SerialNo = uSeriNo
!invoiceno = uInvNo
!invoicedate = uInvDate
!duedate = uDDue
!PoNumber = uPOno
!PODate = uPODate
!SENumber = uSENo
!SEDate = uSEDate
!amount = uAmount
'!Explanations    =
.Update
End With
rsPrintPayUp2.Close

rspaySub.MoveNext
Wend


VarForReportLabe3 = uSeriNo


'-------------    LINE     -----------------

Dim rsPrintPayUpLine2 As New ADODB.Recordset
rsPrintPayUpLine2.Open "Select * from PrintPayUnposted", constring, adOpenDynamic, adLockOptimistic


  With rsPrintPayUpLine2
  .AddNew
!SerialNo = uSeriNo
!xLabel7 = "Explanation :"
.Update
End With
rsPrintPayUpLine2.Close





                  'EXPLANATION
Dim varExplanation
Dim PayExpl As New ADODB.Recordset
PayExpl.Open "Select * from payablesetup where serialno = " & "'" & uSeriNo & "'" & "", constring, adOpenDynamic, adLockOptimistic

varExplanation = PayExpl!Explanation



Dim rstPrintPayUp As New ADODB.Recordset
rstPrintPayUp.Open "Select * from PrintPayUnposted", constring, adOpenDynamic, adLockOptimistic


  With rstPrintPayUp
  .AddNew
!SerialNo = uSeriNo
!Explanations = varExplanation
.Update
End With
rstPrintPayUp.Close

  ReportLabel1 PrintPayUnposted.Sections(1).Controls("label14")
  ReportLabel2 PrintPayUnposted.Sections(1).Controls("label15")
  ReportLabel3 PrintPayUnposted.Sections(1).Controls("label16")



''
On Error Resume Next
DataEnvironment1.rsPrintPayUnpoted.Close

'
'RepPettyUnposted.Show
'On Error Resume Next
PrintPayUnposted.Show

End Sub
Private Sub ReportLabel1(mLabelx As RptLabel)
mLabelx.caption = VarForReportLabel
End Sub
Private Sub ReportLabel2(mLabely As RptLabel)
mLabely.caption = VarForReportLabe2
End Sub
Private Sub ReportLabel3(mLabelz As RptLabel)
mLabelz.caption = VarForReportLabe3
End Sub

Private Sub PtoLedger_Click()

MsgBox "Please Goto the Main Menu, Click the PYB-Payable Setup Journal, Right click the Date and Post it", vbInformation, "HELP"

'If FrmPayableSetup.ListView8.ListItems.Count = 0 Then
'MsgBox "No items To Select"
'Exit Sub
'End If
'
'        Dim rsPayAnal As New ADODB.Recordset
'        Dim ts
'        Dim mark
'        mark = "Unposted"
'        Dim mitem As ListItem
'        mess = MsgBox("Do you want to continue? ", vbQuestion + vbYesNo + vbDefaultButton2, "Please confirm")
'        If mess = vbYes Then
'       '      PostingJournal.Text1.Text = "SAL"
'
'             rsPayAnal.Open "SELECT  ConfirmedDate, COUNT(ConfirmedDate) AS  TotalTRan, SUM(DBAmount) AS DrAmt, SUM(CRAmount) AS CrAmt" _
'             & " From Payjournal where Status = " & "'" & mark & "'" & " GROUP BY ConfirmedDate ORDER BY ConfirmedDate", conString, adOpenKeyset, adLockPessimistic, adCmdText
'            'rsPayAnal.Close
'             Do Until rsPayAnal.EOF = True
'                Set mitem = Posting.ListView1.ListItems.Add(, , Format(rsPayAnal!confirmedDate, "dd/mm/yyyy"))
'                mitem.SubItems(1) = rsPayAnal!TotalTRan
'                mitem.SubItems(2) = FormatNumber(rsPayAnal!drAmt, 2, vbTrue, vbTrue, vbTrue)
'                mitem.SubItems(3) = FormatNumber(rsPayAnal!cramt, 2, vbTrue, vbTrue, vbTrue)
'                mitem.SubItems(4) = "Waiting"
'                rsPayAnal.MoveNext
'            Loop
'     ' Unload FrmPayableSetup
'
'         Posting.Show 1
'        End If

End Sub

Private Sub PurchDel_Click()

If frmPurchaseSetup.CmbPrepBy.Text = "" Then
MsgBox "You have to fill Combo 'Prepared By'"
frmPurchaseSetup.CmbPrepBy.SetFocus
Exit Sub
End If

frmPassword.txtBuffer.Text = "DeletePurchase"
frmPassword.caption = "Enter Password to Delete"
frmPassword.txtPrepBy = frmPurchaseSetup.CmbPrepBy.Text
frmPassword.txtUserId = frmPurchaseSetup.CmbPrepBy.Text

frmPassword.Show 1

End Sub

Private Sub PurchEdit_Click()
varitem = frmPurchaseSetup.ListView1.SelectedItem.Text


If frmPurchaseSetup.cmdedit.caption = Trim("E&dit") Then
     
     xvar = Trim(varitem)
     
If rstPurch.EOF = False Then
rstPurch.MoveFirst
End If
 
      While rstPurch.EOF = False
        If Trim(rstPurch!SerialNo) = Trim(xvar) Then
        
 frmPurchaseSetup.CmdNew.caption = "&Save"
 frmPurchaseSetup.cmdExit1.caption = "&Cancel"
 frmPurchaseSetup.txtAmountDue.Enabled = True
 'txtApprBy.Enabled = True
  frmPurchaseSetup.txtBranch.Enabled = True
  frmPurchaseSetup.txtDocuNo.Enabled = True
  frmPurchaseSetup.mskDateDue.Enabled = True
' txtNotedBy.Enabled = True
 'txtPrepBy.Enabled = True
  frmPurchaseSetup.txtRefNo.Enabled = True
  frmPurchaseSetup.CmbApprovedBy.Enabled = True
  frmPurchaseSetup.cmbCostCenter.Enabled = True
  frmPurchaseSetup.CmbNotedBy.Enabled = True
  frmPurchaseSetup.CmbPrepBy.Enabled = True
  frmPurchaseSetup.CmbSource.Enabled = True
 'Combo5.Enabled = True
 frmPurchaseSetup.cmbProfCenter.Enabled = True
' frmPurchaseSetup.txtPoNo.Enabled = True

 frmPurchaseSetup.mskDate.Enabled = True
 frmPurchaseSetup.List1.Enabled = True
 frmPurchaseSetup.txtStorEntryNo.Enabled = True
 frmPurchaseSetup.mskStoreEnDate.Enabled = True
 'frmPurchaseSetup.txtvenCode.Enabled = True

        frmPurchaseSetup.txtSerialNo.Text = rstPurch!SerialNo
       ' frmPurchaseSetup.CmbSource.Text = rstPurch!payto
        frmPurchaseSetup.txtDocuNo.Text = rstPurch!DocNo
        frmPurchaseSetup.txtBranch.Text = rstPurch!branch
        frmPurchaseSetup.List1.AddItem IIf(IsNull(rstPurch!Explanation), Nill, (rstPurch!Explanation))
        
        frmPurchaseSetup.cmbCostCenter.Text = rstPurch!costcenter
        frmPurchaseSetup.cmbProfCenter.Text = rstPurch!ProfCenter
        frmPurchaseSetup.mskDateDue.Text = rstPurch!DateDue
        frmPurchaseSetup.txtStorEntryNo.Text = rstPurch!StoreEntNo
        frmPurchaseSetup.mskStoreEnDate.Text = rstPurch!StoreEntDate
        frmPurchaseSetup.mskDate.Text = rstPurch!Xdate 'ERRROOr
        frmPurchaseSetup.txtAmtPaid.Text = rstPurch!AmtPaidBefore
        frmPurchaseSetup.txtAmtReq.Text = rstPurch!amtreqested
        frmPurchaseSetup.txtOutBal.Text = rstPurch!outbal
  '      frmPurchaseSetup.txtPoNo.Text = rstPaySetup!PONumber

        frmPurchaseSetup.txtProTax.Text = rstPurch!ProfitTax
        frmPurchaseSetup.txtPercentage.Text = rstPurch!Percentage
        frmPurchaseSetup.txtNoOfAtt.Text = rstPurch!NoOfAttech
        frmPurchaseSetup.txtAmountDue.Text = rstPurch!AmtDue
        frmPurchaseSetup.txtRefNo.Text = rstPurch!RefNo
        frmPurchaseSetup.cmdExit1.caption = "&Cancel"
        Me.Clear.Enabled = True

        End If
     rstPurch.MoveNext
        Wend
     
 frmPurchaseSetup.ListView2.ListItems.Clear
'this is to add Listview2

If RstItem.EOF = False Then
RstItem.MoveFirst
End If


  While RstItem.EOF = False
 If Trim(RstItem!SerialNo) = frmPurchaseSetup.txtSerialNo.Text Then
 Set MItem = frmPurchaseSetup.ListView2.ListItems.Add(, , Format(RstItem!itemcode))
     MItem.SubItems(1) = Format(RstItem!itemdesc)
     MItem.SubItems(2) = Format(RstItem!itemmodelno)
     MItem.SubItems(3) = Format(RstItem!itemdiamention)
     MItem.SubItems(4) = Format(RstItem!itemqty)
     MItem.SubItems(5) = Format(RstItem!itemprice)
     MItem.SubItems(6) = Format(RstItem!SurTax)
     MItem.SubItems(7) = Format(RstItem!vat)
     MItem.SubItems(8) = Format(RstItem!taxcredit)
     MItem.SubItems(9) = Format(RstItem!inventoryinn)


       TotList2 = Val(TotList2) + Val(Trim(RstItem!totalInventoryInn)) 'This is for the Total of the List
 End If
     RstItem.MoveNext
     Wend
'txtTotalInvInn.Text = Trim(TotList2)
frmPurchaseSetup.cmdExit1.caption = "&Cancel"
Me.PurchEdit.Enabled = False
frmPurchaseSetup.txtSerialNo.Enabled = False
End If

  frmPurchaseSetup.SSTab1.SetFocus
  SendKeys "{left}"

End Sub

Private Sub rAdd_Click()
FrmPaymentAnalysis.txtDBAccNo.SetFocus

End Sub

Private Sub RClear_Click()
Me.REdit.caption = "Edit" 'This will chang the Caption So we can Edit Again Inside the Lsit

End Sub

Private Sub Rdel_Click()
If FrmPaymentAnalysis.ListView6.ListItems.Count = 0 Then
MsgBox "List is Empty", vbInformation
Exit Sub
End If

Mymsgx = MsgBox("Mr." & cLogUser & ",Do you want to Delete from the ListView", vbInformation + vbYesNo, "Please Confirm")
If Mymsgx = vbNo Then
 Exit Sub
 End If

Dim LVW6, TotDeb6
TotDeb6 = FrmPaymentAnalysis.txtTotList6

LVW6 = FrmPaymentAnalysis.ListView6.SelectedItem.Index
FrmPaymentAnalysis.ListView6.ListItems.Remove (LVW6)

'If FrmPaymentAnalysis.txtTotList6 Is Not Null Then
FrmPaymentAnalysis.txtTotList6.Text = Val(FrmPaymentAnalysis.txtTotList6.Text) - Val(TotDeb6)
'End If

End Sub

Private Sub REdit_Click()
On Error Resume Next

If FrmPaymentAnalysis.ListView6.SelectedItem.Text = "" Then
MsgBox "listview is Empty or No item Selected", vbInformation, "Edit"
End If


'this is to EDIT XPayment(Inside List3)
If Me.REdit.caption = "Edit" Then
        FrmPaymentAnalysis.txtDBAccNo.Text = FrmPaymentAnalysis.ListView6.SelectedItem.Text
        FrmPaymentAnalysis.txtDBPartic.Text = FrmPaymentAnalysis.ListView6.SelectedItem.SubItems(1)
        FrmPaymentAnalysis.txtDBAmo.Text = FrmPaymentAnalysis.ListView6.SelectedItem.SubItems(2)
        Me.REdit.caption = "Update"
        Me.RClear.caption = "Cancel" 'This will Enable Internal Edit Again"

FrmPaymentAnalysis.txtTemp1List6.Text = FrmPaymentAnalysis.txtDBAccNo.Text
FrmPaymentAnalysis.txtTemp2List6.Text = FrmPaymentAnalysis.txtDBPartic.Text
 FrmPaymentAnalysis.Text4.Text = "Sukran"
   
 frmMenu.Rdel.Enabled = False
    
'This is  "Updating"
ElseIf Me.REdit.caption = "Update" And FrmPaymentAnalysis.txtIdentifyNewData.Text = "" Then

'This is to identify that the Updation is for List3 from frmPassword
frmPassword.txtBuffer.Text = "UpdateList6"
frmPassword.txtUserId.Text = FrmPaymentAnalysis.CmbPrepBy.Text

frmPassword.caption = "Enter Password to Update"
frmPassword.txtPrepBy = FrmPaymentAnalysis.CmbPrepBy.Text

frmPassword.Show 1

ElseIf Me.REdit.caption = "Update" And FrmPaymentAnalysis.txtIdentifyNewData.Text = "New" Then

FrmPaymentAnalysis.ListView6.SelectedItem = FrmPaymentAnalysis.txtDBAccNo.Text
FrmPaymentAnalysis.ListView6.SelectedItem.SubItems(1) = FrmPaymentAnalysis.txtDBPartic.Text
FrmPaymentAnalysis.ListView6.SelectedItem.SubItems(2) = FrmPaymentAnalysis.txtDBAmo.Text
'FrmPaymentAnalysis.ListView6.SelectedItem.ListSubItems.clear
Dim i
i = 0
va = FrmPaymentAnalysis.ListView6.ListItems.Count
 TotList6 = 0
For i = 1 To va
     
     rs = FrmPaymentAnalysis.ListView6.ListItems(i).SubItems(2)
     '.SelectedItem.SubItems(2) '= Trim(txtDBAmo.Text)
     TotList6 = Val(TotList6) + Val(Trim(rs))
Next
FrmPaymentAnalysis.txtTotList6.Text = TotList6


End If
End Sub

Private Sub sAdd_Click()
FrmPaymentAnalysis.txtAccNo.SetFocus
End Sub

Private Sub sclear_Click()
Me.sEdit.caption = "Edit" 'This will change the Caption So we can Edit Again Inside the Lsit


End Sub

Private Sub sDel_Click()
If FrmPaymentAnalysis.ListView3.ListItems.Count = 0 Then
MsgBox "List is Empty", vbInformation
Exit Sub
End If

Mymsgx = MsgBox("Mr." & cLogUser & ",Do you want to Delete from the ListView", vbInformation + vbYesNo, "Please Confirm")
If Mymsgx = vbNo Then
 Exit Sub
 End If

Dim LVW3, TotDeb3
TotDeb3 = FrmPaymentAnalysis.txtTotList3

LVW3 = FrmPaymentAnalysis.ListView3.SelectedItem.Index
FrmPaymentAnalysis.ListView3.ListItems.Remove (LVW3)

'If FrmPaymentAnalysis.txtTotList3 Is Not Null Then
FrmPaymentAnalysis.txtTotList3.Text = Val(FrmPaymentAnalysis.txtTotList3.Text) - Val(TotDeb3)
'End If
End Sub

Private Sub sEdit_Click()

On Error Resume Next
If FrmPaymentAnalysis.ListView3.SelectedItem.Text = "" Then
MsgBox "listview is Empty or No item Selected", vbInformation, "Edit"
End If



'this is to EDIT XPayment(Inside List3)
If Me.sEdit.caption = "Edit" Then
        FrmPaymentAnalysis.txtAccNo.Text = FrmPaymentAnalysis.ListView3.SelectedItem.Text
        FrmPaymentAnalysis.txtPartic.Text = FrmPaymentAnalysis.ListView3.SelectedItem.SubItems(1)
        FrmPaymentAnalysis.txtAmo.Text = FrmPaymentAnalysis.ListView3.SelectedItem.SubItems(2)
        Me.sEdit.caption = "Update"
        Me.sclear.caption = "Cancel" 'This will Enable Internal Edit Again"
        FrmPaymentAnalysis.Text4.Text = "Sthoothy"
Me.sDel.Enabled = False

FrmPaymentAnalysis.txtTemp1List3.Text = FrmPaymentAnalysis.txtAccNo.Text
FrmPaymentAnalysis.txtTemp2List3.Text = FrmPaymentAnalysis.txtPartic.Text
    '=-
    
'This is  "Updating"
ElseIf Me.sEdit.caption = "Update" And FrmPaymentAnalysis.txtIdentifyNewData.Text = "" Then

If FrmPaymentAnalysis.CmbPrepBy.Text = "" Then 'Chck the Combo is Empty
MsgBox "fill the Combo 'Prepared By' "

FrmPaymentAnalysis.CmbPrepBy.SetFocus
Exit Sub
End If

Me.sDel.Enabled = True


'This is to identify that the Updation is for List3 from frmPassword
frmPassword.txtBuffer.Text = "UpdateList3"
frmPassword.txtUserId.Text = FrmPaymentAnalysis.CmbPrepBy.Text
frmPassword.caption = "Enter Password to Update"
frmPassword.txtPrepBy = FrmPaymentAnalysis.CmbPrepBy.Text

frmPassword.Show 1



ElseIf Me.sEdit.caption = "Update" And FrmPaymentAnalysis.txtIdentifyNewData.Text = "New" Then

FrmPaymentAnalysis.ListView3.SelectedItem = FrmPaymentAnalysis.txtAccNo.Text
FrmPaymentAnalysis.ListView3.SelectedItem.SubItems(1) = FrmPaymentAnalysis.txtPartic.Text
FrmPaymentAnalysis.ListView3.SelectedItem.SubItems(2) = FrmPaymentAnalysis.txtAmo.Text
'FrmPaymentAnalysis.ListView6.SelectedItem.ListSubItems.clear
Dim i
i = 0
va = FrmPaymentAnalysis.ListView3.ListItems.Count
 TotList3 = 0
For i = 1 To va
     
     Le = FrmPaymentAnalysis.ListView3.ListItems(i).SubItems(2)
     '.SelectedItem.SubItems(2) '= Trim(txtDBAmo.Text)
     TotList3 = Val(TotList3) + Val(Trim(Le))
Next
FrmPaymentAnalysis.txtTotList3.Text = TotList3

End If
End Sub

Private Sub ShAdd_Click()
  FrmPayableSetup.SSTab1.SetFocus
  SendKeys "{Left}"
End Sub

Private Sub Shdel_Click()

If FrmPayableSetup.ListView1.ListItems.Count = 0 Then
MsgBox "No items To Select"
Exit Sub
End If



frmPassword.txtBuffer.Text = "Delete"
frmPassword.caption = "Enter Password to Delete"
frmPassword.Show 1
End Sub

Private Sub shedit_Click()

If FrmPayableSetup.ListView1.SelectedItem.SubItems(5) = "Printed" Then
MsgBox " You Cannot Edit ,Voucher is  already printed  ", vbInformation, "Editing Cancelled.."
Exit Sub

End If




Dim rstterm As ADODB.Recordset
Set rstpay = New ADODB.Recordset
Set rstterm = New ADODB.Recordset
Dim CON1 As ADODB.Connection
Set CON1 = New ADODB.Connection
conStr = "Provider=MSDASQL;DSN=Finance;UID=sa; PWD=;"

CON1.Open conStr

If FrmPayableSetup.ListView1.ListItems.Count = 0 Then
MsgBox "No items To Select"
Exit Sub
End If


    rstpay.Open "Select * from payablesetup", CON1, adOpenDynamic, adLockOptimistic
    rstterm.Open "Select * from Term", CON1, adOpenDynamic, adLockOptimistic

varitem = FrmPayableSetup.ListView1.SelectedItem.Text

  If FrmPayableSetup.cmdedit.caption = Trim("E&dit") Then
     
     xvar = Trim(varitem)
     
If rstpay.EOF = False Then
rstpay.MoveFirst
End If

     
      While rstpay.EOF = False
        If Trim(rstpay!SerialNo) = Trim(xvar) Then
        
         FrmPayableSetup.txtAmountDue.Enabled = True
 'txtApprBy.Enabled = True
 FrmPayableSetup.txtBranch.Enabled = True
 FrmPayableSetup.txtDocuNo.Enabled = True
 FrmPayableSetup.mskDateDue.Enabled = True
' txtNotedBy.Enabled = True
 'txtPrepBy.Enabled = True
 
 FrmPayableSetup.txtEnter1.Enabled = True
 FrmPayableSetup.TxtEnter2.Enabled = True
 FrmPayableSetup.txtRefNo.Enabled = True
 FrmPayableSetup.CmbApprovedBy.Enabled = True
 FrmPayableSetup.cmbCostCenter.Enabled = True
 FrmPayableSetup.CmbNotedBy.Enabled = True
 FrmPayableSetup.CmbPrepBy.Enabled = True
 FrmPayableSetup.CmbSource.Enabled = True
 
 'Combo5.Enabled = True
'FrmPayableSetup.mskPODate.Enabled = True
FrmPayableSetup.txtCreditNote.Enabled = True
FrmPayableSetup.txtDebitNote.Enabled = True
FrmPayableSetup.cmbPaymode.Enabled = True
FrmPayableSetup.cmbCurrency.Enabled = True
FrmPayableSetup.txtFCamount.Enabled = True
FrmPayableSetup.cmbPayee2.Enabled = True
FrmPayableSetup.txtTaxCredit.Enabled = True
FrmPayableSetup.cmbPayment.Enabled = True
 
 
 
'FrmPayableSetup.txtinvNo.Enabled = True
'FrmPayableSetup.mskInvDate.Enabled = True
'FrmPayableSetup.txtProTax.Enabled = True
'FrmPayableSetup.txtPercentage.Enabled = True
FrmPayableSetup.txtInvAmt.Enabled = True
FrmPayableSetup.txtCreditNote.Enabled = True
FrmPayableSetup.txtDebitNote.Enabled = True
FrmPayableSetup.txtOutBal.Enabled = True
FrmPayableSetup.txtAmtReq.Enabled = True
FrmPayableSetup.txtAmtPaid.Enabled = True

FrmPayableSetup.cmbProfCenter.Enabled = True
FrmPayableSetup.Option1.Enabled = True
FrmPayableSetup.Option2.Enabled = True

FrmPayableSetup.mskDate.Enabled = True
FrmPayableSetup.List1.Enabled = True
'FrmPayableSetup.txtStorEntryNo.Enabled = True
'FrmPayableSetup.mskStoreEnDate.Enabled = True
'FrmPayableSetup.txtvenCode.Enabled = True
FrmPayableSetup.cmbPaymentFor.Enabled = True
'FrmPayableSetup.txtPoNo.Enabled = True

On Error Resume Next
        FrmPayableSetup.txtSerialNo.Text = rstpay!SerialNo
        FrmPayableSetup.CmbSource.Text = rstpay!payto
        FrmPayableSetup.cmbPaymentFor.Text = rstpay!Payee
        FrmPayableSetup.txtDocuNo.Text = rstpay!DocNo
        FrmPayableSetup.txtBranch.Text = rstpay!branch
      '  FrmPayableSetup.List1.AddItem IIf(IsNull(rstpay!Explanation), Nill, (rstpay!Explanation))
        
         FrmPayableSetup.txtEnter1.Text = rstpay!Explanation
        ' FrmPayableSetup.mskPODate.Text = rstpay!PODate
        FrmPayableSetup.cmbCostCenter.Text = rstpay!costcenter
        FrmPayableSetup.cmbProfCenter.Text = rstpay!ProfCenter
        FrmPayableSetup.mskDateDue.Text = rstpay!DateDue
       ' FrmPayableSetup.txtStorEntryNo.Text = rstpay!StoreEntNo
       ' FrmPayableSetup.mskStoreEnDate.Text = rstpay!StoreEntDate
        FrmPayableSetup.mskDate.Text = rstpay!Xdate 'ERRROOrrr
'        FrmPayableSetup.txtinvNo.Text = rstpay!invoiceno
        FrmPayableSetup.txtInvAmt.Text = rstpay!invAmt
       ' FrmPayableSetup.mskInvDate.Text = rstpay!invoicedate
        FrmPayableSetup.txtTaxCredit.Text = rstpay!taxcredit


        FrmPayableSetup.txtNoOfAtt.Text = rstpay!NoOfAttech
        FrmPayableSetup.List1.Text = rstpay!Explanation

        FrmPayableSetup.txtAmtPaid.Text = rstpay!AmtPaidBefore
        FrmPayableSetup.txtAmtReq.Text = rstpay!amtreqested
        FrmPayableSetup.txtOutBal.Text = rstpay!outbal
        'FrmPayableSetup.txtPoNo.Text = IIf(IsNull(rstpay!PoNumber), Nill, (rstpay!PoNumber))
        'FrmPayableSetup.txtProTax.Text = rstpay!ProfitTax
       ' FrmPayableSetup.txtPercentage.Text = rstpay!Percentage
  
        FrmPayableSetup.TxtEnter2.Text = rstpay!NoOfAttech
        FrmPayableSetup.txtAmountDue.Text = rstpay!AmtDue
        FrmPayableSetup.txtRefNo.Text = rstpay!Requester
        FrmPayableSetup.cmbPaymentFor = rstpay!payfor
        FrmPayableSetup.cmdExit1.caption = "&Cancel"
        Me.Clear.Enabled = True

On Error Resume Next

FrmPayableSetup.txtCreditNote.Text = rstpay!creditnote
FrmPayableSetup.txtDebitNote.Text = rstpay!debitnote
FrmPayableSetup.cmbPaymode.Text = rstpay!paymode
FrmPayableSetup.cmbCurrency.Text = rstpay!ExCurrency
'FrmPayableSetup.cmbPayment.Text = rstpay!ExCurrency
FrmPayableSetup.txtFCamount.Text = rstpay!FcAmount
FrmPayableSetup.cmbPaymentLevel.Text = rstpay!PaymentLevels
'frmpayablesetup.txtTaxCredit.Text =
FrmPayableSetup.cmbPayee2.Text = rstpay!Payee
'FrmPayableSetup.Command2.Enabled = False



   On Error GoTo 0
        End If
     rstpay.MoveNext
        Wend
     
     
  'This is to Edit the Term Details
     
     
Dim rstTermX1 As New ADODB.Recordset
Dim VarstTermX1
VarstTermX1 = "Select * from Term where SerialNo = " & "'" & xvar & "'" & ""
 rstTermX1.Open VarstTermX1, CON1, adOpenDynamic, adLockOptimistic
 
If rstTermX1.EOF = False Then
rstTermX1.MoveFirst
End If

  
   While rstTermX1.EOF = False
        
     Set MItem = FrmPayableSetup.ListView2.ListItems.Add(, , Trim(rstTermX1!Rate))
     MItem.SubItems(1) = IIf(IsNull(rstTermX1!descr), "", (rstTermX1!descr))
     MItem.SubItems(2) = IIf(IsNull(rstTermX1!days), "", (rstTermX1!days))
     MItem.SubItems(3) = IIf(IsNull(rstTermX1!Mode), "", (rstTermX1!Mode))
     MItem.SubItems(4) = IIf(IsNull(rstTermX1!xlevel), "", (rstTermX1!xlevel))
     
     rstTermX1.MoveNext
     Wend
End If
rstTermX1.Close

FrmPayableSetup.txtTMDays.Visible = True
FrmPayableSetup.txtTMDes.Visible = True
FrmPayableSetup.txtTMlevel.Visible = True
FrmPayableSetup.txtTMmode.Visible = True
FrmPayableSetup.txtTmRate.Visible = True
'Me.sEdit.Enabled = False
FrmPayableSetup.cmdedit.caption = "&Update"
FrmPayableSetup.CmdNew.caption = "&Update"

FrmPayableSetup.cmdExit1.caption = "&Cancel"
Me.shedit.Enabled = False
FrmPayableSetup.txtSerialNo.Enabled = False
XPundai = FrmPayableSetup.txtAmtReq.Text

'Here i should call frmPayment Analysis
'FrmPaymentAnalysis.Show 1
FrmPaymentAnalysis.txtSerialNo = FrmPayableSetup.txtSerialNo
FrmPaymentAnalysis.MaskEdBox1 = Format(Date, "dd/mm/yyyy")
FrmPaymentAnalysis.txtAmtReq = XPundai
FrmPaymentAnalysis.Text1.Text = FrmPayableSetup.txtInvAmt.Text
FrmPaymentAnalysis.Text2.Text = FrmPayableSetup.txtAmtPaid.Text
FrmPaymentAnalysis.Text3.Text = FrmPayableSetup.txtOutBal.Text
FrmPaymentAnalysis.txtPaymentLevel.Text = FrmPayableSetup.cmbPaymentLevel.Text
 'This is to Edit the Xpayment and Add Details to List3 Details From ListView1(Main)
        
        
        
        
Dim rstList3X1 As New ADODB.Recordset
Dim VarrstList3X1
VarrstList3X1 = "Select * from xpayment where SerialNo = " & "'" & xvar & "'" & ""
 rstList3X1.Open VarrstList3X1, CON1, adOpenDynamic, adLockOptimistic
 
        On Error Resume Next
        Dim stu
        stu = 0
        If rstList3X1.EOF = False Then
        rstList3X1.MoveFirst
        End If
        While rstList3X1.EOF = False
       ' If Trim(rstList3X1!SerialNo) = Trim(xvar) Then
        Set MItem = FrmPaymentAnalysis.ListView3.ListItems.Add(, , Trim(rstList3X1!AccNo))
        MItem.SubItems(1) = IIf(IsNull(rstList3X1!AccName) = True, "", Trim(rstList3X1!AccName))
        MItem.SubItems(2) = Trim(rstList3X1!amount)
        stu = Val(stu) + Val(rstList3X1!amount)
       ' End If
        rstList3X1.MoveNext
         Wend
rstList3X1.Close




 'This is to Edit the XReceipt and Add Details to List3 Details From ListView1(Main)
 
 
 
Dim rstReceiptX1 As New ADODB.Recordset
Dim VarrstReceiptX1
VarrstReceiptX1 = "Select * from xReceipt where SerialNo = " & "'" & xvar & "'" & ""
rstReceiptX1.Open VarrstReceiptX1, CON1, adOpenDynamic, adLockOptimistic
 
 
        Dim ijk
        ijk = 0
        If rstReceiptX1.EOF = False Then
        rstReceiptX1.MoveFirst
        End If
        While rstReceiptX1.EOF = False
 '       If Trim(rstReceipt!SerialNo) = Trim(xvar) Then
        Set MItem = FrmPaymentAnalysis.ListView6.ListItems.Add(, , Trim(rstReceiptX1!AccNo))
        MItem.SubItems(1) = Trim(rstReceiptX1!AccName)
        MItem.SubItems(2) = Trim(rstReceiptX1!amount)
        ijk = Val(ijk) + Val(rstReceiptX1!amount)
    '    End If
        rstReceiptX1.MoveNext
         Wend
         
FrmPaymentAnalysis.txtTotList6 = ijk
FrmPaymentAnalysis.txtTotList3 = stu

FrmPaymentAnalysis.txtAccNo.Enabled = True
FrmPaymentAnalysis.txtPartic.Enabled = True
FrmPaymentAnalysis.txtAmo.Enabled = True

FrmPaymentAnalysis.txtDBAccNo.Enabled = True
FrmPaymentAnalysis.txtDBPartic.Enabled = True
FrmPaymentAnalysis.txtDBAmo.Enabled = True

FrmPayableSetup.txtTMDays.Visible = True
FrmPayableSetup.txtTMDes.Visible = True
FrmPayableSetup.txtTMlevel.Visible = True
FrmPayableSetup.txtTMmode.Visible = True
FrmPayableSetup.txtTmRate.Visible = True
'Me.sEdit.Enabled = False
FrmPayableSetup.CmdNew.caption = "&Update"
FrmPayableSetup.cmdedit.caption = "&Update"

FrmPaymentAnalysis.cmdedit.caption = "&Update"
FrmPaymentAnalysis.CmdNew.caption = "&Update"
FrmPaymentAnalysis.CmdNew.Visible = False
FrmPayableSetup.cmdExit1.caption = "&Cancel"
Me.shedit.Enabled = False
FrmPayableSetup.txtSerialNo.Enabled = False


  FrmPayableSetup.SSTab1.SetFocus
  SendKeys "{Left}"
   SendKeys "{Left}"
FrmPaymentAnalysis.Show 'dont put 1 if u put in the PayAnal form no Popup OK
End Sub

Private Sub Unpaidsss_Click()
RepUnpaids.Show 1
End Sub

Private Sub UPS_Click()

         On Error Resume Next
         ProcedPrepBy UnpaidByDateDue.Sections(2).Controls("lblPrepby")
         UnpaidByDateDue.Show 1
        ' Unload Me


'FrmUnpaidReport.Show 1
End Sub

Private Sub x_Click()
Unload FrmPayableSetup
Unload Me
End Sub

Private Sub xPPV_Click()
DRPpaidvoucher.Show 1
End Sub

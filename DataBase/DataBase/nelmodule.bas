Attribute VB_Name = "nelmodule"
'calling API function to auto dropdown the combo when it its got focus.
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Public FormNo As Integer
Public xCountry As String
Public xCountryNAMe As String
Public xBranchName As String
Public xBranchCode As String
Public xClientCode As String
Public TopLevelCode As String
Public TopLevelName As String
Public Level1Code As String
Public Level2Code As String
Public level3Code As String
Public level4Code As String
Public level5Code As String
Public Level1Name As String
Public Level2Name As String
Public level3Name As String
Public level4Name As String
'Public level5Name As String
Public cTotalItems As String
Public LogSucess As Boolean
Public WhatLevel As String
Public PrevItem1 As String
Public PrevItem2 As String
Public PrevItem3 As String
Public PrevItem4 As String
'Public PrevItem5 As String
Public cErr As Long
Public WhatColumnclick As Integer 'for editing the arab name and remark
Public rstUser As New ADODB.Recordset

'Setup for Machine hour user
Public CancelProcess As Boolean
Public MachineUsed As Integer
Public GenDesc As String
Public CancelGenDesc As Boolean
Public UserRole As String
Public cLogUser As String
Public GenJournalTotalTrn As Long
Public WhoProcess As String

Public FindAcctNAme As Boolean 'Gen,Inv,etc unable to load if user want to find account names

Public CancelAll As Boolean 'control for loading newaccts forms
Public Newform As NewAccts
Global Const constring = "dsN=Finance;UID=SA;PWD=;"
Global Const edittrn = False
Global Const LISTVIEW_MODE0 = "View Large Icons"
Global Const LISTVIEW_MODE1 = "View Small Icons"
Global Const LISTVIEW_MODE2 = "View List"
Global Const LISTVIEW_MODE3 = "View Details"
Sub PrintSalesJOurnal(SelectedDate As Date)
            Dim rstSj As New ADODB.Recordset
            Dim rstDlySumSJ As New ADODB.Recordset
            Dim rstJOurcode As New ADODB.Recordset
            Dim rstTOtTRansDly As New ADODB.Recordset
            Dim PrinterReady As Boolean
            Dim cRow As Integer
            
            
            rstJOurcode.Open "Select * from JOurnalCode order by Code", constring, adOpenKeyset, adLockPessimistic, adCmdText
            rstJOurcode.Move 10, 1
            xCode = rstJOurcode!JOurnalName
            Dim trandate As Date
            trandate = rstJOurcode!lastpostingdate

            rstDlySumSJ.Open "SELECT InvoiceDate,count(InvoiceDate) as TotREc, SUM(TradeRcvble) as TR, SUM(TradeDiscAmt)as TDA , SUM(MgtDiscAmt) as MDA , SUM(GrossSales) as GS, SUM(TranspoCharge)as TC, SUM(NetSales) as NS,  SUM(VAT)as VAT, SUM(SURTaxAmt)as STA From SalesJournal Where transdate=" & "'" & SelectedDate & "'" & "and remarks is null GROUP BY InvoiceDate", _
                                constring, adOpenKeyset, adLockPessimistic, adCmdText
            rstSj.Open "Select * from SalesJOurnal Where transdate=" & "'" & SelectedDate & "'" & "and remarks is null order by INvoiceNo", constring, adOpenKeyset, adLockPessimistic, adCmdText
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


'Printing the Daily Sales Summary
'If rstDlySumSJ.EOF = False Then
'    rstDlySumSJ.MoveLast
'    xto = rstDlySumSJ!transdate
'    rstDlySumSJ.MoveFirst
'End If
Dim InvDate As Date
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
              InvDate = Format(rstDlySumSJ!invoicedate, "dd/mm/yyyy")
              cInvDate = Format(SelectedDate, "dd/mm/yyyy")
              Printer.Print ; Tab(0); cInvDate & " Sales Totals" _
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
                  ; Tab(112 - LenGtTDA); Format(GtTDA, "###,###,###.#0") _
                  ; Tab(127 - LenGtMDA); Format(GtMDA, "###,###.#0") _
                  ; Tab(146 - LenGtGS); Format(GtGS, "###,###,###.#0") _
                  ; Tab(158 - LenGtTC); Format(GtTC, "###.#0") _
                  ; Tab(172 - LenGtNS); Format(GtNS, "###,###,###.#0") _
                  ; Tab(186 - LenGtVat); Format(GtVat, "###,###.#0") _
                  ; Tab(198 - LenGtSTA); Format(GtStA, "###,###.#0")
                  Printer.Print ; Tab(0); "=============================================================================================================================================================================="


Printer.Print ""
Printer.Print ; Tab(0); "Total Debit Amount           : " & Format(GtTR + GtMDA + GtTDA, "###,###,###.#0"); Tab(90); "___________                                              ____________                                              ___________                       "
Printer.Print ; Tab(0); "Total Credit Amount          : " & Format(GtTR + GtMDA + GtTDA, "###,###,###.#0"); Tab(90); "Prepared by                                                   Checked by                                                  Approved by                     "
Printer.Print ; Tab(0); "=============================================================================================================================================================================="
Printer.Print ; Tab(103); "***End of the Report***"
Printer.Print ""
Printer.Print ""
Printer.Print ""



'printing by Profit Center
Dim rsByProfitCenter As New ADODB.Recordset
rsByProfitCenter.Open "SELECT  ProfitCenter, COUNT(ProfitCenter) AS TotalbyProfitCenter, SUM(TradeRcvble) AS TR, SUM(TradeDiscAmt) AS TDA,SUM(MgtDiscAmt) AS MDA," _
             & " SUM(GrossSales) AS GS, SUM(TranspoCharge) AS TC,SUM(NetSales) AS NS,SUM(VAt) AS VAT,SUM(SurTaxAmt) AS STA" _
             & " From SalesJournal where transdate=" & "'" & SelectedDate & "'" & "and Remarks is null GROUP BY ProfitCenter ORDER BY ProfitCenter ", constring, adOpenKeyset, adLockPessimistic, adCmdText


Printer.FontBold = True
Printer.Print ; Tab(0); "DAILY SALES by Profit Center " & cInvDate
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
                  ; Tab(112 - LenGtTDA); Format(GtTDA, "###,###,###.#0") _
                  ; Tab(127 - LenGtMDA); Format(GtMDA, "###,###.#0") _
                  ; Tab(146 - LenGtGS); Format(GtGS, "###,###,###.#0") _
                  ; Tab(158 - LenGtTC); Format(GtTC, "###.#0") _
                  ; Tab(173 - LenGtNS); Format(GtNS, "###,###,###.#0") _
                  ; Tab(185 - LenGtVat); Format(GtVat, "###,###.#0") _
                  ; Tab(197 - LenGtSTA); Format(GtStA, "###,###.#0")
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
    Printer.Print ; Tab(0); "Sales Journal Reports(Unposted)"
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
Sub DisplayCats(Prevcap As String, acctNo As String, catName As String)
Dim recfindaccount As New ADODB.Recordset
recfindaccount.Open "select * from level6 where accountcode = " & "'" & acctNo & "'", constring, adOpenKeyset, adLockOptimistic
If recfindaccount.BOF = True Then
   recfindaccount.Close
   recfindaccount.Open "select * from level5 where accountcode = " & "'" & acctNo & "'", constring, adOpenKeyset, adLockOptimistic
    If recfindaccount.BOF = True Then
      recfindaccount.Close
        recfindaccount.Open "select * from level4 where accountcode = " & "'" & acctNo & "'", constring, adOpenKeyset, adLockOptimistic
        If recfindaccount.BOF = True Then
             recfindaccount.Close
             recfindaccount.Open "select * from level3 where accountcode = " & "'" & acctNo & "'", constring, adOpenKeyset, adLockOptimistic
             If recfindaccount.BOF = False Then
               catName = recfindaccount!Level1Name & Chr(187) & recfindaccount!Level2Name
               recfindaccount.Close
            End If
        Else
        catName = recfindaccount!Level1Name & Chr(187) & recfindaccount!Level2Name & Chr(187) & recfindaccount!level3Name
        recfindaccount.Close
        End If
      Else
        catName = recfindaccount!Level2Name & Chr(187) & recfindaccount!level3Name & Chr(187) & recfindaccount!level4Name
        recfindaccount.Close
     End If
    Else
        catName = recfindaccount!level3Name & Chr(187) & recfindaccount!level4Name & Chr(187) & recfindaccount!level5Name
        recfindaccount.Close
   End If

End Sub
Sub DisplayCatsName(Prevcap As String, acctNo As String, catName As String)
Dim recfindaccount As New ADODB.Recordset
recfindaccount.Open "select * from level6 where Ltrim(accountNAmeEng) = " & "'" & acctNo & "'", constring, adOpenKeyset, adLockOptimistic
If recfindaccount.BOF = True Then
   recfindaccount.Close
   recfindaccount.Open "select * from level5 where Ltrim(accountNAmeEng) = " & "'" & acctNo & "'", constring, adOpenKeyset, adLockOptimistic
    If recfindaccount.BOF = True Then
      recfindaccount.Close
        recfindaccount.Open "select * from level4 where Ltrim(accountNAmeEng)= " & "'" & acctNo & "'", constring, adOpenKeyset, adLockOptimistic
        If recfindaccount.BOF = True Then
             recfindaccount.Close
             recfindaccount.Open "select * from level3 where Ltrim(accountNAmeEng)= " & "'" & acctNo & "'", constring, adOpenKeyset, adLockOptimistic
             If recfindaccount.BOF = False Then
                catName = recfindaccount!Level1Name & Chr(187) & recfindaccount!Level2Name
                 recfindaccount.Close
             End If
            Else
               catName = recfindaccount!Level1Name & Chr(187) & recfindaccount!Level2Name & Chr(187) & recfindaccount!level3Name
               recfindaccount.Close
        End If
      Else
        catName = recfindaccount!Level2Name & Chr(187) & recfindaccount!level3Name & Chr(187) & recfindaccount!level4Name
        recfindaccount.Close
      End If
     Else
       catName = recfindaccount!level3Name & Chr(187) & recfindaccount!level4Name & Chr(187) & recfindaccount!level5Name
       recfindaccount.Close
     End If
End Sub
Sub EnableMenu(cUser As String, cROle As String)
If UCase(cROle) = UCase("General") Then
    'under Transaction menu
    Mainform.xGEnJOurn.Enabled = True
    Mainform.xFixedASset.Enabled = False
    Mainform.xSAles.Enabled = False
    Mainform.xInventory.Enabled = False
    Mainform.xPaySetup.Enabled = False
    Mainform.xPurchaseJOurn.Enabled = False
    Mainform.xPurchaseSEtup.Enabled = False
    Mainform.xcAShRct.Enabled = True
    Mainform.xCashPmt.Enabled = True

    'Mainform.NewAcct.Enabled = False
    Mainform.RegisternewAsset.Enabled = False
    
    'Under File Menu
    Mainform.BankAccount.Enabled = False
    Mainform.Payee.Enabled = False
    Mainform.taxdetails.Enabled = False
    Mainform.PMTCat.Enabled = False
    Mainform.assigningcheque.Enabled = False
    
    'under Transaction button
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(1).Enabled = True
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(2).Enabled = False
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(3).Enabled = False
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(4).Enabled = False
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(5).Enabled = False
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(6).Enabled = False
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(7).Enabled = False
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(8).Enabled = False
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(9).Enabled = False
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(10).Enabled = False
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(11).Enabled = False
     'under Post button
    Mainform.xPost.Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(1).Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(2).Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(3).Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(4).Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(5).Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(6).Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(7).Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(8).Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(9).Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(10).Enabled = False
    
    
ElseIf UCase(cROle) = UCase("Asset") Then
    'under Transaction menu
    Mainform.xGEnJOurn.Enabled = False
    Mainform.xFixedASset.Enabled = True
    Mainform.xSAles.Enabled = False
    Mainform.xInventory.Enabled = False
    Mainform.xPaySetup.Enabled = False
    Mainform.xPurchaseJOurn.Enabled = False
    Mainform.xPurchaseSEtup.Enabled = False
    Mainform.xcAShRct.Enabled = False
    Mainform.xCashPmt.Enabled = False
    'Mainform.NewAcct.Enabled = False
    Mainform.RegisternewAsset.Enabled = True
    
    'Under File Menu
    Mainform.BankAccount.Enabled = False
    Mainform.Payee.Enabled = False
    Mainform.taxdetails.Enabled = False
    Mainform.PMTCat.Enabled = False
    Mainform.assigningcheque.Enabled = False
    
    'under Transaction button
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(1).Enabled = False
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(2).Enabled = True
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(3).Enabled = False
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(4).Enabled = False
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(5).Enabled = False
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(6).Enabled = False
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(7).Enabled = False
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(8).Enabled = False
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(9).Enabled = False
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(10).Enabled = False
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(11).Enabled = False
     'under Post button
    Mainform.xPost.Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(1).Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(2).Enabled = True
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(3).Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(4).Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(5).Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(6).Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(7).Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(8).Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(9).Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(10).Enabled = False
    
    
ElseIf UCase(cROle) = UCase("Sales") Then
    'under Transaction menu
    Mainform.xGEnJOurn.Enabled = False
    Mainform.xFixedASset.Enabled = False
    Mainform.xSAles.Enabled = True
    Mainform.xInventory.Enabled = False
    Mainform.xPaySetup.Enabled = False
    Mainform.xPurchaseJOurn.Enabled = False
    Mainform.xPurchaseSEtup.Enabled = False
    Mainform.xcAShRct.Enabled = True
    Mainform.xCashPmt.Enabled = True
    'Mainform.NewAcct.Enabled = False
    Mainform.RegisternewAsset.Enabled = False
    
    'Under File Menu
    Mainform.BankAccount.Enabled = False
    Mainform.Payee.Enabled = False
    Mainform.taxdetails.Enabled = False
    Mainform.PMTCat.Enabled = False
    Mainform.assigningcheque.Enabled = False
    
    'under Transaction button
   
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(1).Enabled = False
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(2).Enabled = False
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(3).Enabled = True
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(4).Enabled = False
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(5).Enabled = False
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(6).Enabled = False
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(7).Enabled = False
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(8).Enabled = False
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(9).Enabled = False
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(10).Enabled = False
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(11).Enabled = False
     'under Post button
    Mainform.xPost.Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(1).Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(2).Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(3).Enabled = True
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(4).Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(5).Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(6).Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(7).Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(8).Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(9).Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(10).Enabled = False
    
    
ElseIf UCase(cROle) = UCase("Inventory") Then
    'under Transaction menu
    Mainform.xGEnJOurn.Enabled = False
    Mainform.xFixedASset.Enabled = False
    Mainform.xSAles.Enabled = False
    Mainform.xInventory.Enabled = True
    Mainform.xPaySetup.Enabled = False
    Mainform.xPurchaseJOurn.Enabled = False
    Mainform.xPurchaseSEtup.Enabled = False
    Mainform.xcAShRct.Enabled = False
    Mainform.xCashPmt.Enabled = False
    'Mainform.NewAcct.Enabled = False
    Mainform.RegisternewAsset.Enabled = False
    
    'Under File Menu
    Mainform.BankAccount.Enabled = False
    Mainform.Payee.Enabled = False
    Mainform.taxdetails.Enabled = False
    Mainform.PMTCat.Enabled = False
    Mainform.assigningcheque.Enabled = False
    
    'under Transaction button
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(1).Enabled = False
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(2).Enabled = False
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(3).Enabled = False
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(4).Enabled = True
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(5).Enabled = False
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(6).Enabled = False
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(7).Enabled = False
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(8).Enabled = False
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(9).Enabled = False
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(10).Enabled = False
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(11).Enabled = False
     'under Post button
    Mainform.xPost.Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(1).Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(2).Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(3).Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(4).Enabled = True
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(5).Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(6).Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(7).Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(8).Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(9).Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(10).Enabled = False
    
   
ElseIf UCase(cROle) = UCase("Cashier") Then ' this is for anu
    Mainform.Timer2.Interval = 0
    'under Transaction menu
    Mainform.xGEnJOurn.Enabled = False
    Mainform.xFixedASset.Enabled = False
    Mainform.xSAles.Enabled = False
    Mainform.xInventory.Enabled = False
    Mainform.xPaySetup.Enabled = False
    Mainform.xPurchaseJOurn.Enabled = False
    Mainform.xPurchaseSEtup.Enabled = False
    Mainform.xcAShRct.Enabled = True
    Mainform.xCashPmt.Enabled = True
    'Mainform.NewAcct.Enabled = False
    Mainform.RegisternewAsset.Enabled = False
    Mainform.CashPosition.Enabled = True
    Mainform.creditnote.Enabled = False
    Mainform.debitnote.Enabled = False
    
    'Under File Menu
    Mainform.BankAccount.Enabled = False
    Mainform.Payee.Enabled = False
    Mainform.taxdetails.Enabled = False
    Mainform.PMTCat.Enabled = False
    Mainform.assigningcheque.Enabled = False
    
    'under Transaction button
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(1).Enabled = False
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(2).Enabled = False
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(3).Enabled = False
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(4).Enabled = False
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(5).Enabled = False
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(6).Enabled = False
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(7).Enabled = False
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(8).Enabled = False
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(9).Enabled = True
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(10).Enabled = True
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(11).Enabled = True
     'under Post button
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(1).Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(2).Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(3).Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(4).Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(5).Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(6).Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(7).Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(8).Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(9).Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(10).Enabled = False

ElseIf UCase(cROle) = UCase("Payables") Then
    'under Transaction menu
    Mainform.xGEnJOurn.Enabled = False
    Mainform.xFixedASset.Enabled = False
    Mainform.xSAles.Enabled = False
    Mainform.xInventory.Enabled = False
    Mainform.xPaySetup.Enabled = True
    Mainform.xPurchaseJOurn.Enabled = False
    Mainform.xPurchaseSEtup.Enabled = False
    Mainform.xcAShRct.Enabled = True
    Mainform.xCashPmt.Enabled = True
    'Mainform.NewAcct.Enabled = False
    Mainform.RegisternewAsset.Enabled = False
    
    'Under File Menu
    Mainform.BankAccount.Enabled = False
    Mainform.Payee.Enabled = False
    Mainform.taxdetails.Enabled = False
    Mainform.PMTCat.Enabled = False
    Mainform.assigningcheque.Enabled = True
    
    'under Transaction button
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(1).Enabled = False
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(2).Enabled = False
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(3).Enabled = False
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(4).Enabled = False
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(5).Enabled = False
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(6).Enabled = True
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(7).Enabled = True
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(8).Enabled = True
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(9).Enabled = False
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(10).Enabled = True
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(11).Enabled = True
   
     'under Post button
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(1).Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(2).Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(3).Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(4).Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(5).Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(6).Enabled = True
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(7).Enabled = True
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(8).Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(9).Enabled = True
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(10).Enabled = True

ElseIf UCase(cROle) = UCase("Accounts1") Or UCase(cROle) = UCase("Accounts2") Or UCase(cROle) = UCase("Admin") Then
    'under Transaction menu
    Mainform.xAcctGrouping.Enabled = True
    Mainform.xGEnJOurn.Enabled = True
    Mainform.xFixedASset.Enabled = True
    Mainform.xSAles.Enabled = True
    Mainform.xInventory.Enabled = True
    Mainform.xPaySetup.Enabled = True
    Mainform.xPurchaseJOurn.Enabled = True
    Mainform.xPurchaseSEtup.Enabled = True
    Mainform.xcAShRct.Enabled = True
    Mainform.xCashPmt.Enabled = True
    'Mainform.NewAcct.Enabled = True
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
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(11).Enabled = True
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

ElseIf UCase(cROle) = UCase("Auditor") Then
    'under Transaction menu
    Mainform.xAcctGrouping.Enabled = False
    Mainform.xGEnJOurn.Enabled = False
    Mainform.xFixedASset.Enabled = False
    Mainform.xSAles.Enabled = False
    Mainform.xInventory.Enabled = False
    Mainform.xPaySetup.Enabled = False
    Mainform.xPurchaseJOurn.Enabled = False
    Mainform.xPurchaseSEtup.Enabled = False
    Mainform.xcAShRct.Enabled = False
    Mainform.xCashPmt.Enabled = False
    'Mainform.NewAcct.Enabled = False
    Mainform.RegisternewAsset.Enabled = False
    Mainform.BankTrans.Enabled = False
    Mainform.creditnote.Enabled = False
    Mainform.debitnote.Enabled = False
    
    'Under File Menu
    Mainform.BankAccount.Enabled = False
    Mainform.Payee.Enabled = False
    Mainform.taxdetails.Enabled = False
    Mainform.PMTCat.Enabled = False
    Mainform.assigningcheque.Enabled = False
    Mainform.xcAShRct.Enabled = False
    Mainform.xCashPmt.Enabled = False
    'under Transaction button
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(1).Enabled = False
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(2).Enabled = False
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(3).Enabled = False
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(4).Enabled = False
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(5).Enabled = False
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(6).Enabled = False
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(7).Enabled = False
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(8).Enabled = False
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(9).Enabled = False
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(10).Enabled = False
    Mainform.Toolbar1.Buttons(4).ButtonMenus.Item(11).Enabled = False
     'under Post button
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(1).Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(2).Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(3).Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(4).Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(5).Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(6).Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(7).Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(8).Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(9).Enabled = False
    Mainform.Toolbar1.Buttons(5).ButtonMenus.Item(10).Enabled = False

End If
End Sub

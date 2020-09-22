VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HabitatClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mvarKupal As String 'local copy
Public Property Let Kupal(ByVal vData As String)
    mvarKupal = vData
End Property
Public Property Get Kupal() As String
    Kupal = mvarKupal
End Property
Public Sub ArabicAmountInWords(AmtInFigure As Currency, AmtInwords As Variant, mCurrency As Variant, mCents As Variant)
On Error GoTo MyMsg
Dim wordit As String
Dim Amt As Currency
AmtInFigure = Format(AmtInFigure, "###########.#0")
If AmtInFigure > 999999999.99 Then
    xmsg = MsgBox("Amount too big", vbOKOnly + vbExclamation)
    Exit Sub
End If
On Error Resume Next
cLen = InStr(1, AmtInFigure, ".")
'Manipulating the decimal and Whole the number
If cLen = 0 Then
 Amt = AmtInFigure
 mDec = "00"
 Else
 Amt = Left(AmtInFigure, cLen - 1)
 mDec = (Mid(AmtInFigure, cLen + 1, 2))
End If

If Len(mDec) = 1 Then
    mDec = mDec & "0"
End If
If cLen = 0 Then
  figure = AmtInFigure
 Else
  figure = Left(AmtInFigure, cLen - 1)
End If
Dim Unitword() As String
Dim mHundred As String
Dim nUnitword() As String
Dim UnitwordHun() As String
Dim Tenword() As String
ReDim nUnitword(20)
ReDim UnitwordHun(20)
ReDim Unitword(20)
ReDim Tenword(9)

Unitword(1) = " æÇÍÏ "
Unitword(2) = " ÇËäÇä"
Unitword(3) = " ËáÇËÉ "
Unitword(4) = " ÇÑÈÚÉ "
Unitword(5) = " ÎãÓÉ "
Unitword(6) = " ÓÊÉ "
Unitword(7) = " ÓÈÚÉ "
Unitword(8) = " ËãÇäíÉ "
Unitword(9) = " ÊÓÚÉ "
Unitword(10) = " ÚÔÑÉ "
Unitword(11) = " ÇÍÏ ÚÔÑ "
Unitword(12) = " ÇËäÇ ÚÔÑ "
Unitword(13) = " ËáÇËÉ ÚÔÑ "
Unitword(14) = " ÇÑÈÚÉ ÚÔÑ "
Unitword(15) = " ÎãÓÉ ÚÔÑ "
Unitword(16) = " ÓÊÉ ÚÔÑ "
Unitword(17) = " ÓÈÚÉ ÚÔÑ "
Unitword(18) = " ËãÇäíÉ ÚÔÑ "
Unitword(19) = " ÊÓÚÉ ÚÔÑ "
Tenword(1) = "ÚÔÑÉ"
Tenword(2) = "ÚÔÑæä"
Tenword(3) = "ËáÇËæä"
Tenword(4) = "ÇÑÈÚæä"
Tenword(5) = "ÎãÓæä"
Tenword(6) = "ÓÊæä"
Tenword(7) = "ÓÈÚæä"
Tenword(8) = "ËãÇäæä"
Tenword(9) = "ÊÓÚæä"
mHundred = "ãÇÆÉ"
mthousand = " ÇáÝ "
mMillion = " ãáíæä "
mOnly = " ÝÞØ áÇÛíÑ "
wordit = ""
    
         
         
         If figure >= 1000000 And figure < 1000000000 Then
                a = Int(figure / 1000000)
               'for hundred million
                   If Len(Trim(a)) = 3 And Left(a, 1) = 1 And Val(Right(a, 2)) = 0 Then ' for hundred million
                       wordit = wordit & mHundred & " " & mMillion
                      
                      ElseIf Len(Trim(a)) = 3 And Val(Right(a, 2)) > 0 And Left(a, 1) < 3 Then 'And Val(Mid(a, 2, 1)) <> 0 And Val(Right(a, 1)) <> 0 Then
                        
                        If Len(Trim(a)) = 3 And Val(Right(a, 2)) < 21 Then
                            lastmh = Val(Right(a, 2))
                            wah = " æ "
                            If Val(Right(a, 1)) = 0 Then
                             wah = " "
                              lastmh = Val(Mid(a, 2, 1))
                              wordit = wordit & "ãÇÆÉ" & wah & Tenword(lastmh) & mMillion
                             Else
                              wordit = wordit & "ãÇÆÉ" & wah & Unitword(lastmh) & mMillion
                            End If
                          Else
                            Lastdigit = Val(Right(a, 1))
                            Middigit = Val(Mid(a, 2, 1))
                            wah = " æ "
                            If Val(Left(a, 1)) = 1 Then
                               wordit = wordit & "ãÇÆÉ" & wah & Unitword(Lastdigit) & wah & Tenword(Middigit) & mMillion
                             ElseIf Val(Left(a, 1)) = 2 Then
                               wordit = wordit & "ãÇÆÊÇä" & wah & Unitword(Lastdigit) & wah & Tenword(Middigit) & mMillion
                            End If
                         End If
                      ElseIf Len(Trim(a)) = 3 Then 'And Left(a, 1) = 1 And Val(Mid(a, 2, 1)) = 0 And Val(Right(a, 1)) <> 0 Then
                      
                     ' Else
                        hm = Val(Left(figure, 1))
                        If Left(figure, 1) = 3 Then
                          UnitwordHun(hm) = "ËáÇË"
                        ElseIf Left(figure, 1) = 4 Then
                          UnitwordHun(hm) = "ÇÑÈÚ"
                        ElseIf Left(figure, 1) = 5 Then
                          UnitwordHun(hm) = "ÍãÓ"
                        ElseIf Left(figure, 1) = 6 Then
                          UnitwordHun(hm) = "ÓÊ"
                        ElseIf Left(figure, 1) = 7 Then
                          UnitwordHun(hm) = "ÓÈÚ"
                        ElseIf Left(figure, 1) = 8 Then
                          UnitwordHun(hm) = "ËãÇä"
                        ElseIf Left(figure, 1) = 9 Then
                          UnitwordHun(hm) = "ÊÓÚ"
                       End If
                       
                       wordit = wordit & UnitwordHun(hm) & Trim(mHundred) '& " " & mMillion
                      
                      If Len(Trim(a)) = 3 And Val(Right(a, 2)) < 21 Then
                            lastmh = Val(Right(a, 2))
                            wah = " æ "
'                           If Len(Trim(Amt)) = 9 And Val(Mid(Amt, 4, 6)) = 0 Then
'                             wah6 = " "
'                             Else
'                             wah6 = " æ "
'                            End If
                            If Val(Right(a, 1)) = 0 Then
                              wah = " "
                              lastmh = Val(Mid(a, 2, 1))
                              wordit = wordit & wah & Tenword(lastmh) & mMillion
                             Else
                              
                              wordit = wordit & wah & Unitword(lastmh) & mMillion & wah6
                            End If
                          Else
                            'wah = " æ "
                            Lastdigit = Val(Right(a, 1))
                            Middigit = Val(Mid(a, 2, 1))
                            If Val(Mid(figure, 2, 1)) = 0 And Val(Mid(figure, 3, 1)) = 0 Then
                              WAH4 = "  "
                              wah = "  "
                              wah5 = " "
                              ElseIf Val(Mid(figure, 2, 1)) = 0 And Val(Mid(figure, 3, 1)) <> 0 Then
                              WAH4 = " æ "
                              wah5 = " "
                             ElseIf Val(Mid(figure, 2, 1)) <> 0 And Val(Mid(figure, 3, 1)) = 0 Then
                              WAH4 = " æ "
                              wah5 = " "
                              wah = ""
                             Else
                             WAH4 = " æ "
                             wah5 = " "
                             wah = " æ "
                            End If
'                            If Val(Mid(Amt, 4, 6)) <> 0 Then
'                               wah5 = " æ "
'                            End If
                            wordit = wordit & wah & Unitword(Lastdigit) & WAH4 & Tenword(Middigit) & mMillion & wah5
                         End If
                       
                    End If
                      
                
                If a > 0 Then
                    If Len(figure) = 7 Then
                        If Left(figure, 1) = 1 Then
                            If Val(Mid(figure, 2, 6)) > 0 Then
                               'wah7 = " æ "
                            End If
                            wordit = wordit & " " & mMillion & wah7
                          ElseIf Left(figure, 1) = 2 Then
                            wah = " æ "
                            l = Val(Left(figure, 1))
                            wordit = wordit & " " & " ãáíæäÇä "
                          Else
                           wah = " æ "
                           wordit = wordit & " " & Unitword(a) & mMillion
                      End If
                     
                     ElseIf Len(figure) = 8 Then
                         If Val(Mid(figure, 2, 1)) = 0 Then
                            wah = " æ "
                            l = Val(Left(figure, 1))
                            wordit = wordit & " " & Tenword(l) & mMillion
                          ElseIf Val(Mid(figure, 2, 1)) <> 0 Then
                            wah = " æ "
                            l = Val(Mid(figure, 2, 1))
                            f = Val(Left(figure, 1))
                            First2 = f & l
                            First2 = Val(First2)
                            
                            If Left(figure, 2) < 21 Then
                              wordit = wordit & " " & Unitword(First2) & " " & IIf(Mid(figure, 3, 1) <> 0, mMillion, mMillion)
                             Else
                             wordit = wordit & " " & Unitword(l) & wah & Tenword(f) & IIf(Mid(figure, 3, 1) <> 0, mMillion, mMillion)
                            End If
                          
                          End If
                     
                     ElseIf Len(d) = 3 Then
                         l = Val(Left(d, 1))
                         c = Val(Mid(d, 2, 1))
                         r = Val(Right(d, 1))
                         first3 = Val(l) & Val(c) & Val(r)
                         first3 = first3
                         wordit = wordit & " " & Unitword(first3) & " " & " ãáÇííä "
                         wah3 = ""
                     End If
                    
                     wAHcHAR = Trim(Right(Trim(wordit), 2))
                    If wAHcHAR <> "æ" Then
                        wah = " æ "
                      Else
                        wah = "  "
                    End If
                    figure = figure - (1000000 * a)
                    If figure <> 0 Then
                     wordit = wordit & wah   'IIf(FIGURE < 10000000, wah3, WAH)
                    End If
                End If
        End If
        
        
        If figure >= 100000 And figure < 1000000 Then
                a = Int(figure / 100000)
                If a > 0 Then
                    wAHcHAR = Trim(Right(Trim(wordit), 2))
                    If wAHcHAR <> "æ" And wordit <> "" Then
                        wah = " æ "
                      Else
                        wah = "  "
                    End If
                
                    'mHundred = "ãÇÆÉ"
                   If Left(figure, 1) = 1 Then
                    UnitwordHun(a) = "ãÇÆÉ"
                    ElseIf Left(figure, 1) = 2 Then
                     UnitwordHun(a) = "ãÇÆÊÇä"
                    ElseIf Left(figure, 1) = 3 Then
                      UnitwordHun(a) = "ËáÇË"
                    ElseIf Left(figure, 1) = 4 Then
                      UnitwordHun(a) = "ÇÑÈÚ"
                    ElseIf Left(figure, 1) = 5 Then
                      UnitwordHun(a) = "ÍãÓ"
                    ElseIf Left(figure, 1) = 6 Then
                      UnitwordHun(a) = "ÓÊ"
                    ElseIf Left(figure, 1) = 7 Then
                      UnitwordHun(a) = "ÓÈÚ"
                    ElseIf Left(figure, 1) = 8 Then
                      UnitwordHun(a) = "ËãÇä"
                    ElseIf Left(figure, 1) = 9 Then
                      UnitwordHun(a) = "ÊÓÚ"
                   End If
                   'mHundred = "ãÇÆÊÇä"
                   If Mid(figure, 3, 1) <> 0 And Len(Trim(Amt)) < 7 Then
                     wah2 = " æ "
                    Else
                     
                     wah2 = " "
                     If Len(Trim(Amt)) = 6 And Mid(Amt, 2, 2) <> 0 Then
                        wah2 = " æ "
                     End If
                   End If
'                    If wah6 <> "" Then
'                        WAH = " "
'                       Else
'                        WAH = " æ "
'                    End If
'                    If Len(Trim(Amt)) > 6 Then
'                       wah8 = " æ "
'                      Else
'                       wah8 = " "
'                    End If
'                    If Val(Mid(FIGURE, 2, 2)) = 0 Then
'                      WAH = " "
'                     Else
'                     WAH = " æ "
'                    End If
                    
'                    If Val(Mid(Amt, 2, 2)) = 0 Then
'                      wah8 = " æ "
'                      Else
'                      wah8 = " "
'                    End If
                    wordit = wordit & wah8 & Trim(UnitwordHun(a)) & IIf(Val(Left(figure, 1)) > 2, mHundred, "") & wah  '& mHundred
                    'If Mid(figure, 3, 1) <> 0 And Len(Trim(figure)) = 6 Then
                      'for alaf and alf
                      If Val(Mid(figure, 2, 2)) > 2 And Val(Mid(figure, 2, 2)) < 11 Then
                         ht = " ÇáÇÝ "
                        Else
                         ht = " ÇáÝ "
                      End If
                      If Val(Mid(figure, 4, 1)) = 0 Then
                         wah = " "
                        Else
                        wah = " æ "
                      End If
                      If Val(Mid(figure, 3, 1)) = 0 Then
                        a = Val(Mid(figure, 2, 1))
                        If a = 0 Then
                          wah = " "
                        End If
                        wordit = wordit & wah & Tenword(a) & ht & wah
                       ElseIf Val(Mid(figure, 3, 1)) <> 0 And Val(Mid(figure, 2, 1)) > 1 Then
                        wah = " æ "
                        mid3 = Val(Mid(figure, 3, 1))
                        mid2 = Val(Mid(figure, 2, 1))
                        wordit = wordit & wah & Unitword(mid3) & " æ " & Tenword(mid2) & ht & wah
                       Else
                        If Val(Right(figure, 3)) = 0 Then
                         wah = " "
                         Else
                        wah = " æ "
                        End If
                        a = Val(Mid(figure, 2, 2))
                        wordit = wordit & wah & Unitword(a) & ht & wah
                       End If
                      
                     Else
                      If Len(Trim(Amt)) < 7 Then
                        a = Val(Mid(Amt, 4, 1))
                       Else
                        a = 0
                      End If
                    'End If
                    
                    If Mid(figure, 3, 1) <> 0 And Mid(figure, 2, 1) = 0 Then
                      
                       If Mid(figure, 3, 1) = 1 Or Mid(figure, 2, 1) > 1 Then
                         mthousand = " ÇáÝ "
                       ElseIf Mid(figure, 3, 1) = 2 Then
                         mthousand = " ÇáÝÇä "
                       Else
                       
                         mthousand = " ÇáÇÝ "
                       End If
                       
                       a = Val(Mid(figure, 3, 1))
                       
                     ElseIf Mid(Amt, 3, 1) = 0 And Mid(Amt, 2, 1) <> 0 Then
                       If Len(Trim(Amt)) > 6 Then
                        wordit = wordit & Unitword(a)
                       End If
                       a = Val(Mid(figure, 2, 1))
                       
                       wah = " æ "
                       
                     Else
                     
                      If Val(Mid(figure, 2, 2)) > 0 And Val(Mid(figure, 2, 2)) < 21 Or Val(Mid(figure, 3, 1)) < 2 Then
                         a = Val(Mid(figure, 2, 2))
                       Else
                         a = Val(Mid(figure, 3, 1))
                      End If
                       mthousand = " ÇáÝ  "
                       wah = " æ "
                     End If
                      If Len(Trim(Amt)) < 7 Then
                       If Val(Mid(figure, 3, 1)) = 0 Then
                          wordit = wordit & Tenword(a)
                        ElseIf Val(Mid(figure, 3, 1)) <> 0 And Val(Mid(figure, 2, 1)) > 1 Then
                          wordit = wordit & Unitword(mid3) & " æ " & Tenword(mid2)
                          'wah = " "
                        Else
                          wordit = wordit & Unitword(a)
                          wah = " æ "
                       End If
                        wordit = wordit & mthousand & wah
                        wah = " æ "
                       Else
                        If Mid(figure, 2, 1) = 0 And Len(Trim(Amt)) > 6 Then
                            wah = " "
                          Else
                           wah = " æ "
                        End If
                        wordit = wordit & Unitword(a) & wah ' & Tenword(a) & mThousand, IIf(Mid(Amt, 2, 1) = "0", mThousand, " ") & wah)
                        a = Val(Mid(figure, 2, 1))
                        wordit = wordit & Tenword(a) '& mThousand
                        figure = Val(Right(figure, 3))
                        wah = ""
                        wah1 = " æ "
                      End If
                      
                    wordit = wordit & IIf(figure < 1000, wah & mthousand & wah11, "")
                    If Val(Mid(Amt, 2, 1)) <> 0 And Val(Mid(Amt, 3, 1)) <> 0 Then
                      If Len(Trim(Amt)) > 6 Then
                        a = Val(Left(figure, 1))
                        wordit = wordit '& UnitwordHun(a) '(a) & mThousand & " "
                        
                       Else
                        a = Val(Mid(Amt, 2, 1))
                        'wordit = wordit & Tenword(a) & mThousand & " "
                       End If
                       
                      ElseIf Mid(Amt, 2, 1) <> 0 Or Val(Mid(Amt, 3, 1)) = 0 Then
                       a = Val(Mid(Amt, 2, 1))
                      Else
                       a = Val(Right(figure, 3))
                       'wordit = wordit & Tenword(a) & mThousand & " "
                    End If
                    If Len(Trim(Amt)) = 6 And Val(Right(Amt, 2)) = 0 Then
                      figure = Val(Right(figure, 3))
                     ElseIf Len(Trim(Amt)) = 6 And Val(Right(Amt, 2)) <> 0 Then
                      figure = Val(Right(figure, 3))
                     Else
                      If Val(Mid(figure, 2, 1)) = 0 And Right(figure, 2) <> 0 Then
                        'figure = Val(Right(Amt, 1)) '- (100000 * a)
                       ElseIf Right(figure, 2) <> 0 Then
                       'figure = Val(Right(Amt, 3)) '- (100000 * a)
                      End If
                    End If
                'End
                 End If
                 figure = Val(Right(Amt, 3))
        End If
    
        If figure >= 20000 And figure < 100000 Then
                a = Int(figure / 10000)
                If Mid(figure, 3, 1) > 0 Then
                  b = Val(Right(figure, 1))
                End If
                If Mid(figure, 2, 1) > 0 Then
                  b = Val(Mid(figure, 2, 1))
                End If
                If a > 0 Then
                     wAHcHAR = Trim(Right(Trim(wordit), 2))
                    If wAHcHAR <> "æ" And wordit <> "" Then
                        wah = " æ "
                      Else
                        wah = "  "
                    End If
                    If Mid(figure, 3, 1) <> 0 And Mid(figure, 2, 1) = 0 Then
                      Unitword(b) = ""
                    End If
                    If Mid(figure, 2, 1) <> 0 And Mid(figure, 3, 1) <> 0 Then
                      b = Val(Mid(figure, 2, 1))
                      wah = " æ "
                      
                    End If
                    If Mid(Amt, 3, Len(Amt)) = 0 Then
                     wordit = wordit & Unitword(b) & wah & Tenword(a)
                    Else
                     If Mid(Amt, 2, 1) <> 0 Then
                        wordit = wordit & Unitword(b) & " æ " & Tenword(a)
                       Else
                        wordit = wordit & IIf(Mid(figure, 3, 1) > 0, Unitword(b) & wah & Tenword(a), Tenword(a))
                     End If
                    End If
                    figure = figure - (10000 * a)
                    wordit = wordit & IIf(figure < 1000, " " & mthousand & " ", mthousand)
                    figure = Val(Right(figure, 3))
                End If
        End If
        
        If figure > 10000 Then
                a = Int(figure / 1000)
                If a > 0 Then
                    X = Err.Number
                    c = Err.Description
                    wordit = wordit & Unitword(a)
                    wAHcHAR = Trim(Right(Trim(wordit), 2))
                    If wAHcHAR <> "æ" And wordit <> "" Then
                        wah = " æ "
                      Else
                        wah = "  "
                    End If
                    If Left(figure, 2) = "10" Then
                        mthousand = " ÇáÇÝ "
                     End If
                    figure = figure - (1000 * a)
                    
                    wordit = wordit & cwah & IIf(figure < 1000, " " & mthousand & " ", "")
                End If
        End If
        
        
        If figure >= 1000 Then
               a = Int(figure / 1000)
                If Len(figure) = 4 Then
                    If a = 1 Then
                        Unitword(a) = " ÇáÝ  " 'for 1000
                      ElseIf a = 2 Then
                        Unitword(a) = " ÇáÝÇä " 'for 2000
                    End If
                 End If
                If a > 0 Then
                   wAHcHAR = Trim(Right(Trim(wordit), 2))
                    If wAHcHAR <> "æ" And wordit <> "" Then
                        wah = " æ "
                      Else
                        wah = "  "
                    End If
                    
                    'If Amt < 11000 And Amt > 2000 Then
                     If figure < 11000 And figure > 2000 Then
                      mthousand = " ÇáÇÝ "
                     Else
                      mthousand = "ÇáÝ "
                      If Len(Trim(Amt)) = 9 And Val(Mid(figure, 7, 1)) = 0 Then
                        wah2 = "  "
                       Else
                         If Len(Trim(Amt)) = 4 And Val(Right(Amt, 3)) = 0 Then
                           wah2 = " "
                           mthousand = ""
                          Else
                           If Val(Left(figure, 1)) = 1 Then
                             mthousand = ""
                             wah2 = " "
                            Else
                            wah2 = " æ "
                           End If
                         End If
                      End If
                    End If
                    'wordit = wordit & wah & Unitword(A) & IIf(Left(Amt, 1) > 2 Or Left(Amt, 2) = 10, " " & mthousand & " ", "")
                     
                    wordit = wordit & wah & Unitword(a) & mthousand & wah2
                    figure = figure - (1000 * a)
                End If
        End If


        If figure >= 100 Then
              a = Int(figure / 100)
                If Len(figure) = 3 Then
                    If Left(figure, 1) = 1 Then
                        mHundred = " ãÇÆÉ "
                     ElseIf Left(figure, 1) = 2 Then
                        mHundred = " ãÇÆÊÇä "
                     End If
                 End If
                                   
                If a > 0 Then
                    wAHcHAR = Trim(Right(Trim(wordit), 2))
                    If wAHcHAR <> "æ" And wordit <> "" Then
                        wah = " æ "
                      Else
                        wah = "  "
                    End If
                    If Len(figure) > 2 Then
                      If Left(figure, 1) = 3 Then
                         nUnitword(a) = "ËáÇË"
                       ElseIf Left(figure, 1) = 4 Then
                         nUnitword(a) = "ÇÑÈÚ"
                       ElseIf Left(figure, 1) = 5 Then
                         nUnitword(a) = "ÍãÓ"
                       ElseIf Left(figure, 1) = 6 Then
                         nUnitword(a) = "ÓÊ"
                       ElseIf Left(figure, 1) = 7 Then
                         nUnitword(a) = "ÓÈÚ"
                       ElseIf Left(figure, 1) = 8 Then
                         nUnitword(a) = "ËãÇä"
                       ElseIf Left(figure, 1) = 9 Then
                         nUnitword(a) = "ÊÓÚ"
                      End If
                    End If
                    wordit = wordit & wah & nUnitword(a) & Trim(mHundred)   '& ' " "
                    figure = figure - (100 * a)
                End If
        End If
        
        If figure >= 20 And figure < 100 Then
                a = Int(figure / 10)
                If Right(figure, 1) > 0 Then
                  b = Val(Right(figure, 1))
                End If
                a = Int(figure / 10)
                If a > 0 Then
                    wAHcHAR = Trim(Right(Trim(wordit), 2))
                    If wAHcHAR <> "æ" And wordit <> "" Then
                        wah = " æ "
                      Else
                        wah = "  "
                    End If
                    If Right(Amt, 1) = 1 Then
                     Unitword(1) = " æÇÍÏ "
                    ElseIf Right(Amt, 1) = 2 Then
                     Unitword(2) = " ÇËäÇä "
                    ElseIf Right(Amt, 1) = 3 Then
                     Unitword(3) = " ËáÇËÉ "
                    ElseIf Right(Amt, 1) = 4 Then
                     Unitword(4) = " ÇÑÈÚÉ "
                    ElseIf Right(Amt, 1) = 5 Then
                     Unitword(5) = " ÎãÓÉ "
                    ElseIf Right(Amt, 1) = 6 Then
                     Unitword(6) = " ÓÊÉ "
                    ElseIf Right(Amt, 1) = 7 Then
                     Unitword(7) = " ÓÈÚÉ "
                    ElseIf Right(Amt, 1) = 8 Then
                     Unitword(8) = " ËãÇäíÉ "
                    ElseIf Right(Amt, 1) = 9 Then
                     Unitword(9) = " ÊÓÚÉ "
                    End If
                    wordit = wordit & wah & IIf(Right(figure, 1) > 0, Unitword(b) & " æ " & Tenword(a), Tenword(a))
                    If Mid(Amt, 3, 1) <> 0 And Right(Amt, 1) = 0 Then
                      figure = figure - (10 * a)
                     Else
                      figure = 0
                    End If
                End If
        End If
    
        If figure >= 1 Then
                a = Int(figure / 1)
                If a > 0 Then
                    wAHcHAR = Trim(Right(Trim(wordit), 2))
                    If wAHcHAR <> "æ" And wordit <> "" Then
                        wah = " æ "
                      Else
                        wah = "  "
                    End If
                    Dim LTDNZ As String
                    LTDNZ = Val(Right(Amt, 2))
                    Unitword(9) = " ÊÓÚÉ "
                    Unitword(1) = " æÇÍÏ "
                    Unitword(2) = " ÇËäÇä"
                    If Val(Left(LTDNZ, 1)) > 0 Then '< 99 And Amt > 1000 Then
                      wordit = wordit & wah & Unitword(a)
                    End If
                    figure = figure - a
                End If
        End If
       
       'manipulating the decimal
       If Val(mDec) <> 0 Then
          WordDec = mCents
        Else
          WordDec = ""
       End If
       If Amt = 0 Then
        mCurrency = ""
        wah = ""
       Else
       wah = " æ "
       End If
       'wordit = " ÝÞØ " & wordit & " " & " " & mCurrency & " " & IIf(Val(mDec) <> 0, wah, "")
       wordit = " ÝÞØ " & wordit & " " & " " & mCurrency & " " & IIf(Val(mDec) <> 0, wah & " " & mDec, "") & " " & WordDec & " áÇÛíÑ "
'       If Val(Right(mdec, 1)) <> 0 And Val(Left(mdec, 1)) = 0 Or Val(Left(mdec, 1)) = 1 Then
'          a = Val(Right(mdec, 1))
'          wordit = wordit & Unitword(a) & " áÇÛíÑ "
'        ElseIf Val(Right(mdec, 1)) = 0 And Val(Left(mdec, 1)) <> 0 Then
'          a = Val(Left(mdec, 1))
'          wordit = wordit & Tenword(a) & " áÇÛíÑ "
'        Else
'          TW = Left(mdec, 1)
'          UW = Right(mdec, 1)
'          wah = " æ "
'          wordit = wordit & Unitword(UW) & wah & Tenword(TW) & " áÇÛíÑ "
'        End If
       'My Final Word

       mvarKupal = wordit
       'MsgBox (wordit)
MyMsg:
c = Err.Number
d = Err.Description
If Err.Number = 9 Then
    msg = MsgBox("Amount too big", vbOKOnly + vbExclamation, Err.Description)
End If

'Form1.Show
Exit Sub
End Sub

Public Sub EnglishAmountInWords(AmtInFigure As Currency, AmtInwords As Variant, mCurrency As Variant, mCents As Variant)
On Error GoTo MyMsg
Dim wordit As String
Dim Amt As Currency
Amt = AmtInFigure
figure = AmtInFigure
Dim Unitword() As String
Dim Tenword() As String
ReDim Unitword(20)
ReDim Tenword(9)

Unitword(1) = "One"
Unitword(2) = "Two"
Unitword(3) = "Three"
Unitword(4) = "Four"
Unitword(5) = "Five"
Unitword(6) = "Six"
Unitword(7) = "Seven"
Unitword(8) = "Eight"
Unitword(9) = "Nine"
Unitword(10) = "Ten"
Unitword(11) = "Eleven"
Unitword(12) = "Twelve"
Unitword(13) = "Thirteen"
Unitword(14) = "Fourteen"
Unitword(15) = "Fifteen"
Unitword(16) = "Sixteen"
Unitword(17) = "Seventeen"
Unitword(18) = "Eighteen"
Unitword(19) = "Nineteen"
Unitword(20) = "Twenty"
Tenword(1) = "Ten"
Tenword(2) = "Twenty"
Tenword(3) = "Thirty"
Tenword(4) = "Forty"
Tenword(5) = "Fifty"
Tenword(6) = "Sixty"
Tenword(7) = "Seventy"
Tenword(8) = "Eighty"
Tenword(9) = "Ninety"

wordit = ""
        
        If figure >= 1000000 And figure < 1000000000 Then
                a = Int(figure / 1000000)
                If a > 0 Then
                    d = Trim(Str(a))
                    If Len(d) = 1 Then
                      wordit = wordit & " " & Unitword(a)
                     ElseIf Len(d) = 2 Then
                      l = Val(Left(d, 1))
                      r = Val(Right(d, 1))
                      wordit = wordit & " " & Tenword(l) & " " & Unitword(r)
                      ' wordit = wordit & " " & Unitword(a)
                     ElseIf Len(d) = 3 Then
                      l = Val(Left(d, 1))
                      c = Val(Mid(d, 2, 1))
                      r = Val(Right(d, 1))
                      wordit = wordit & " " & Unitword(l) & " Hundred " & Tenword(c) & " " & Unitword(r)
                     End If
                    figure = figure - (1000000 * a)
                    wordit = wordit & " " & IIf(figure < 10000000, "Million", "")
                End If
        End If
        
        If figure >= 100000 And figure < 1000000 Then
                a = Int(figure / 100000)
                If a > 0 Then
                    wordit = wordit & " " & Unitword(a) & " Hundred"
                    figure = figure - (100000 * a)
                    wordit = wordit & " " & IIf(figure < 1000, " Thousand ", "")
                End If
        End If
    
        If figure >= 20000 And figure < 100000 Then
                a = Int(figure / 10000)
                If a > 0 Then
                    wordit = wordit & " " & Tenword(a)
                    figure = figure - (10000 * a)
                    wordit = wordit & " " & IIf(figure < 1000, " Thousand ", "")
                End If
        End If
        
        If figure > 10000 Then
                a = Int(figure / 1000)
                If a > 0 Then
                    X = Err.Number
                    c = Err.Description
                    wordit = wordit & Unitword(a)
                    
                    figure = figure - (1000 * a)
                    wordit = wordit & " " & IIf(figure < 1000, " Thousand ", "")
                End If
        End If
        
        If figure >= 1000 Then
                a = Int(figure / 1000)
                If a > 0 Then
                    wordit = wordit & " " & Unitword(a) & " Thousand "
                    figure = figure - (1000 * a)
                End If
        End If

        If figure >= 100 Then
                a = Int(figure / 100)
                If a > 0 Then
                    wordit = wordit & " " & Unitword(a) & " Hundred "
                    figure = figure - (100 * a)
                End If
        End If
        
        If figure >= 20 And figure < 100 Then
                a = Int(figure / 10)
                If a > 0 Then
                    wordit = wordit & " " & Tenword(a)
                    figure = figure - (10 * a)
                End If
        End If
    
        If figure >= 1 Then
                a = Int(figure / 1)
                If a > 0 Then
                    wordit = wordit & " " & Unitword(a)
                    figure = figure - a
                End If
        End If
       
       'manipulating the Piastres
        If figure > 0# And figure < 0.1 And Amt >= 1 Then
            wordit = wordit & " and " & "0" & Left(figure * 100, 2) & "/100 " & mCurrency & " Only"
          ElseIf figure < 1 And figure > 0.09 And Amt >= 1 Then
            wordit = wordit & " and " & Left(figure * 100, 2) & "/100 " & mCurrency & " Only"
          ElseIf figure = 0 And Amt > 0 Then
            wordit = wordit & " and " & Left(figure * 100, 2) & "0/100 " & mCurrency & " Only"
          ElseIf figure > 0 And figure < 0.1 Then
            wordit = wordit & "0" & Left(figure * 100, 2) & "/100 " & mCents & " only"
          Else
            wordit = wordit & Left(figure * 100, 2) & "/100 " & mCents & " only"
        End If
        AmtInwords = wordit
        mvarKupal = wordit
        'MsgBox (wordit)
MyMsg:
Exit Sub
End Sub
Public Sub GetTables(rst As ADODB.Recordset, CON1 As ADODB.Connection, xtable As String, constring As String, sqltable As Boolean)
On Error GoTo cErrorOnConek
Dim cErr As Long
CON1.Open constring
rst.CursorLocation = adUseClient
rst.CursorType = adOpenDynamic
rst.LockType = adLockOptimistic
rst.Open xtable, CON1, , , IIf(sqltable = True, adCmdText, adCmdTable)

'if error occurs
cErrorOnConek:
cErr = Err.Number
xx = Err.Description

 If cErr = -2147467259 Then
    X = Err.Description
  xmsg = MsgBox("DSN not found. Please Check it at ODBC", vbOK + vbExclamation, "Connection Error")
  Exit Sub
 ElseIf cErr = -2147217865 Then
  xmsg = MsgBox("Table not found, Please check it out.", vbOKOnly + vbExclamation, "Connection Error")
  Exit Sub
 ElseIf cErr = -2147217900 Then
  xmsg = MsgBox(xx, vbOKOnly + vbExclamation, "Connection Error")
  Exit Sub
   
 ElseIf cErr = 91 Then
  xmsg = MsgBox(Err.Description, vbRetryCancel + vbExclamation, "Connection Error")
 ElseIf cErr = 0 Then
  mConnect = True
  Exit Sub
 End If
 If xmsg = vbRetry Then
   Set rst = New ADODB.Recordset
   Set CON1 = New ADODB.Connection
   GetTables rst, CON1, xtable, constring, sqltable
  Else
   mConnect = False
  End If
 End Sub


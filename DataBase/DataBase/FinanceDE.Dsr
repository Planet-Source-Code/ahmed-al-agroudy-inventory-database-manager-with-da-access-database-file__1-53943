VERSION 5.00
Begin {C0E45035-5775-11D0-B388-00A0C9055D8E} FinanceDE 
   ClientHeight    =   9435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   _ExtentX        =   21167
   _ExtentY        =   16642
   FolderFlags     =   7
   TypeLibGuid     =   "{41C6361F-65B2-4FCE-9CAF-353C82B9B293}"
   TypeInfoGuid    =   "{C12790DA-016F-4262-96D5-38C2C76B4B62}"
   TypeInfoCookie  =   0
   Version         =   4
   NumConnections  =   1
   BeginProperty Connection1 
      ConnectionName  =   "FinanceCON"
      ConnDispId      =   1001
      SourceOfData    =   3
      ConnectionSource=   "Provider=MSDASQL.1;Password="""";Persist Security Info=True;User ID=sa;Data Source=finance"
      Expanded        =   -1  'True
      IsSQL           =   -1  'True
      QuoteChar       =   34
      SeparatorChar   =   46
   EndProperty
   NumRecordsets   =   21
   BeginProperty Recordset1 
      CommandName     =   "TopLevel"
      CommDispId      =   1002
      RsDispId        =   1288
      CommandText     =   "select * from Toplevel where country=?"
      ActiveConnectionName=   "FinanceCON"
      CommandType     =   1
      GroupingName    =   "Level1_Grouping"
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   7
      BeginProperty Field1 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   202
         Name            =   "Country"
         Caption         =   "Country"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "CountryName"
         Caption         =   "CountryName"
      EndProperty
      BeginProperty Field3 
         Precision       =   10
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "code"
         Caption         =   "code"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   202
         Name            =   "AccountCode"
         Caption         =   "AccountCode"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "AccountNameEng"
         Caption         =   "AccountNameEng"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "AccountNameArab"
         Caption         =   "AccountNameArab"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   202
         Name            =   "Remarks"
         Caption         =   "Remarks"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   2
         Scale           =   0
         Size            =   2
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset2 
      CommandName     =   "level1"
      CommDispId      =   -1
      RsDispId        =   -1
      CommandText     =   "Select * from Level1 order by AccountCode"
      ActiveConnectionName=   "FinanceCON"
      CommandType     =   1
      GroupingName    =   "level2_Grouping"
      RelateToParent  =   -1  'True
      ParentCommandName=   "TopLevel"
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   9
      BeginProperty Field1 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   202
         Name            =   "Country"
         Caption         =   "Country"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "CountryName"
         Caption         =   "CountryName"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "TopLevelCode"
         Caption         =   "TopLevelCode"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "TopLevelName"
         Caption         =   "TopLevelName"
      EndProperty
      BeginProperty Field5 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "Code"
         Caption         =   "Code"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   202
         Name            =   "AccountCode"
         Caption         =   "AccountCode"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "AccountNameEng"
         Caption         =   "AccountNameEng"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "AccountNameArab"
         Caption         =   "AccountNameArab"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   202
         Name            =   "Remarks"
         Caption         =   "Remarks"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   1
      BeginProperty Relation1 
         ParentField     =   "code"
         ChildField      =   "TopLevelCode"
         ParentType      =   0
         ChildType       =   0
      EndProperty
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset3 
      CommandName     =   "Level2"
      CommDispId      =   -1
      RsDispId        =   -1
      CommandText     =   "Select * from Level2 order by AccountCode"
      ActiveConnectionName=   "FinanceCON"
      CommandType     =   1
      GroupingName    =   "Level3_Grouping"
      RelateToParent  =   -1  'True
      ParentCommandName=   "level1"
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   11
      BeginProperty Field1 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   202
         Name            =   "Country"
         Caption         =   "Country"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   202
         Name            =   "CountryName"
         Caption         =   "CountryName"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   202
         Name            =   "TopLevelCode"
         Caption         =   "TopLevelCode"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "TopLevelName"
         Caption         =   "TopLevelName"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   202
         Name            =   "Level1Code"
         Caption         =   "Level1Code"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "Level1Name"
         Caption         =   "Level1Name"
      EndProperty
      BeginProperty Field7 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "Code"
         Caption         =   "Code"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   202
         Name            =   "AccountCode"
         Caption         =   "AccountCode"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "AccountNameEng"
         Caption         =   "AccountNameEng"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "AccountNameArab"
         Caption         =   "AccountNameArab"
      EndProperty
      BeginProperty Field11 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   202
         Name            =   "Remarks"
         Caption         =   "Remarks"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   2
      BeginProperty Relation1 
         ParentField     =   "TopLevelCode"
         ChildField      =   "TopLevelCode"
         ParentType      =   0
         ChildType       =   0
      EndProperty
      BeginProperty Relation2 
         ParentField     =   "Code"
         ChildField      =   "Level1Code"
         ParentType      =   0
         ChildType       =   0
      EndProperty
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset4 
      CommandName     =   "level3"
      CommDispId      =   -1
      RsDispId        =   -1
      CommandText     =   "Select * from Level3 order by AccountCode"
      ActiveConnectionName=   "FinanceCON"
      CommandType     =   1
      GroupingName    =   "level4_Grouping"
      RelateToParent  =   -1  'True
      ParentCommandName=   "Level2"
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   14
      BeginProperty Field1 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   202
         Name            =   "MainAcct"
         Caption         =   "MainAcct"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   202
         Name            =   "Country"
         Caption         =   "Country"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "CountryName"
         Caption         =   "CountryName"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   202
         Name            =   "TopLevelCode"
         Caption         =   "TopLevelCode"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "TopLevelName"
         Caption         =   "TopLevelName"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   202
         Name            =   "Level1Code"
         Caption         =   "Level1Code"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "Level1Name"
         Caption         =   "Level1Name"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   202
         Name            =   "Level2Code"
         Caption         =   "Level2Code"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "Level2Name"
         Caption         =   "Level2Name"
      EndProperty
      BeginProperty Field10 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "Code"
         Caption         =   "Code"
      EndProperty
      BeginProperty Field11 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   202
         Name            =   "AccountCode"
         Caption         =   "AccountCode"
      EndProperty
      BeginProperty Field12 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "AccountNameEng"
         Caption         =   "AccountNameEng"
      EndProperty
      BeginProperty Field13 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "AccountNameArab"
         Caption         =   "AccountNameArab"
      EndProperty
      BeginProperty Field14 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   202
         Name            =   "Remarks"
         Caption         =   "Remarks"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   3
      BeginProperty Relation1 
         ParentField     =   "TopLevelCode"
         ChildField      =   "TopLevelCode"
         ParentType      =   0
         ChildType       =   0
      EndProperty
      BeginProperty Relation2 
         ParentField     =   "Level1Code"
         ChildField      =   "Level1Code"
         ParentType      =   0
         ChildType       =   0
      EndProperty
      BeginProperty Relation3 
         ParentField     =   "Code"
         ChildField      =   "Level2Code"
         ParentType      =   0
         ChildType       =   0
      EndProperty
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset5 
      CommandName     =   "GenJournal"
      CommDispId      =   1032
      RsDispId        =   1132
      CommandText     =   "Select * from Genjournaltrans where transdate = ? and prepby =? and remarks is null order by serialno"
      ActiveConnectionName=   "FinanceCON"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   21
      BeginProperty Field1 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   202
         Name            =   "SerialNo"
         Caption         =   "SerialNo"
      EndProperty
      BeginProperty Field2 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "Ticket"
         Caption         =   "Ticket"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   202
         Name            =   "DeleteMark"
         Caption         =   "DeleteMark"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   202
         Name            =   "Classification"
         Caption         =   "Classification"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "AccountNumber"
         Caption         =   "AccountNumber"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   202
         Name            =   "AccountName"
         Caption         =   "AccountName"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "GenDesc"
         Caption         =   "GenDesc"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   1073741823
         Scale           =   0
         Type            =   203
         Name            =   "Description"
         Caption         =   "Description"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   202
         Name            =   "TransDate"
         Caption         =   "TransDate"
      EndProperty
      BeginProperty Field10 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "BeginningBal"
         Caption         =   "BeginningBal"
      EndProperty
      BeginProperty Field11 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "DebitAmount"
         Caption         =   "DebitAmount"
      EndProperty
      BeginProperty Field12 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "CreditAmount"
         Caption         =   "CreditAmount"
      EndProperty
      BeginProperty Field13 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "EndingBals"
         Caption         =   "EndingBals"
      EndProperty
      BeginProperty Field14 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "TransType"
         Caption         =   "TransType"
      EndProperty
      BeginProperty Field15 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "Remarks"
         Caption         =   "Remarks"
      EndProperty
      BeginProperty Field16 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "Fcdebit"
         Caption         =   "Fcdebit"
      EndProperty
      BeginProperty Field17 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "FcCredit"
         Caption         =   "FcCredit"
      EndProperty
      BeginProperty Field18 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   202
         Name            =   "status"
         Caption         =   "status"
      EndProperty
      BeginProperty Field19 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "PrepBy"
         Caption         =   "PrepBy"
      EndProperty
      BeginProperty Field20 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "NotedBy"
         Caption         =   "NotedBy"
      EndProperty
      BeginProperty Field21 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "AppBy"
         Caption         =   "AppBy"
      EndProperty
      NumGroups       =   0
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   20
         Scale           =   0
         Size            =   20
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         Direction       =   1
         Precision       =   50
         Scale           =   0
         Size            =   50
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset6 
      CommandName     =   "AssetJournal"
      CommDispId      =   1040
      RsDispId        =   1066
      CommandText     =   $"FinanceDE.dsx":0000
      ActiveConnectionName=   "FinanceCON"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   17
      BeginProperty Field1 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "SerialNo"
         Caption         =   "SerialNo"
      EndProperty
      BeginProperty Field2 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "Ticket"
         Caption         =   "Ticket"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "AccountNumber"
         Caption         =   "AccountNumber"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   202
         Name            =   "AccountName"
         Caption         =   "AccountName"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   202
         Name            =   "Description"
         Caption         =   "Description"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "TransDate"
         Caption         =   "TransDate"
      EndProperty
      BeginProperty Field7 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "BeginningBal"
         Caption         =   "BeginningBal"
      EndProperty
      BeginProperty Field8 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "DebitAmount"
         Caption         =   "DebitAmount"
      EndProperty
      BeginProperty Field9 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "CreditAmount"
         Caption         =   "CreditAmount"
      EndProperty
      BeginProperty Field10 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "EndingBal"
         Caption         =   "EndingBal"
      EndProperty
      BeginProperty Field11 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "Remarks"
         Caption         =   "Remarks"
      EndProperty
      BeginProperty Field12 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "Fcdebit"
         Caption         =   "Fcdebit"
      EndProperty
      BeginProperty Field13 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "FcCredit"
         Caption         =   "FcCredit"
      EndProperty
      BeginProperty Field14 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   202
         Name            =   "status"
         Caption         =   "status"
      EndProperty
      BeginProperty Field15 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "PrepBy"
         Caption         =   "PrepBy"
      EndProperty
      BeginProperty Field16 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "NotedBy"
         Caption         =   "NotedBy"
      EndProperty
      BeginProperty Field17 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "AppBy"
         Caption         =   "AppBy"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   50
         Scale           =   0
         Size            =   50
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset7 
      CommandName     =   "SalesJournal"
      CommDispId      =   1052
      RsDispId        =   1060
      CommandText     =   "select * from SalesJournal  where debitamount<>0 and transdate > ? order by ticket"
      ActiveConnectionName=   "FinanceCON"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   34
      BeginProperty Field1 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   202
         Name            =   "SerialNo"
         Caption         =   "SerialNo"
      EndProperty
      BeginProperty Field2 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "Ticket"
         Caption         =   "Ticket"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   202
         Name            =   "DeleteMark"
         Caption         =   "DeleteMark"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "AccountNumber"
         Caption         =   "AccountNumber"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   202
         Name            =   "AccountName"
         Caption         =   "AccountName"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   1073741823
         Scale           =   0
         Type            =   203
         Name            =   "Description"
         Caption         =   "Description"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   202
         Name            =   "TransDate"
         Caption         =   "TransDate"
      EndProperty
      BeginProperty Field8 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "BeginningBal"
         Caption         =   "BeginningBal"
      EndProperty
      BeginProperty Field9 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "DebitAmount"
         Caption         =   "DebitAmount"
      EndProperty
      BeginProperty Field10 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "CreditAmount"
         Caption         =   "CreditAmount"
      EndProperty
      BeginProperty Field11 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "EndingBals"
         Caption         =   "EndingBals"
      EndProperty
      BeginProperty Field12 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   202
         Name            =   "Remarks"
         Caption         =   "Remarks"
      EndProperty
      BeginProperty Field13 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "InvoiceDate"
         Caption         =   "InvoiceDate"
      EndProperty
      BeginProperty Field14 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "InvoiceNo"
         Caption         =   "InvoiceNo"
      EndProperty
      BeginProperty Field15 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "ClientCode"
         Caption         =   "ClientCode"
      EndProperty
      BeginProperty Field16 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "TradeRcvble"
         Caption         =   "TradeRcvble"
      EndProperty
      BeginProperty Field17 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "TradeDisc"
         Caption         =   "TradeDisc"
      EndProperty
      BeginProperty Field18 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "TradeDiscAmt"
         Caption         =   "TradeDiscAmt"
      EndProperty
      BeginProperty Field19 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "MgtDisc"
         Caption         =   "MgtDisc"
      EndProperty
      BeginProperty Field20 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "MgtDiscAmt"
         Caption         =   "MgtDiscAmt"
      EndProperty
      BeginProperty Field21 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "GrossSales"
         Caption         =   "GrossSales"
      EndProperty
      BeginProperty Field22 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "NetSales"
         Caption         =   "NetSales"
      EndProperty
      BeginProperty Field23 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "TranspoCharge"
         Caption         =   "TranspoCharge"
      EndProperty
      BeginProperty Field24 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "VAT"
         Caption         =   "VAT"
      EndProperty
      BeginProperty Field25 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "SurTaxRate"
         Caption         =   "SurTaxRate"
      EndProperty
      BeginProperty Field26 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "SURTaxAmt"
         Caption         =   "SURTaxAmt"
      EndProperty
      BeginProperty Field27 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "TaxCr"
         Caption         =   "TaxCr"
      EndProperty
      BeginProperty Field28 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "ProfitCenter"
         Caption         =   "ProfitCenter"
      EndProperty
      BeginProperty Field29 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "Fcdebit"
         Caption         =   "Fcdebit"
      EndProperty
      BeginProperty Field30 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "FcCredit"
         Caption         =   "FcCredit"
      EndProperty
      BeginProperty Field31 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   202
         Name            =   "status"
         Caption         =   "status"
      EndProperty
      BeginProperty Field32 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "PrepBy"
         Caption         =   "PrepBy"
      EndProperty
      BeginProperty Field33 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "NotedBy"
         Caption         =   "NotedBy"
      EndProperty
      BeginProperty Field34 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "AppBy"
         Caption         =   "AppBy"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   20
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset8 
      CommandName     =   "InventoryJournal"
      CommDispId      =   1061
      RsDispId        =   1282
      CommandText     =   "Select * from InventoryJournal where TRANSDATE=? AND remarks is null order by Serialno "
      ActiveConnectionName=   "FinanceCON"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   29
      BeginProperty Field1 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "SerialNo"
         Caption         =   "SerialNo"
      EndProperty
      BeginProperty Field2 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "Ticket"
         Caption         =   "Ticket"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "AccountNumber"
         Caption         =   "AccountNumber"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   202
         Name            =   "AccountName"
         Caption         =   "AccountName"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   300
         Scale           =   0
         Type            =   202
         Name            =   "Classification"
         Caption         =   "Classification"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   500
         Scale           =   0
         Type            =   202
         Name            =   "Description"
         Caption         =   "Description"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "TransDate"
         Caption         =   "TransDate"
      EndProperty
      BeginProperty Field8 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "BeginningBal"
         Caption         =   "BeginningBal"
      EndProperty
      BeginProperty Field9 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "DebitAmount"
         Caption         =   "DebitAmount"
      EndProperty
      BeginProperty Field10 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "CreditAmount"
         Caption         =   "CreditAmount"
      EndProperty
      BeginProperty Field11 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "EndingBal"
         Caption         =   "EndingBal"
      EndProperty
      BeginProperty Field12 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "Remarks"
         Caption         =   "Remarks"
      EndProperty
      BeginProperty Field13 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "Fcdebit"
         Caption         =   "Fcdebit"
      EndProperty
      BeginProperty Field14 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "FcCredit"
         Caption         =   "FcCredit"
      EndProperty
      BeginProperty Field15 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   202
         Name            =   "status"
         Caption         =   "status"
      EndProperty
      BeginProperty Field16 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "PrepBy"
         Caption         =   "PrepBy"
      EndProperty
      BeginProperty Field17 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "NotedBy"
         Caption         =   "NotedBy"
      EndProperty
      BeginProperty Field18 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "AppBy"
         Caption         =   "AppBy"
      EndProperty
      BeginProperty Field19 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "GRno"
         Caption         =   "GRno"
      EndProperty
      BeginProperty Field20 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "InventCat"
         Caption         =   "InventCat"
      EndProperty
      BeginProperty Field21 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "Voucher"
         Caption         =   "Voucher"
      EndProperty
      BeginProperty Field22 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "DONo"
         Caption         =   "DONo"
      EndProperty
      BeginProperty Field23 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "Fr_CostCenter"
         Caption         =   "Fr_CostCenter"
      EndProperty
      BeginProperty Field24 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "Fr_Dept"
         Caption         =   "Fr_Dept"
      EndProperty
      BeginProperty Field25 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "To_CostCenter"
         Caption         =   "To_CostCenter"
      EndProperty
      BeginProperty Field26 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "To_Dept"
         Caption         =   "To_Dept"
      EndProperty
      BeginProperty Field27 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "Purpose"
         Caption         =   "Purpose"
      EndProperty
      BeginProperty Field28 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "WONo"
         Caption         =   "WONo"
      EndProperty
      BeginProperty Field29 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "TranType"
         Caption         =   "TranType"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   50
         Scale           =   0
         Size            =   50
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset9 
      CommandName     =   "Level4"
      CommDispId      =   -1
      RsDispId        =   -1
      CommandText     =   "Select * from Level4 order by AccountCode"
      ActiveConnectionName=   "FinanceCON"
      CommandType     =   1
      RelateToParent  =   -1  'True
      ParentCommandName=   "level3"
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   16
      BeginProperty Field1 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   202
         Name            =   "MainAcct"
         Caption         =   "MainAcct"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   202
         Name            =   "Country"
         Caption         =   "Country"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "CountryName"
         Caption         =   "CountryName"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   202
         Name            =   "TopLevelCode"
         Caption         =   "TopLevelCode"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "TopLevelName"
         Caption         =   "TopLevelName"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   202
         Name            =   "Level1Code"
         Caption         =   "Level1Code"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "Level1Name"
         Caption         =   "Level1Name"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   202
         Name            =   "Level2Code"
         Caption         =   "Level2Code"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "Level2Name"
         Caption         =   "Level2Name"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   202
         Name            =   "Level3Code"
         Caption         =   "Level3Code"
      EndProperty
      BeginProperty Field11 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "Level3Name"
         Caption         =   "Level3Name"
      EndProperty
      BeginProperty Field12 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "Code"
         Caption         =   "Code"
      EndProperty
      BeginProperty Field13 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   202
         Name            =   "AccountCode"
         Caption         =   "AccountCode"
      EndProperty
      BeginProperty Field14 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "AccountNameEng"
         Caption         =   "AccountNameEng"
      EndProperty
      BeginProperty Field15 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "AccountNameArab"
         Caption         =   "AccountNameArab"
      EndProperty
      BeginProperty Field16 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   202
         Name            =   "Remarks"
         Caption         =   "Remarks"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   4
      BeginProperty Relation1 
         ParentField     =   "TopLevelCode"
         ChildField      =   "TopLevelCode"
         ParentType      =   0
         ChildType       =   0
      EndProperty
      BeginProperty Relation2 
         ParentField     =   "Level1Code"
         ChildField      =   "Level1Code"
         ParentType      =   0
         ChildType       =   0
      EndProperty
      BeginProperty Relation3 
         ParentField     =   "Level2Code"
         ChildField      =   "Level2Code"
         ParentType      =   0
         ChildType       =   0
      EndProperty
      BeginProperty Relation4 
         ParentField     =   "Code"
         ChildField      =   "Level3Code"
         ParentType      =   0
         ChildType       =   0
      EndProperty
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset10 
      CommandName     =   "Level5"
      CommDispId      =   -1
      RsDispId        =   -1
      CommandText     =   "Select * from Level5 order by AccountCode"
      ActiveConnectionName=   "FinanceCON"
      CommandType     =   1
      RelateToParent  =   -1  'True
      ParentCommandName=   "Level4"
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   18
      BeginProperty Field1 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   202
         Name            =   "MainAcct"
         Caption         =   "MainAcct"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   202
         Name            =   "Country"
         Caption         =   "Country"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "CountryName"
         Caption         =   "CountryName"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   202
         Name            =   "TopLevelCode"
         Caption         =   "TopLevelCode"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "TopLevelName"
         Caption         =   "TopLevelName"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   202
         Name            =   "Level1Code"
         Caption         =   "Level1Code"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "Level1Name"
         Caption         =   "Level1Name"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   202
         Name            =   "Level2Code"
         Caption         =   "Level2Code"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "Level2Name"
         Caption         =   "Level2Name"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   202
         Name            =   "Level3Code"
         Caption         =   "Level3Code"
      EndProperty
      BeginProperty Field11 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "Level3Name"
         Caption         =   "Level3Name"
      EndProperty
      BeginProperty Field12 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   202
         Name            =   "Level4Code"
         Caption         =   "Level4Code"
      EndProperty
      BeginProperty Field13 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "Level4Name"
         Caption         =   "Level4Name"
      EndProperty
      BeginProperty Field14 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "Code"
         Caption         =   "Code"
      EndProperty
      BeginProperty Field15 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   202
         Name            =   "AccountCode"
         Caption         =   "AccountCode"
      EndProperty
      BeginProperty Field16 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "AccountNameEng"
         Caption         =   "AccountNameEng"
      EndProperty
      BeginProperty Field17 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "AccountNameArab"
         Caption         =   "AccountNameArab"
      EndProperty
      BeginProperty Field18 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   202
         Name            =   "Remarks"
         Caption         =   "Remarks"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   5
      BeginProperty Relation1 
         ParentField     =   "TopLevelCode"
         ChildField      =   "TopLevelCode"
         ParentType      =   0
         ChildType       =   0
      EndProperty
      BeginProperty Relation2 
         ParentField     =   "Level1Code"
         ChildField      =   "Level1Code"
         ParentType      =   0
         ChildType       =   0
      EndProperty
      BeginProperty Relation3 
         ParentField     =   "Level2Code"
         ChildField      =   "Level2Code"
         ParentType      =   0
         ChildType       =   0
      EndProperty
      BeginProperty Relation4 
         ParentField     =   "Level3Code"
         ChildField      =   "Level3Code"
         ParentType      =   0
         ChildType       =   0
      EndProperty
      BeginProperty Relation5 
         ParentField     =   "Code"
         ChildField      =   "Level4Code"
         ParentType      =   0
         ChildType       =   0
      EndProperty
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset11 
      CommandName     =   "Level6"
      CommDispId      =   -1
      RsDispId        =   -1
      CommandText     =   "Select * from Level6 order by AccountCode"
      ActiveConnectionName=   "FinanceCON"
      CommandType     =   1
      RelateToParent  =   -1  'True
      ParentCommandName=   "Level5"
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   20
      BeginProperty Field1 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   202
         Name            =   "MainAcct"
         Caption         =   "MainAcct"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   202
         Name            =   "Country"
         Caption         =   "Country"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "CountryName"
         Caption         =   "CountryName"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   202
         Name            =   "TopLevelCode"
         Caption         =   "TopLevelCode"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "TopLevelName"
         Caption         =   "TopLevelName"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   202
         Name            =   "Level1Code"
         Caption         =   "Level1Code"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "Level1Name"
         Caption         =   "Level1Name"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   202
         Name            =   "Level2Code"
         Caption         =   "Level2Code"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "Level2Name"
         Caption         =   "Level2Name"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   202
         Name            =   "Level3Code"
         Caption         =   "Level3Code"
      EndProperty
      BeginProperty Field11 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "Level3Name"
         Caption         =   "Level3Name"
      EndProperty
      BeginProperty Field12 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   202
         Name            =   "Level4Code"
         Caption         =   "Level4Code"
      EndProperty
      BeginProperty Field13 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "Level4Name"
         Caption         =   "Level4Name"
      EndProperty
      BeginProperty Field14 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   202
         Name            =   "Level5Code"
         Caption         =   "Level5Code"
      EndProperty
      BeginProperty Field15 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "Level5Name"
         Caption         =   "Level5Name"
      EndProperty
      BeginProperty Field16 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "Code"
         Caption         =   "Code"
      EndProperty
      BeginProperty Field17 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   202
         Name            =   "AccountCode"
         Caption         =   "AccountCode"
      EndProperty
      BeginProperty Field18 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "AccountNameEng"
         Caption         =   "AccountNameEng"
      EndProperty
      BeginProperty Field19 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "AccountNameArab"
         Caption         =   "AccountNameArab"
      EndProperty
      BeginProperty Field20 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   202
         Name            =   "Remarks"
         Caption         =   "Remarks"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   6
      BeginProperty Relation1 
         ParentField     =   "TopLevelCode"
         ChildField      =   "TopLevelCode"
         ParentType      =   0
         ChildType       =   0
      EndProperty
      BeginProperty Relation2 
         ParentField     =   "Level1Code"
         ChildField      =   "Level1Code"
         ParentType      =   0
         ChildType       =   0
      EndProperty
      BeginProperty Relation3 
         ParentField     =   "Level2Code"
         ChildField      =   "Level2Code"
         ParentType      =   0
         ChildType       =   0
      EndProperty
      BeginProperty Relation4 
         ParentField     =   "Level3Code"
         ChildField      =   "Level3Code"
         ParentType      =   0
         ChildType       =   0
      EndProperty
      BeginProperty Relation5 
         ParentField     =   "Level4Code"
         ChildField      =   "Level4Code"
         ParentType      =   0
         ChildType       =   0
      EndProperty
      BeginProperty Relation6 
         ParentField     =   "Code"
         ChildField      =   "Level5Code"
         ParentType      =   0
         ChildType       =   0
      EndProperty
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset12 
      CommandName     =   "TempLevel"
      CommDispId      =   1079
      RsDispId        =   1089
      CommandText     =   "select * from templevel order by AccountNameEng"
      ActiveConnectionName=   "FinanceCON"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   5
      BeginProperty Field1 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "AccountCode"
         Caption         =   "AccountCode"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   202
         Name            =   "AccountNameEng"
         Caption         =   "AccountNameEng"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   202
         Name            =   "AccountNameArab"
         Caption         =   "AccountNameArab"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   202
         Name            =   "Remarks"
         Caption         =   "Remarks"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   202
         Name            =   "MotherCat"
         Caption         =   "MotherCat"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset13 
      CommandName     =   "AssetRegistered"
      CommDispId      =   1090
      RsDispId        =   1095
      CommandText     =   "select * from AssetRegistered order by idno"
      ActiveConnectionName=   "FinanceCON"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   10
      BeginProperty Field1 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "IDNo"
         Caption         =   "IDNo"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "Code"
         Caption         =   "Code"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "Category"
         Caption         =   "Category"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "Sub_Category"
         Caption         =   "Sub_Category"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "NameEng"
         Caption         =   "NameEng"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "NameArab"
         Caption         =   "NameArab"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "SerialNo"
         Caption         =   "SerialNo"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "ModelNo"
         Caption         =   "ModelNo"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "GLAcctCode"
         Caption         =   "GLAcctCode"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "DateRegistered"
         Caption         =   "DateRegistered"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset14 
      CommandName     =   "GenJournalPosted"
      CommDispId      =   1098
      RsDispId        =   1284
      CommandText     =   "Select * from GLMaster where Recorddate = ? and left(journalno,3)=? order by autonumber "
      ActiveConnectionName=   "FinanceCON"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   13
      BeginProperty Field1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "autonumber"
         Caption         =   "autonumber"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "JournalNo"
         Caption         =   "JournalNo"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   202
         Name            =   "AccountCode"
         Caption         =   "AccountCode"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   1073741823
         Scale           =   0
         Type            =   203
         Name            =   "AccountName"
         Caption         =   "AccountName"
      EndProperty
      BeginProperty Field5 
         Precision       =   16
         Size            =   16
         Scale           =   0
         Type            =   135
         Name            =   "PostDate"
         Caption         =   "PostDate"
      EndProperty
      BeginProperty Field6 
         Precision       =   16
         Size            =   16
         Scale           =   0
         Type            =   135
         Name            =   "recorddate"
         Caption         =   "recorddate"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   200
         Scale           =   0
         Type            =   202
         Name            =   "Particulars"
         Caption         =   "Particulars"
      EndProperty
      BeginProperty Field8 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "DebitAmount"
         Caption         =   "DebitAmount"
      EndProperty
      BeginProperty Field9 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "CreditAmount"
         Caption         =   "CreditAmount"
      EndProperty
      BeginProperty Field10 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "Balance"
         Caption         =   "Balance"
      EndProperty
      BeginProperty Field11 
         Precision       =   0
         Size            =   1073741823
         Scale           =   0
         Type            =   203
         Name            =   "Cashdetails"
         Caption         =   "Cashdetails"
      EndProperty
      BeginProperty Field12 
         Precision       =   0
         Size            =   1073741823
         Scale           =   0
         Type            =   203
         Name            =   "Category"
         Caption         =   "Category"
      EndProperty
      BeginProperty Field13 
         Precision       =   0
         Size            =   5
         Scale           =   0
         Type            =   202
         Name            =   "Printed"
         Caption         =   "Printed"
      EndProperty
      NumGroups       =   0
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   16
         Scale           =   0
         Size            =   16
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         Direction       =   1
         Precision       =   3
         Scale           =   0
         Size            =   3
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset15 
      CommandName     =   "BankJournalUnpost"
      CommDispId      =   1134
      RsDispId        =   1146
      CommandText     =   "select * from BankJOurnal where transdate=? and prepby = ? and remarks is null  order by serialno"
      ActiveConnectionName=   "FinanceCON"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   20
      BeginProperty Field1 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   202
         Name            =   "SerialNo"
         Caption         =   "SerialNo"
      EndProperty
      BeginProperty Field2 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "Ticket"
         Caption         =   "Ticket"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   202
         Name            =   "DeleteMark"
         Caption         =   "DeleteMark"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   202
         Name            =   "Classification"
         Caption         =   "Classification"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "AccountNumber"
         Caption         =   "AccountNumber"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   202
         Name            =   "AccountName"
         Caption         =   "AccountName"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "GenDesc"
         Caption         =   "GenDesc"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   1073741823
         Scale           =   0
         Type            =   203
         Name            =   "Description"
         Caption         =   "Description"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   202
         Name            =   "TransDate"
         Caption         =   "TransDate"
      EndProperty
      BeginProperty Field10 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "BeginningBal"
         Caption         =   "BeginningBal"
      EndProperty
      BeginProperty Field11 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "DebitAmount"
         Caption         =   "DebitAmount"
      EndProperty
      BeginProperty Field12 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "CreditAmount"
         Caption         =   "CreditAmount"
      EndProperty
      BeginProperty Field13 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "EndingBals"
         Caption         =   "EndingBals"
      EndProperty
      BeginProperty Field14 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "Remarks"
         Caption         =   "Remarks"
      EndProperty
      BeginProperty Field15 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "Fcdebit"
         Caption         =   "Fcdebit"
      EndProperty
      BeginProperty Field16 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "FcCredit"
         Caption         =   "FcCredit"
      EndProperty
      BeginProperty Field17 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   202
         Name            =   "status"
         Caption         =   "status"
      EndProperty
      BeginProperty Field18 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "PrepBy"
         Caption         =   "PrepBy"
      EndProperty
      BeginProperty Field19 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "NotedBy"
         Caption         =   "NotedBy"
      EndProperty
      BeginProperty Field20 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "AppBy"
         Caption         =   "AppBy"
      EndProperty
      NumGroups       =   0
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   20
         Scale           =   0
         Size            =   20
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         Direction       =   1
         Precision       =   50
         Scale           =   0
         Size            =   50
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset16 
      CommandName     =   "BankBalances"
      CommDispId      =   1147
      RsDispId        =   1190
      CommandText     =   "Select * from BankAccountBalances order by GlAcctCode"
      ActiveConnectionName=   "FinanceCON"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   14
      BeginProperty Field1 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "RateCode"
         Caption         =   "RateCode"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "BankAcctNameeng"
         Caption         =   "BankAcctNameeng"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "BankAcctNameArab"
         Caption         =   "BankAcctNameArab"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "GLAcctCode"
         Caption         =   "GLAcctCode"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "AccountType"
         Caption         =   "AccountType"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "Currency"
         Caption         =   "Currency"
      EndProperty
      BeginProperty Field7 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "LE1"
         Caption         =   "LE1"
      EndProperty
      BeginProperty Field8 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "USDollar1"
         Caption         =   "USDollar1"
      EndProperty
      BeginProperty Field9 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "USDollar2"
         Caption         =   "USDollar2"
      EndProperty
      BeginProperty Field10 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "LE2"
         Caption         =   "LE2"
      EndProperty
      BeginProperty Field11 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "Euro"
         Caption         =   "Euro"
      EndProperty
      BeginProperty Field12 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "LE3"
         Caption         =   "LE3"
      EndProperty
      BeginProperty Field13 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "Balances"
         Caption         =   "Balances"
      EndProperty
      BeginProperty Field14 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "AsOf"
         Caption         =   "AsOf"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset17 
      CommandName     =   "TempGenJournal"
      CommDispId      =   1154
      RsDispId        =   1270
      CommandText     =   "select * from TempGenJournal where prepby=? order by ticket"
      ActiveConnectionName=   "FinanceCON"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   18
      BeginProperty Field1 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   202
         Name            =   "SerialNo"
         Caption         =   "SerialNo"
      EndProperty
      BeginProperty Field2 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "Ticket"
         Caption         =   "Ticket"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   202
         Name            =   "DeleteMark"
         Caption         =   "DeleteMark"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   202
         Name            =   "Classification"
         Caption         =   "Classification"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "AccountNumber"
         Caption         =   "AccountNumber"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   202
         Name            =   "AccountName"
         Caption         =   "AccountName"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   300
         Scale           =   0
         Type            =   202
         Name            =   "GenDesc"
         Caption         =   "GenDesc"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   1073741823
         Scale           =   0
         Type            =   203
         Name            =   "Description"
         Caption         =   "Description"
      EndProperty
      BeginProperty Field9 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "TransDate"
         Caption         =   "TransDate"
      EndProperty
      BeginProperty Field10 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "BeginningBal"
         Caption         =   "BeginningBal"
      EndProperty
      BeginProperty Field11 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "DebitAmount"
         Caption         =   "DebitAmount"
      EndProperty
      BeginProperty Field12 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "CreditAmount"
         Caption         =   "CreditAmount"
      EndProperty
      BeginProperty Field13 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "TransType"
         Caption         =   "TransType"
      EndProperty
      BeginProperty Field14 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "Fcdebit"
         Caption         =   "Fcdebit"
      EndProperty
      BeginProperty Field15 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "FcCredit"
         Caption         =   "FcCredit"
      EndProperty
      BeginProperty Field16 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "EndingBal"
         Caption         =   "EndingBal"
      EndProperty
      BeginProperty Field17 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   202
         Name            =   "status"
         Caption         =   "status"
      EndProperty
      BeginProperty Field18 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "Prepby"
         Caption         =   "Prepby"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   50
         Scale           =   0
         Size            =   50
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset18 
      CommandName     =   "Command1"
      CommDispId      =   1235
      RsDispId        =   1256
      CommandText     =   "Select * from tempBankjournal  where transdate >= ? and transdate < = ? order by transdate"
      ActiveConnectionName=   "FinanceCON"
      CommandType     =   1
      Grouping        =   -1  'True
      GroupingName    =   "Command1_Grouping"
      SummaryExpanded =   -1  'True
      DetailExpanded  =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   27
      BeginProperty Field1 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   202
         Name            =   "SerialNo"
         Caption         =   "SerialNo"
      EndProperty
      BeginProperty Field2 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "Ticket"
         Caption         =   "Ticket"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   202
         Name            =   "DeleteMark"
         Caption         =   "DeleteMark"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   202
         Name            =   "Classification"
         Caption         =   "Classification"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "AccountNumber"
         Caption         =   "AccountNumber"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   202
         Name            =   "AccountName"
         Caption         =   "AccountName"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "ORNo"
         Caption         =   "ORNo"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   1073741823
         Scale           =   0
         Type            =   203
         Name            =   "Description"
         Caption         =   "Description"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "CheckOwnedBy"
         Caption         =   "CheckOwnedBy"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "CheckNo"
         Caption         =   "CheckNo"
      EndProperty
      BeginProperty Field11 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "PAyee"
         Caption         =   "PAyee"
      EndProperty
      BeginProperty Field12 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "CheckType"
         Caption         =   "CheckType"
      EndProperty
      BeginProperty Field13 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   202
         Name            =   "TransDate"
         Caption         =   "TransDate"
      EndProperty
      BeginProperty Field14 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "DatePosted"
         Caption         =   "DatePosted"
      EndProperty
      BeginProperty Field15 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   202
         Name            =   "CodeType"
         Caption         =   "CodeType"
      EndProperty
      BeginProperty Field16 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "TranType"
         Caption         =   "TranType"
      EndProperty
      BeginProperty Field17 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "BeginningBal"
         Caption         =   "BeginningBal"
      EndProperty
      BeginProperty Field18 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "DebitAmount"
         Caption         =   "DebitAmount"
      EndProperty
      BeginProperty Field19 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "CreditAmount"
         Caption         =   "CreditAmount"
      EndProperty
      BeginProperty Field20 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "EndingBal"
         Caption         =   "EndingBal"
      EndProperty
      BeginProperty Field21 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "Remarks"
         Caption         =   "Remarks"
      EndProperty
      BeginProperty Field22 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "Fcdebit"
         Caption         =   "Fcdebit"
      EndProperty
      BeginProperty Field23 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "FcCredit"
         Caption         =   "FcCredit"
      EndProperty
      BeginProperty Field24 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   202
         Name            =   "status"
         Caption         =   "status"
      EndProperty
      BeginProperty Field25 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "PrepBy"
         Caption         =   "PrepBy"
      EndProperty
      BeginProperty Field26 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "NotedBy"
         Caption         =   "NotedBy"
      EndProperty
      BeginProperty Field27 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "AppBy"
         Caption         =   "AppBy"
      EndProperty
      NumGroups       =   1
      BeginProperty Grouping1 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "TranType"
         Caption         =   "TranType"
      EndProperty
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   20
         Scale           =   0
         Size            =   20
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         Direction       =   1
         Precision       =   20
         Scale           =   0
         Size            =   20
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   2
      BeginProperty Aggregate1 
         Name            =   "Aggregate1"
         AggOn           =   "Command1"
         AggField        =   "DebitAmount"
         AggType         =   2
         AggFunction     =   2
         Precision       =   8
         Size            =   4
         Scale           =   8
         Type            =   131
         Name            =   "Aggregate1"
         Caption         =   "Aggregate1"
         Control         =   "TextBox"
         ControlGuid     =   "{0CA5C786-7C71-11D0-B223-00A0C908FB55}"
      EndProperty
      BeginProperty Aggregate2 
         Name            =   "Aggregate2"
         AggOn           =   "Command1"
         AggField        =   "CreditAmount"
         AggType         =   2
         AggFunction     =   7
         Precision       =   8
         Size            =   4
         Scale           =   8
         Type            =   131
         Name            =   "Aggregate2"
         Caption         =   "Aggregate2"
         Control         =   "TextBox"
         ControlGuid     =   "{0CA5C786-7C71-11D0-B223-00A0C908FB55}"
      EndProperty
   EndProperty
   BeginProperty Recordset19 
      CommandName     =   "BankLEdgerRep"
      CommDispId      =   1247
      RsDispId        =   1251
      CommandText     =   "Select * from tempBankjournal  where transdate >= ? and transdate < = ? order by transdate "
      ActiveConnectionName=   "FinanceCON"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   26
      BeginProperty Field1 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   202
         Name            =   "SerialNo"
         Caption         =   "SerialNo"
      EndProperty
      BeginProperty Field2 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "Ticket"
         Caption         =   "Ticket"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   202
         Name            =   "DeleteMark"
         Caption         =   "DeleteMark"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   202
         Name            =   "Classification"
         Caption         =   "Classification"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "AccountNumber"
         Caption         =   "AccountNumber"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   202
         Name            =   "AccountName"
         Caption         =   "AccountName"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "ORNo"
         Caption         =   "ORNo"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   1073741823
         Scale           =   0
         Type            =   203
         Name            =   "Description"
         Caption         =   "Description"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "CheckOwnedBy"
         Caption         =   "CheckOwnedBy"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "CheckNo"
         Caption         =   "CheckNo"
      EndProperty
      BeginProperty Field11 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "PAyee"
         Caption         =   "PAyee"
      EndProperty
      BeginProperty Field12 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "CheckType"
         Caption         =   "CheckType"
      EndProperty
      BeginProperty Field13 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   202
         Name            =   "TransDate"
         Caption         =   "TransDate"
      EndProperty
      BeginProperty Field14 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "DatePosted"
         Caption         =   "DatePosted"
      EndProperty
      BeginProperty Field15 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "TranType"
         Caption         =   "TranType"
      EndProperty
      BeginProperty Field16 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "BeginningBal"
         Caption         =   "BeginningBal"
      EndProperty
      BeginProperty Field17 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "DebitAmount"
         Caption         =   "DebitAmount"
      EndProperty
      BeginProperty Field18 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "CreditAmount"
         Caption         =   "CreditAmount"
      EndProperty
      BeginProperty Field19 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "EndingBal"
         Caption         =   "EndingBal"
      EndProperty
      BeginProperty Field20 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "Remarks"
         Caption         =   "Remarks"
      EndProperty
      BeginProperty Field21 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "Fcdebit"
         Caption         =   "Fcdebit"
      EndProperty
      BeginProperty Field22 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "FcCredit"
         Caption         =   "FcCredit"
      EndProperty
      BeginProperty Field23 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   202
         Name            =   "status"
         Caption         =   "status"
      EndProperty
      BeginProperty Field24 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "PrepBy"
         Caption         =   "PrepBy"
      EndProperty
      BeginProperty Field25 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "NotedBy"
         Caption         =   "NotedBy"
      EndProperty
      BeginProperty Field26 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "AppBy"
         Caption         =   "AppBy"
      EndProperty
      NumGroups       =   0
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   20
         Scale           =   0
         Size            =   20
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         Direction       =   1
         Precision       =   20
         Scale           =   0
         Size            =   20
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset20 
      CommandName     =   "ByGenDescription"
      CommDispId      =   1257
      RsDispId        =   1285
      CommandText     =   "Select * from Genjournaltrans where transdate = ? and remarks is null order by transdate,ticket "
      ActiveConnectionName=   "FinanceCON"
      CommandType     =   1
      Grouping        =   -1  'True
      GroupingName    =   "ByGenDescription_Grouping"
      IsRSReturning   =   -1  'True
      NumFields       =   21
      BeginProperty Field1 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   202
         Name            =   "SerialNo"
         Caption         =   "SerialNo"
      EndProperty
      BeginProperty Field2 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "Ticket"
         Caption         =   "Ticket"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   202
         Name            =   "DeleteMark"
         Caption         =   "DeleteMark"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   202
         Name            =   "Classification"
         Caption         =   "Classification"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "AccountNumber"
         Caption         =   "AccountNumber"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   202
         Name            =   "AccountName"
         Caption         =   "AccountName"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   1073741823
         Scale           =   0
         Type            =   203
         Name            =   "GenDesc"
         Caption         =   "GenDesc"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   1073741823
         Scale           =   0
         Type            =   203
         Name            =   "Description"
         Caption         =   "Description"
      EndProperty
      BeginProperty Field9 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "TransDate"
         Caption         =   "TransDate"
      EndProperty
      BeginProperty Field10 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "BeginningBal"
         Caption         =   "BeginningBal"
      EndProperty
      BeginProperty Field11 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "DebitAmount"
         Caption         =   "DebitAmount"
      EndProperty
      BeginProperty Field12 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "CreditAmount"
         Caption         =   "CreditAmount"
      EndProperty
      BeginProperty Field13 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "EndingBals"
         Caption         =   "EndingBals"
      EndProperty
      BeginProperty Field14 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "TransType"
         Caption         =   "TransType"
      EndProperty
      BeginProperty Field15 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "Remarks"
         Caption         =   "Remarks"
      EndProperty
      BeginProperty Field16 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "Fcdebit"
         Caption         =   "Fcdebit"
      EndProperty
      BeginProperty Field17 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "FcCredit"
         Caption         =   "FcCredit"
      EndProperty
      BeginProperty Field18 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   202
         Name            =   "status"
         Caption         =   "status"
      EndProperty
      BeginProperty Field19 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "PrepBy"
         Caption         =   "PrepBy"
      EndProperty
      BeginProperty Field20 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "NotedBy"
         Caption         =   "NotedBy"
      EndProperty
      BeginProperty Field21 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "AppBy"
         Caption         =   "AppBy"
      EndProperty
      NumGroups       =   3
      BeginProperty Grouping1 
         Precision       =   0
         Size            =   300
         Scale           =   0
         Type            =   202
         Name            =   "GenDesc"
         Caption         =   "GenDesc"
      EndProperty
      BeginProperty Grouping2 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   202
         Name            =   "SerialNo"
         Caption         =   "SerialNo"
      EndProperty
      BeginProperty Grouping3 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "PrepBy"
         Caption         =   "PrepBy"
      EndProperty
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   20
         DataType        =   202
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset21 
      CommandName     =   "SetupAsset"
      CommDispId      =   1271
      RsDispId        =   1281
      CommandText     =   $"FinanceDE.dsx":0057
      ActiveConnectionName=   "FinanceCON"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   37
      BeginProperty Field1 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "IDNo"
         Caption         =   "IDNo"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "Code"
         Caption         =   "Code"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "Category"
         Caption         =   "Category"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "Sub_Category"
         Caption         =   "Sub_Category"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "NameEng"
         Caption         =   "NameEng"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "NameArab"
         Caption         =   "NameArab"
      EndProperty
      BeginProperty Field7 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "DateRcvd"
         Caption         =   "DateRcvd"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "SerialNo"
         Caption         =   "SerialNo"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "ModelNo"
         Caption         =   "ModelNo"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "GLAcctCode"
         Caption         =   "GLAcctCode"
      EndProperty
      BeginProperty Field11 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "DateRegistered"
         Caption         =   "DateRegistered"
      EndProperty
      BeginProperty Field12 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "ExpenseCode"
         Caption         =   "ExpenseCode"
      EndProperty
      BeginProperty Field13 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "ExpenseName"
         Caption         =   "ExpenseName"
      EndProperty
      BeginProperty Field14 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "AccumulatedCode"
         Caption         =   "AccumulatedCode"
      EndProperty
      BeginProperty Field15 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "AccumulatedNAme"
         Caption         =   "AccumulatedNAme"
      EndProperty
      BeginProperty Field16 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "AssetNo"
         Caption         =   "AssetNo"
      EndProperty
      BeginProperty Field17 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "AssetCode"
         Caption         =   "AssetCode"
      EndProperty
      BeginProperty Field18 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "AssetName"
         Caption         =   "AssetName"
      EndProperty
      BeginProperty Field19 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "ModelNo"
         Caption         =   "ModelNo"
      EndProperty
      BeginProperty Field20 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "SerialNo"
         Caption         =   "SerialNo"
      EndProperty
      BeginProperty Field21 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "InventoryCat"
         Caption         =   "InventoryCat"
      EndProperty
      BeginProperty Field22 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "AssetType"
         Caption         =   "AssetType"
      EndProperty
      BeginProperty Field23 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "AcquisitionDate"
         Caption         =   "AcquisitionDate"
      EndProperty
      BeginProperty Field24 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "AcquisitionValue"
         Caption         =   "AcquisitionValue"
      EndProperty
      BeginProperty Field25 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "SalvageValue"
         Caption         =   "SalvageValue"
      EndProperty
      BeginProperty Field26 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "ComputationMethod"
         Caption         =   "ComputationMethod"
      EndProperty
      BeginProperty Field27 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "UsefullLife"
         Caption         =   "UsefullLife"
      EndProperty
      BeginProperty Field28 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "YearlyFixedRate"
         Caption         =   "YearlyFixedRate"
      EndProperty
      BeginProperty Field29 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "MachineHourUsed"
         Caption         =   "MachineHourUsed"
      EndProperty
      BeginProperty Field30 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "UnitProd"
         Caption         =   "UnitProd"
      EndProperty
      BeginProperty Field31 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "AccumulatedDep"
         Caption         =   "AccumulatedDep"
      EndProperty
      BeginProperty Field32 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "LAstTransDate"
         Caption         =   "LAstTransDate"
      EndProperty
      BeginProperty Field33 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "DebitAcct"
         Caption         =   "DebitAcct"
      EndProperty
      BeginProperty Field34 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "CreditAcct"
         Caption         =   "CreditAcct"
      EndProperty
      BeginProperty Field35 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "AdditionalCost"
         Caption         =   "AdditionalCost"
      EndProperty
      BeginProperty Field36 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   202
         Name            =   "Remarks"
         Caption         =   "Remarks"
      EndProperty
      BeginProperty Field37 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "AssignTo"
         Caption         =   "AssignTo"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
End
Attribute VB_Name = "FinanceDE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

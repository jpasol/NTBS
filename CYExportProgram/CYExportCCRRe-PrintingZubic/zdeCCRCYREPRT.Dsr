VERSION 5.00
Begin {C0E45035-5775-11D0-B388-00A0C9055D8E} zdeCCRCYREPRT 
   ClientHeight    =   11040
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20805
   _ExtentX        =   36698
   _ExtentY        =   19473
   FolderFlags     =   1
   TypeLibGuid     =   "{2D835FC3-0BA4-11D3-BD67-00105A64485A}"
   TypeInfoGuid    =   "{2D835FC4-0BA4-11D3-BD67-00105A64485A}"
   TypeInfoCookie  =   0
   Version         =   4
   NumConnections  =   1
   BeginProperty Connection1 
      ConnectionName  =   "Billing"
      ConnDispId      =   1001
      SourceOfData    =   3
      ConnectionSource=   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=billing;Data Source=sbitcbilling"
      Expanded        =   -1  'True
      IsSQL           =   -1  'True
      QuoteChar       =   34
      SeparatorChar   =   46
   EndProperty
   NumRecordsets   =   11
   BeginProperty Recordset1 
      CommandName     =   "getTotal"
      CommDispId      =   1002
      RsDispId        =   1081
      CommandText     =   $"zdeCCRCYREPRT.dsx":0000
      ActiveConnectionName=   "Billing"
      CommandType     =   1
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   1
      BeginProperty Field1 
         Precision       =   38
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "TotalAmt"
         Caption         =   "TotalAmt"
      EndProperty
      NumGroups       =   0
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   8
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         Direction       =   1
         Precision       =   3
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset2 
      CommandName     =   "getCCRList"
      CommDispId      =   1009
      RsDispId        =   1015
      CommandText     =   $"zdeCCRCYREPRT.dsx":00E6
      ActiveConnectionName=   "Billing"
      CommandType     =   1
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   5
      BeginProperty Field1 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "cusnam"
         Caption         =   "cusnam"
      EndProperty
      BeginProperty Field2 
         Precision       =   8
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "refnum"
         Caption         =   "refnum"
      EndProperty
      BeginProperty Field3 
         Precision       =   3
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "seqnum"
         Caption         =   "seqnum"
      EndProperty
      BeginProperty Field4 
         Precision       =   8
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "ccrnum"
         Caption         =   "ccrnum"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "ccrmod"
         Caption         =   "ccrmod"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   8
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset3 
      CommandName     =   "getCCRdetails"
      CommDispId      =   1016
      RsDispId        =   1082
      CommandText     =   "SELECT * FROM CCRcyx WHERE refnum = ? AND seqnum = ? AND status <> 'CAN' AND CompanyCode IS NOT NULL"
      ActiveConnectionName=   "Billing"
      CommandType     =   1
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   46
      BeginProperty Field1 
         Precision       =   8
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "refnum"
         Caption         =   "refnum"
      EndProperty
      BeginProperty Field2 
         Precision       =   3
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "seqnum"
         Caption         =   "seqnum"
      EndProperty
      BeginProperty Field3 
         Precision       =   1
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "itmnum"
         Caption         =   "itmnum"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   12
         Scale           =   0
         Type            =   129
         Name            =   "cntnum"
         Caption         =   "cntnum"
      EndProperty
      BeginProperty Field5 
         Precision       =   8
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "ccrnum"
         Caption         =   "ccrnum"
      EndProperty
      BeginProperty Field6 
         Precision       =   2
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "cntsze"
         Caption         =   "cntsze"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "fulemp"
         Caption         =   "fulemp"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "dgrcls"
         Caption         =   "dgrcls"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   12
         Scale           =   0
         Type            =   129
         Name            =   "vslcde"
         Caption         =   "vslcde"
      EndProperty
      BeginProperty Field10 
         Precision       =   6
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "whfamt"
         Caption         =   "whfamt"
      EndProperty
      BeginProperty Field11 
         Precision       =   8
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "arramt"
         Caption         =   "arramt"
      EndProperty
      BeginProperty Field12 
         Precision       =   8
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "ovzamt"
         Caption         =   "ovzamt"
      EndProperty
      BeginProperty Field13 
         Precision       =   8
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "dgramt"
         Caption         =   "dgramt"
      EndProperty
      BeginProperty Field14 
         Precision       =   8
         Size            =   19
         Scale           =   3
         Type            =   131
         Name            =   "arrvat"
         Caption         =   "arrvat"
      EndProperty
      BeginProperty Field15 
         Precision       =   8
         Size            =   19
         Scale           =   3
         Type            =   131
         Name            =   "arrtax"
         Caption         =   "arrtax"
      EndProperty
      BeginProperty Field16 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "vatcde"
         Caption         =   "vatcde"
      EndProperty
      BeginProperty Field17 
         Precision       =   6
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "cntovzl"
         Caption         =   "cntovzl"
      EndProperty
      BeginProperty Field18 
         Precision       =   6
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "cntovzw"
         Caption         =   "cntovzw"
      EndProperty
      BeginProperty Field19 
         Precision       =   6
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "cntovzh"
         Caption         =   "cntovzh"
      EndProperty
      BeginProperty Field20 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "ovzums"
         Caption         =   "ovzums"
      EndProperty
      BeginProperty Field21 
         Precision       =   6
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "revton"
         Caption         =   "revton"
      EndProperty
      BeginProperty Field22 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "trncde"
         Caption         =   "trncde"
      EndProperty
      BeginProperty Field23 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "whfcde"
         Caption         =   "whfcde"
      EndProperty
      BeginProperty Field24 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "guarntycde"
         Caption         =   "guarntycde"
      EndProperty
      BeginProperty Field25 
         Precision       =   5
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "dolrte"
         Caption         =   "dolrte"
      EndProperty
      BeginProperty Field26 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "exprtr"
         Caption         =   "exprtr"
      EndProperty
      BeginProperty Field27 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "broker"
         Caption         =   "broker"
      EndProperty
      BeginProperty Field28 
         Precision       =   0
         Size            =   80
         Scale           =   0
         Type            =   200
         Name            =   "entnum"
         Caption         =   "entnum"
      EndProperty
      BeginProperty Field29 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "commod"
         Caption         =   "commod"
      EndProperty
      BeginProperty Field30 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "remark"
         Caption         =   "remark"
      EndProperty
      BeginProperty Field31 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "trknam"
         Caption         =   "trknam"
      EndProperty
      BeginProperty Field32 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "pltnum"
         Caption         =   "pltnum"
      EndProperty
      BeginProperty Field33 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "trkchs"
         Caption         =   "trkchs"
      EndProperty
      BeginProperty Field34 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   129
         Name            =   "status"
         Caption         =   "status"
      EndProperty
      BeginProperty Field35 
         Precision       =   8
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "ovrccr"
         Caption         =   "ovrccr"
      EndProperty
      BeginProperty Field36 
         Precision       =   8
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "ppanum"
         Caption         =   "ppanum"
      EndProperty
      BeginProperty Field37 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   129
         Name            =   "userid"
         Caption         =   "userid"
      EndProperty
      BeginProperty Field38 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "sysdttm"
         Caption         =   "sysdttm"
      EndProperty
      BeginProperty Field39 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "updcde"
         Caption         =   "updcde"
      EndProperty
      BeginProperty Field40 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "outdttm"
         Caption         =   "outdttm"
      EndProperty
      BeginProperty Field41 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   129
         Name            =   "supvsr"
         Caption         =   "supvsr"
      EndProperty
      BeginProperty Field42 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "IsN4BillingPermissionGranted"
         Caption         =   "IsN4BillingPermissionGranted"
      EndProperty
      BeginProperty Field43 
         Precision       =   8
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "wghamt"
         Caption         =   "wghamt"
      EndProperty
      BeginProperty Field44 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "IsN4BillingDGPermissionGranted"
         Caption         =   "IsN4BillingDGPermissionGranted"
      EndProperty
      BeginProperty Field45 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "IsN4BillingOOGPermissionGranted"
         Caption         =   "IsN4BillingOOGPermissionGranted"
      EndProperty
      BeginProperty Field46 
         Precision       =   0
         Size            =   5
         Scale           =   0
         Type            =   202
         Name            =   "CompanyCode"
         Caption         =   "CompanyCode"
      EndProperty
      NumGroups       =   0
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   8
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         Direction       =   1
         Precision       =   3
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset4 
      CommandName     =   "getAdrAmt"
      CommDispId      =   1022
      RsDispId        =   1027
      CommandText     =   "select * from CCRpay where refnum = ?"
      ActiveConnectionName=   "Billing"
      CommandType     =   1
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   13
      BeginProperty Field1 
         Precision       =   8
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "refnum"
         Caption         =   "refnum"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   6
         Scale           =   0
         Type            =   129
         Name            =   "cuscde"
         Caption         =   "cuscde"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "cusnam"
         Caption         =   "cusnam"
      EndProperty
      BeginProperty Field4 
         Precision       =   10
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "cshamt"
         Caption         =   "cshamt"
      EndProperty
      BeginProperty Field5 
         Precision       =   10
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "adramt"
         Caption         =   "adramt"
      EndProperty
      BeginProperty Field6 
         Precision       =   8
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "adrnum"
         Caption         =   "adrnum"
      EndProperty
      BeginProperty Field7 
         Precision       =   9
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "chgamt"
         Caption         =   "chgamt"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "status"
         Caption         =   "status"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "rectag"
         Caption         =   "rectag"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   129
         Name            =   "userid"
         Caption         =   "userid"
      EndProperty
      BeginProperty Field11 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "sysdttm"
         Caption         =   "sysdttm"
      EndProperty
      BeginProperty Field12 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "updcde"
         Caption         =   "updcde"
      EndProperty
      BeginProperty Field13 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "ccrtyp"
         Caption         =   "ccrtyp"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "Reference"
         Direction       =   1
         Precision       =   8
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset5 
      CommandName     =   "CCRCyx"
      CommDispId      =   1030
      RsDispId        =   1035
      CommandText     =   "select * from ccrcyx"
      ActiveConnectionName=   "Billing"
      CommandType     =   1
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   39
      BeginProperty Field1 
         Precision       =   8
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "refnum"
         Caption         =   "refnum"
      EndProperty
      BeginProperty Field2 
         Precision       =   3
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "seqnum"
         Caption         =   "seqnum"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   12
         Scale           =   0
         Type            =   129
         Name            =   "cntnum"
         Caption         =   "cntnum"
      EndProperty
      BeginProperty Field4 
         Precision       =   8
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "ccrnum"
         Caption         =   "ccrnum"
      EndProperty
      BeginProperty Field5 
         Precision       =   2
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "cntsze"
         Caption         =   "cntsze"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "fulemp"
         Caption         =   "fulemp"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "dgrcls"
         Caption         =   "dgrcls"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   7
         Scale           =   0
         Type            =   129
         Name            =   "vslcde"
         Caption         =   "vslcde"
      EndProperty
      BeginProperty Field9 
         Precision       =   6
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "whfamt"
         Caption         =   "whfamt"
      EndProperty
      BeginProperty Field10 
         Precision       =   8
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "arramt"
         Caption         =   "arramt"
      EndProperty
      BeginProperty Field11 
         Precision       =   8
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "ovzamt"
         Caption         =   "ovzamt"
      EndProperty
      BeginProperty Field12 
         Precision       =   8
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "dgramt"
         Caption         =   "dgramt"
      EndProperty
      BeginProperty Field13 
         Precision       =   8
         Size            =   19
         Scale           =   3
         Type            =   131
         Name            =   "arrvat"
         Caption         =   "arrvat"
      EndProperty
      BeginProperty Field14 
         Precision       =   8
         Size            =   19
         Scale           =   3
         Type            =   131
         Name            =   "arrtax"
         Caption         =   "arrtax"
      EndProperty
      BeginProperty Field15 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "vatcde"
         Caption         =   "vatcde"
      EndProperty
      BeginProperty Field16 
         Precision       =   6
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "cntovzl"
         Caption         =   "cntovzl"
      EndProperty
      BeginProperty Field17 
         Precision       =   6
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "cntovzw"
         Caption         =   "cntovzw"
      EndProperty
      BeginProperty Field18 
         Precision       =   6
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "cntovzh"
         Caption         =   "cntovzh"
      EndProperty
      BeginProperty Field19 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "ovzums"
         Caption         =   "ovzums"
      EndProperty
      BeginProperty Field20 
         Precision       =   6
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "revton"
         Caption         =   "revton"
      EndProperty
      BeginProperty Field21 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "trncde"
         Caption         =   "trncde"
      EndProperty
      BeginProperty Field22 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "whfcde"
         Caption         =   "whfcde"
      EndProperty
      BeginProperty Field23 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "guarntycde"
         Caption         =   "guarntycde"
      EndProperty
      BeginProperty Field24 
         Precision       =   5
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "dolrte"
         Caption         =   "dolrte"
      EndProperty
      BeginProperty Field25 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "exprtr"
         Caption         =   "exprtr"
      EndProperty
      BeginProperty Field26 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "broker"
         Caption         =   "broker"
      EndProperty
      BeginProperty Field27 
         Precision       =   0
         Size            =   80
         Scale           =   0
         Type            =   200
         Name            =   "entnum"
         Caption         =   "entnum"
      EndProperty
      BeginProperty Field28 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "commod"
         Caption         =   "commod"
      EndProperty
      BeginProperty Field29 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "remark"
         Caption         =   "remark"
      EndProperty
      BeginProperty Field30 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "trknam"
         Caption         =   "trknam"
      EndProperty
      BeginProperty Field31 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "pltnum"
         Caption         =   "pltnum"
      EndProperty
      BeginProperty Field32 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "trkchs"
         Caption         =   "trkchs"
      EndProperty
      BeginProperty Field33 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   129
         Name            =   "status"
         Caption         =   "status"
      EndProperty
      BeginProperty Field34 
         Precision       =   8
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "ovrccr"
         Caption         =   "ovrccr"
      EndProperty
      BeginProperty Field35 
         Precision       =   8
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "ppanum"
         Caption         =   "ppanum"
      EndProperty
      BeginProperty Field36 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   129
         Name            =   "userid"
         Caption         =   "userid"
      EndProperty
      BeginProperty Field37 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "sysdttm"
         Caption         =   "sysdttm"
      EndProperty
      BeginProperty Field38 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "updcde"
         Caption         =   "updcde"
      EndProperty
      BeginProperty Field39 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "outdttm"
         Caption         =   "outdttm"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset6 
      CommandName     =   "CCRPay"
      CommDispId      =   1036
      RsDispId        =   1041
      CommandText     =   "select * from ccrpay"
      ActiveConnectionName=   "Billing"
      CommandType     =   1
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   28
      BeginProperty Field1 
         Precision       =   8
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "refnum"
         Caption         =   "refnum"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   6
         Scale           =   0
         Type            =   129
         Name            =   "cuscde"
         Caption         =   "cuscde"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "cusnam"
         Caption         =   "cusnam"
      EndProperty
      BeginProperty Field4 
         Precision       =   10
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "cshamt"
         Caption         =   "cshamt"
      EndProperty
      BeginProperty Field5 
         Precision       =   10
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "adramt"
         Caption         =   "adramt"
      EndProperty
      BeginProperty Field6 
         Precision       =   8
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "adrnum"
         Caption         =   "adrnum"
      EndProperty
      BeginProperty Field7 
         Precision       =   9
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "chgamt"
         Caption         =   "chgamt"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "chkno1"
         Caption         =   "chkno1"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "chkno2"
         Caption         =   "chkno2"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "chkno3"
         Caption         =   "chkno3"
      EndProperty
      BeginProperty Field11 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "chkno4"
         Caption         =   "chkno4"
      EndProperty
      BeginProperty Field12 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "chkno5"
         Caption         =   "chkno5"
      EndProperty
      BeginProperty Field13 
         Precision       =   10
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "chkamt1"
         Caption         =   "chkamt1"
      EndProperty
      BeginProperty Field14 
         Precision       =   10
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "chkamt2"
         Caption         =   "chkamt2"
      EndProperty
      BeginProperty Field15 
         Precision       =   10
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "chkamt3"
         Caption         =   "chkamt3"
      EndProperty
      BeginProperty Field16 
         Precision       =   10
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "chkamt4"
         Caption         =   "chkamt4"
      EndProperty
      BeginProperty Field17 
         Precision       =   10
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "chkamt5"
         Caption         =   "chkamt5"
      EndProperty
      BeginProperty Field18 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "chkbnk1"
         Caption         =   "chkbnk1"
      EndProperty
      BeginProperty Field19 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "chkbnk2"
         Caption         =   "chkbnk2"
      EndProperty
      BeginProperty Field20 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "chkbnk3"
         Caption         =   "chkbnk3"
      EndProperty
      BeginProperty Field21 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "chkbnk4"
         Caption         =   "chkbnk4"
      EndProperty
      BeginProperty Field22 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "chkbnk5"
         Caption         =   "chkbnk5"
      EndProperty
      BeginProperty Field23 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   129
         Name            =   "status"
         Caption         =   "status"
      EndProperty
      BeginProperty Field24 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "rectag"
         Caption         =   "rectag"
      EndProperty
      BeginProperty Field25 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   129
         Name            =   "userid"
         Caption         =   "userid"
      EndProperty
      BeginProperty Field26 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "sysdttm"
         Caption         =   "sysdttm"
      EndProperty
      BeginProperty Field27 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "updcde"
         Caption         =   "updcde"
      EndProperty
      BeginProperty Field28 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "ccrtyp"
         Caption         =   "ccrtyp"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset7 
      CommandName     =   "CountCCR"
      CommDispId      =   1042
      RsDispId        =   1048
      CommandText     =   "SELECT noccr = COUNT(DISTINCT ccrnum) FROM ccrcyx WHERE refnum = ? AND seqnum = ?"
      ActiveConnectionName=   "Billing"
      CommandType     =   1
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   1
      BeginProperty Field1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "noccr"
         Caption         =   "noccr"
      EndProperty
      NumGroups       =   0
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "pRefnum"
         Direction       =   1
         Precision       =   8
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "pSeqnum"
         Direction       =   1
         Precision       =   3
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset8 
      CommandName     =   "CCRDetails"
      CommDispId      =   1049
      RsDispId        =   1054
      CommandText     =   "SELECT * FROM ccrcyx WHERE refnum = ? AND seqnum = ? ORDER BY refnum,seqnum,itmnum "
      ActiveConnectionName=   "Billing"
      CommandType     =   1
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   40
      BeginProperty Field1 
         Precision       =   8
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "refnum"
         Caption         =   "refnum"
      EndProperty
      BeginProperty Field2 
         Precision       =   3
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "seqnum"
         Caption         =   "seqnum"
      EndProperty
      BeginProperty Field3 
         Precision       =   1
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "itmnum"
         Caption         =   "itmnum"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   12
         Scale           =   0
         Type            =   129
         Name            =   "cntnum"
         Caption         =   "cntnum"
      EndProperty
      BeginProperty Field5 
         Precision       =   8
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "ccrnum"
         Caption         =   "ccrnum"
      EndProperty
      BeginProperty Field6 
         Precision       =   2
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "cntsze"
         Caption         =   "cntsze"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "fulemp"
         Caption         =   "fulemp"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "dgrcls"
         Caption         =   "dgrcls"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   7
         Scale           =   0
         Type            =   129
         Name            =   "vslcde"
         Caption         =   "vslcde"
      EndProperty
      BeginProperty Field10 
         Precision       =   6
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "whfamt"
         Caption         =   "whfamt"
      EndProperty
      BeginProperty Field11 
         Precision       =   8
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "arramt"
         Caption         =   "arramt"
      EndProperty
      BeginProperty Field12 
         Precision       =   8
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "ovzamt"
         Caption         =   "ovzamt"
      EndProperty
      BeginProperty Field13 
         Precision       =   8
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "dgramt"
         Caption         =   "dgramt"
      EndProperty
      BeginProperty Field14 
         Precision       =   8
         Size            =   19
         Scale           =   3
         Type            =   131
         Name            =   "arrvat"
         Caption         =   "arrvat"
      EndProperty
      BeginProperty Field15 
         Precision       =   8
         Size            =   19
         Scale           =   3
         Type            =   131
         Name            =   "arrtax"
         Caption         =   "arrtax"
      EndProperty
      BeginProperty Field16 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "vatcde"
         Caption         =   "vatcde"
      EndProperty
      BeginProperty Field17 
         Precision       =   6
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "cntovzl"
         Caption         =   "cntovzl"
      EndProperty
      BeginProperty Field18 
         Precision       =   6
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "cntovzw"
         Caption         =   "cntovzw"
      EndProperty
      BeginProperty Field19 
         Precision       =   6
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "cntovzh"
         Caption         =   "cntovzh"
      EndProperty
      BeginProperty Field20 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "ovzums"
         Caption         =   "ovzums"
      EndProperty
      BeginProperty Field21 
         Precision       =   6
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "revton"
         Caption         =   "revton"
      EndProperty
      BeginProperty Field22 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "trncde"
         Caption         =   "trncde"
      EndProperty
      BeginProperty Field23 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "whfcde"
         Caption         =   "whfcde"
      EndProperty
      BeginProperty Field24 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "guarntycde"
         Caption         =   "guarntycde"
      EndProperty
      BeginProperty Field25 
         Precision       =   5
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "dolrte"
         Caption         =   "dolrte"
      EndProperty
      BeginProperty Field26 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "exprtr"
         Caption         =   "exprtr"
      EndProperty
      BeginProperty Field27 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "broker"
         Caption         =   "broker"
      EndProperty
      BeginProperty Field28 
         Precision       =   0
         Size            =   80
         Scale           =   0
         Type            =   200
         Name            =   "entnum"
         Caption         =   "entnum"
      EndProperty
      BeginProperty Field29 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "commod"
         Caption         =   "commod"
      EndProperty
      BeginProperty Field30 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "remark"
         Caption         =   "remark"
      EndProperty
      BeginProperty Field31 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "trknam"
         Caption         =   "trknam"
      EndProperty
      BeginProperty Field32 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "pltnum"
         Caption         =   "pltnum"
      EndProperty
      BeginProperty Field33 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "trkchs"
         Caption         =   "trkchs"
      EndProperty
      BeginProperty Field34 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   129
         Name            =   "status"
         Caption         =   "status"
      EndProperty
      BeginProperty Field35 
         Precision       =   8
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "ovrccr"
         Caption         =   "ovrccr"
      EndProperty
      BeginProperty Field36 
         Precision       =   8
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "ppanum"
         Caption         =   "ppanum"
      EndProperty
      BeginProperty Field37 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   129
         Name            =   "userid"
         Caption         =   "userid"
      EndProperty
      BeginProperty Field38 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "sysdttm"
         Caption         =   "sysdttm"
      EndProperty
      BeginProperty Field39 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "updcde"
         Caption         =   "updcde"
      EndProperty
      BeginProperty Field40 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "outdttm"
         Caption         =   "outdttm"
      EndProperty
      NumGroups       =   0
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   8
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         Direction       =   1
         Precision       =   3
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset9 
      CommandName     =   "GetUserInfo"
      CommDispId      =   1055
      RsDispId        =   1059
      CommandText     =   "select workstation=host_name()"
      ActiveConnectionName=   "Billing"
      CommandType     =   1
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   1
      BeginProperty Field1 
         Precision       =   0
         Size            =   128
         Scale           =   0
         Type            =   202
         Name            =   "workstation"
         Caption         =   "workstation"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset10 
      CommandName     =   "GetDetails"
      CommDispId      =   1060
      RsDispId        =   1084
      CommandText     =   "select * from ccrcyx where ccrnum = ? and CompanyCode is not null order by itmnum"
      ActiveConnectionName=   "Billing"
      CommandType     =   1
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   46
      BeginProperty Field1 
         Precision       =   8
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "refnum"
         Caption         =   "refnum"
      EndProperty
      BeginProperty Field2 
         Precision       =   3
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "seqnum"
         Caption         =   "seqnum"
      EndProperty
      BeginProperty Field3 
         Precision       =   1
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "itmnum"
         Caption         =   "itmnum"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   12
         Scale           =   0
         Type            =   129
         Name            =   "cntnum"
         Caption         =   "cntnum"
      EndProperty
      BeginProperty Field5 
         Precision       =   8
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "ccrnum"
         Caption         =   "ccrnum"
      EndProperty
      BeginProperty Field6 
         Precision       =   2
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "cntsze"
         Caption         =   "cntsze"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "fulemp"
         Caption         =   "fulemp"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "dgrcls"
         Caption         =   "dgrcls"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   12
         Scale           =   0
         Type            =   129
         Name            =   "vslcde"
         Caption         =   "vslcde"
      EndProperty
      BeginProperty Field10 
         Precision       =   6
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "whfamt"
         Caption         =   "whfamt"
      EndProperty
      BeginProperty Field11 
         Precision       =   8
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "arramt"
         Caption         =   "arramt"
      EndProperty
      BeginProperty Field12 
         Precision       =   8
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "ovzamt"
         Caption         =   "ovzamt"
      EndProperty
      BeginProperty Field13 
         Precision       =   8
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "dgramt"
         Caption         =   "dgramt"
      EndProperty
      BeginProperty Field14 
         Precision       =   8
         Size            =   19
         Scale           =   3
         Type            =   131
         Name            =   "arrvat"
         Caption         =   "arrvat"
      EndProperty
      BeginProperty Field15 
         Precision       =   8
         Size            =   19
         Scale           =   3
         Type            =   131
         Name            =   "arrtax"
         Caption         =   "arrtax"
      EndProperty
      BeginProperty Field16 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "vatcde"
         Caption         =   "vatcde"
      EndProperty
      BeginProperty Field17 
         Precision       =   6
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "cntovzl"
         Caption         =   "cntovzl"
      EndProperty
      BeginProperty Field18 
         Precision       =   6
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "cntovzw"
         Caption         =   "cntovzw"
      EndProperty
      BeginProperty Field19 
         Precision       =   6
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "cntovzh"
         Caption         =   "cntovzh"
      EndProperty
      BeginProperty Field20 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "ovzums"
         Caption         =   "ovzums"
      EndProperty
      BeginProperty Field21 
         Precision       =   6
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "revton"
         Caption         =   "revton"
      EndProperty
      BeginProperty Field22 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "trncde"
         Caption         =   "trncde"
      EndProperty
      BeginProperty Field23 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "whfcde"
         Caption         =   "whfcde"
      EndProperty
      BeginProperty Field24 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "guarntycde"
         Caption         =   "guarntycde"
      EndProperty
      BeginProperty Field25 
         Precision       =   5
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "dolrte"
         Caption         =   "dolrte"
      EndProperty
      BeginProperty Field26 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "exprtr"
         Caption         =   "exprtr"
      EndProperty
      BeginProperty Field27 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "broker"
         Caption         =   "broker"
      EndProperty
      BeginProperty Field28 
         Precision       =   0
         Size            =   80
         Scale           =   0
         Type            =   200
         Name            =   "entnum"
         Caption         =   "entnum"
      EndProperty
      BeginProperty Field29 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "commod"
         Caption         =   "commod"
      EndProperty
      BeginProperty Field30 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "remark"
         Caption         =   "remark"
      EndProperty
      BeginProperty Field31 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "trknam"
         Caption         =   "trknam"
      EndProperty
      BeginProperty Field32 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "pltnum"
         Caption         =   "pltnum"
      EndProperty
      BeginProperty Field33 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "trkchs"
         Caption         =   "trkchs"
      EndProperty
      BeginProperty Field34 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   129
         Name            =   "status"
         Caption         =   "status"
      EndProperty
      BeginProperty Field35 
         Precision       =   8
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "ovrccr"
         Caption         =   "ovrccr"
      EndProperty
      BeginProperty Field36 
         Precision       =   8
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "ppanum"
         Caption         =   "ppanum"
      EndProperty
      BeginProperty Field37 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   129
         Name            =   "userid"
         Caption         =   "userid"
      EndProperty
      BeginProperty Field38 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "sysdttm"
         Caption         =   "sysdttm"
      EndProperty
      BeginProperty Field39 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "updcde"
         Caption         =   "updcde"
      EndProperty
      BeginProperty Field40 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "outdttm"
         Caption         =   "outdttm"
      EndProperty
      BeginProperty Field41 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   129
         Name            =   "supvsr"
         Caption         =   "supvsr"
      EndProperty
      BeginProperty Field42 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "IsN4BillingPermissionGranted"
         Caption         =   "IsN4BillingPermissionGranted"
      EndProperty
      BeginProperty Field43 
         Precision       =   8
         Size            =   19
         Scale           =   2
         Type            =   131
         Name            =   "wghamt"
         Caption         =   "wghamt"
      EndProperty
      BeginProperty Field44 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "IsN4BillingDGPermissionGranted"
         Caption         =   "IsN4BillingDGPermissionGranted"
      EndProperty
      BeginProperty Field45 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "IsN4BillingOOGPermissionGranted"
         Caption         =   "IsN4BillingOOGPermissionGranted"
      EndProperty
      BeginProperty Field46 
         Precision       =   0
         Size            =   5
         Scale           =   0
         Type            =   202
         Name            =   "CompanyCode"
         Caption         =   "CompanyCode"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   8
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset11 
      CommandName     =   "GetRefDetails"
      CommDispId      =   1066
      RsDispId        =   1083
      CommandText     =   $"zdeCCRCYREPRT.dsx":01C3
      ActiveConnectionName=   "Billing"
      CommandType     =   1
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   6
      BeginProperty Field1 
         Precision       =   8
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "refnum"
         Caption         =   "refnum"
      EndProperty
      BeginProperty Field2 
         Precision       =   3
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "seqnum"
         Caption         =   "seqnum"
      EndProperty
      BeginProperty Field3 
         Precision       =   8
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "ccrnum"
         Caption         =   "ccrnum"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "exprtr"
         Caption         =   "exprtr"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "broker"
         Caption         =   "broker"
      EndProperty
      BeginProperty Field6 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "sysdttm"
         Caption         =   "sysdttm"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   8
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
End
Attribute VB_Name = "zdeCCRCYREPRT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

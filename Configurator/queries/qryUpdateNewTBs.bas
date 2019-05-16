Operation =4
Option =0
Begin InputTables
    Name ="tblNewPrductConstruction"
End
Begin OutputColumns
    Name ="tblNewPrductConstruction.[New Oracle Part #]"
    Expression ="Left(Forms!frmCreateItem![NewOracle#],7) & Mid([New Oracle Part #],8,5)"
    Name ="tblNewPrductConstruction.[Item No]"
    Expression ="Left(Forms!frmCreateItem!ItemNo,9) & Mid([Item No],10,2) & Mid(Forms!frmCreateIt"
        "em!ItemNo,12,15)"
    Name ="tblNewPrductConstruction.[Revision Date]"
    Expression ="Date()"
    Name ="tblNewPrductConstruction.[Revision Letter]"
    Expression ="\"1\""
    Name ="tblNewPrductConstruction.[Revision History]"
    Expression ="Date() & \" - INITIAL ISSUE \" & Forms!frmSwitchBoard!Initials"
    Name ="tblNewPrductConstruction.OracleStatus"
    Expression ="\"Active\""
    Name ="tblNewPrductConstruction.[Jacket Material]"
    Expression ="Forms!frmCreateItem!Jacket"
    Name ="tblNewPrductConstruction.[Fiber Type]"
    Expression ="IIf(Forms!frmCreateItem!MixedFiber=1,\"MIXED (\" & Forms!frmCreateItem!Fiber1 & "
        "\"/\" & Forms!frmCreateItem!Fiber2 & \")\",Forms!frmFiberSpec!FiberType)"
    Name ="tblNewPrductConstruction.[Fiber Oracle item]"
    Expression ="Forms!frmCreateItem!Fiber1"
    Name ="tblNewPrductConstruction.FiberType2"
    Expression ="IIf(Forms!frmCreateItem!MixedFiber=1,Forms!frmCreateItem!Fiber2,\"NONE\")"
    Name ="tblNewPrductConstruction.[1st Req Freq]"
    Expression ="Forms!frmFiberSpec![1stReqFreq]"
    Name ="tblNewPrductConstruction.[1st Max Atten]"
    Expression ="Forms!frmFiberSpec![1stMaxAtten]"
    Name ="tblNewPrductConstruction.[1st Min BandW]"
    Expression ="Forms!frmFiberSpec![1stMinBandW]"
    Name ="tblNewPrductConstruction.[2nd Req Freq]"
    Expression ="Forms!frmFiberSpec![2ndReqFreq]"
    Name ="tblNewPrductConstruction.[2nd Max Atten]"
    Expression ="Forms!frmFiberSpec![2ndMaxAtten]"
    Name ="tblNewPrductConstruction.[2nd Min BandW]"
    Expression ="Forms!frmFiberSpec![2ndMinBandW]"
    Name ="tblNewPrductConstruction.[SM 1300 Max Atten]"
    Expression ="\"\""
    Name ="tblNewPrductConstruction.[SM 1550 Max Atten]"
    Expression ="\"\""
    Name ="tblNewPrductConstruction.Active"
    Expression ="-1"
    Name ="tblNewPrductConstruction.Reason"
    Expression ="Null"
    Name ="tblNewPrductConstruction.[zzOracle Part#]"
    Expression ="Null"
    Name ="tblNewPrductConstruction.[Listing Company]"
    Expression ="Forms!frmCreateItem!ListingCompany"
    Name ="tblNewPrductConstruction.[Label Type]"
    Expression ="IIf(Forms!frmCreateItem!ListingCompany<>\"NONE\",\"(\" & Forms!frmCreateItem!Lis"
        "tingCompany & \")c(\" & Forms!frmCreateItem!ListingCompany & \")\",\"\")"
    Name ="tblNewPrductConstruction.CustomerRev"
    Expression ="Forms!frmCreateItem!CustomerRev"
    Name ="tblNewPrductConstruction.FiberType3"
    Expression ="\"NONE\""
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0xc389f6d006ab474bb47d9e008b0f7b21
End
Begin
    Begin
        dbText "Name" ="tblNewPrductConstruction.[Fiber Type]"
        dbInteger "ColumnWidth" ="3870"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.NewPrintLine4"
        dbInteger "ColumnWidth" ="1980"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.[Revision Date]"
        dbInteger "ColumnWidth" ="2505"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.[New Oracle Part #]"
        dbInteger "ColumnWidth" ="3510"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.[Item No]"
        dbInteger "ColumnWidth" ="2595"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.[Revision History]"
        dbInteger "ColumnWidth" ="6165"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =98
    Top =297
    Right =1207
    Bottom =884
    Left =-1
    Top =-1
    Right =1077
    Bottom =356
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =92
        Top =91
        Right =552
        Bottom =329
        Top =0
        Name ="tblNewPrductConstruction"
        Name =""
    End
End

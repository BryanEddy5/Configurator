Operation =4
Option =0
Begin InputTables
    Name ="tblNewPrductConstruction"
End
Begin OutputColumns
    Name ="tblNewPrductConstruction.[New Oracle Part #]"
    Expression ="Forms!frmCreateItem![NewOracle#]"
    Name ="tblNewPrductConstruction.[Item No]"
    Expression ="Forms!frmCreateItem!ItemNo"
    Name ="tblNewPrductConstruction.Customer"
    Expression ="Forms!frmCreateItem!Customer"
    Name ="tblNewPrductConstruction.[Customer Part#]"
    Expression ="Forms!frmCreateItem!PID"
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
    Name ="tblNewPrductConstruction.[Jacket Color]"
    Expression ="Forms!frmCreateItem!Color"
    Name ="tblNewPrductConstruction.[Print Type (base)]"
    Expression ="IIf(Forms!frmCreateItem!StripePrint=1,\"SPECIAL: One (1) BLACK STRIPE\",IIf(Form"
        "s!frmCreateItem!PrintWhite=1,\"SPECIAL: WHITE PRINT\",\"Standard black print\"))"
    Name ="tblNewPrductConstruction.[Print Reel No]"
    Expression ="IIf(Forms!frmCreateItem!PrintType=0,0,-1)"
    Name ="tblNewPrductConstruction.[Print Item No]"
    Expression ="IIf(Forms!frmCreateItem!PrintType=1 Or Forms!frmCreateItem![NewOracle#] Like \"p"
        "a*\",-1,0)"
    Name ="tblNewPrductConstruction.[Print Line 1]"
    Expression ="Forms!frmCreateItem!PrintLine1"
    Name ="tblNewPrductConstruction.[Print Line 2]"
    Expression ="Forms!frmCreateItem!PrintLine2"
    Name ="tblNewPrductConstruction.[Print Line 3]"
    Expression ="Forms!frmCreateItem!PrintLine3"
    Name ="tblNewPrductConstruction.[Print Line 4]"
    Expression ="Forms!frmCreateItem!PrintLine4"
    Name ="tblNewPrductConstruction.NewPrintLine4"
    Expression ="IIf(Forms!frmCreateItem!PrintLine4 Like \"*meter*\",\"[DATE2] [METER] METERS\",\""
        "[DATE2] [2FT] FEET\")"
    Name ="tblNewPrductConstruction.[Print Spacing]"
    Expression ="IIf([Forms]![frmCreateItem]![PrintLine4] Like \"*meter*\",\"METER\",\"2FT\")"
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
    Expression ="0"
    Name ="tblNewPrductConstruction.Reason"
    Expression ="\"IP \" & Forms!frmSwitchBoard!Initials"
    Name ="tblNewPrductConstruction.ENumber"
    Expression ="Forms!frmCreateItem!ListingNumber"
    Name ="tblNewPrductConstruction.[zzOracle Part#]"
    Expression ="Null"
    Name ="tblNewPrductConstruction.[UL-ETL-Section]"
    Expression ="Forms!frmCreateItem!ListingSection"
    Name ="tblNewPrductConstruction.[UL-ETL-Const]"
    Expression ="Forms!frmCreateItem!ListingConstruction"
    Name ="tblNewPrductConstruction.[UL-ETL-Listing]"
    Expression ="Forms!frmCreateItem!Listing"
    Name ="tblNewPrductConstruction.[Listing Company]"
    Expression ="[Forms]![frmCreateItem]![ListingCompany]"
    Name ="tblNewPrductConstruction.[Label Type]"
    Expression ="getNewItemListingLabel()"
    Name ="tblNewPrductConstruction.CustomerRev"
    Expression ="Forms!frmCreateItem!CustomerRev"
    Name ="tblNewPrductConstruction.FiberType3"
    Expression ="\"NONE\""
    Name ="tblNewPrductConstruction.[Special Instr Product1]"
    Expression ="\"NONE\""
    Name ="tblNewPrductConstruction.[zzSpecial Instr Product3]"
    Expression ="\"NONE\""
    Name ="tblNewPrductConstruction.[zzSpecial Instr Product4]"
    Expression ="\"NONE\""
    Name ="tblNewPrductConstruction.Base"
    Expression ="[Forms]![frmCreateItem]![Base]"
    Name ="tblNewPrductConstruction.[TB Color Series]"
    Expression ="\"NONE\""
    Name ="tblNewPrductConstruction.[Color Chip ID]"
    Expression ="Null"
    Name ="tblNewPrductConstruction.VendorSpecificFiber"
    Expression ="GetisVendorSpecificPrint()"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x964e45acfbb42547a2d3a900e4a3a0ca
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
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.[Special Instr Product1]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.[zzSpecial Instr Product3]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.[zzSpecial Instr Product4]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.ENumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.FiberType3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.[Print Spacing]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.Base"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.[TB Color Series]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.[Color Chip ID]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.[Print Type (base)]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.CustomerRev"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.[Label Type]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.[Listing Company]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.Colored_Fiber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.VendorSpecificFiber"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =142
    Top =111
    Right =1251
    Bottom =740
    Left =-1
    Top =-1
    Right =1077
    Bottom =469
    Left =0
    Top =34
    ColumnsShown =579
    Begin
        Left =92
        Top =57
        Right =552
        Bottom =295
        Top =0
        Name ="tblNewPrductConstruction"
        Name =""
    End
End

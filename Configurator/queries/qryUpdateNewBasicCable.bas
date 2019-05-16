Operation =4
Option =0
Begin InputTables
    Name ="tblNewBasicCable_New"
End
Begin OutputColumns
    Name ="tblNewBasicCable_New.[New Oracle Part #]"
    Expression ="Forms!frmCreateItem![NewOracle#]"
    Name ="tblNewBasicCable_New.[Item No]"
    Expression ="Forms!frmCreateItem!ItemNo"
    Name ="tblNewBasicCable_New.Customer"
    Expression ="Forms!frmCreateItem!Customer"
    Name ="tblNewBasicCable_New.[Customer Part#]"
    Expression ="Forms!frmCreateItem!PID"
    Name ="tblNewBasicCable_New.[Revision Date]"
    Expression ="Date()"
    Name ="tblNewBasicCable_New.[Revision Letter]"
    Expression ="1"
    Name ="tblNewBasicCable_New.[Revision History]"
    Expression ="Date() & \" - INITIAL ISSUE \" & Forms!frmSwitchBoard!Initials"
    Name ="tblNewBasicCable_New.OracleStatus"
    Expression ="\"Active\""
    Name ="tblNewBasicCable_New.[Jacket Color]"
    Expression ="Forms!frmCreateItem!Color"
    Name ="tblNewBasicCable_New.[Print Type (base)]"
    Expression ="IIf(Forms!frmCreateItem!PrintWhite=1,\"SPECIAL: WHITE PRINT\",\"Standard black p"
        "rint\")"
    Name ="tblNewBasicCable_New.[Print Reel No]"
    Expression ="CreateNewItemIsPrinted()"
    Name ="tblNewBasicCable_New.[Print Item No]"
    Expression ="IIf(Forms!frmCreateItem!PrintType=1,-1,0)"
    Name ="tblNewBasicCable_New.[Print Line 1]"
    Expression ="Forms!frmCreateItem!PrintLine1"
    Name ="tblNewBasicCable_New.[Print Line 2]"
    Expression ="Forms!frmCreateItem!PrintLine2"
    Name ="tblNewBasicCable_New.[Print Line 3]"
    Expression ="Forms!frmCreateItem!PrintLine3"
    Name ="tblNewBasicCable_New.[Print Line 4]"
    Expression ="Forms!frmCreateItem!PrintLine4"
    Name ="tblNewBasicCable_New.NewPrintLine4"
    Expression ="IIf(Forms!frmCreateItem!PrintLine4 Like \"*meter*\",\"[DATE2] [METER] METERS\",\""
        "[DATE2] [2FT] FEET\")"
    Name ="tblNewBasicCable_New.[Print Spacing]"
    Expression ="IIf([Forms]![frmCreateItem]![PrintLine4] Like \"*meter*\",\"METER\",\"2FT\")"
    Name ="tblNewBasicCable_New.[Fiber Oracle item]"
    Expression ="Forms!frmFiberSpec![Fiber Oracle item]"
    Name ="tblNewBasicCable_New.[1st Req Freq]"
    Expression ="Forms!frmFiberSpec![1stReqFreq]"
    Name ="tblNewBasicCable_New.[1st Max Atten]"
    Expression ="Forms!frmFiberSpec![1stMaxAtten]"
    Name ="tblNewBasicCable_New.[1st Min BandW]"
    Expression ="IIf(Forms!frmFiberSpec![1stMinBandW]=\"\",Null,Forms!frmFiberSpec![1stMinBandW])"
    Name ="tblNewBasicCable_New.[2nd Req Freq]"
    Expression ="Forms!frmFiberSpec![2ndReqFreq]"
    Name ="tblNewBasicCable_New.[2nd Max Atten]"
    Expression ="Forms!frmFiberSpec![2ndMaxAtten]"
    Name ="tblNewBasicCable_New.[2nd Min BandW]"
    Expression ="IIf(Forms!frmFiberSpec![2ndMinBandW]=\"\",Null,Forms!frmFiberSpec![2ndMinBandW])"
    Name ="tblNewBasicCable_New.[Unit Series]"
    Expression ="Forms!frmCreateItem!UnitSeries"
    Name ="tblNewBasicCable_New.Active"
    Expression ="0"
    Name ="tblNewBasicCable_New.Reason"
    Expression ="\"IP \" & Forms!frmSwitchBoard!Initials"
    Name ="tblNewBasicCable_New.ENumber"
    Expression ="Forms!frmCreateItem!ListingNumber"
    Name ="tblNewBasicCable_New.[Print Height]"
    Expression ="IIf(CDbl([Forms]![frmCreateitem]![NomOD])>9,\"Print Height:  2.9 +/-0.1mm\",\"Pr"
        "int Height:  1.9 +/-0.1mm\")"
    Name ="tblNewBasicCable_New.[UL-ETL-Section]"
    Expression ="Forms!frmCreateItem!ListingSection"
    Name ="tblNewBasicCable_New.[UL-ETL-Const]"
    Expression ="Forms!frmCreateItem!ListingConstruction"
    Name ="tblNewBasicCable_New.[UL-ETL-Listing]"
    Expression ="Forms!frmCreateItem!Listing"
    Name ="tblNewBasicCable_New.[Listing Company]"
    Expression ="Forms!frmCreateItem!ListingCompany"
    Name ="tblNewBasicCable_New.[Label Type]"
    Expression ="getNewItemListingLabel()"
    Name ="tblNewBasicCable_New.CustomerRev"
    Expression ="Forms!frmCreateItem!CustomerRev"
    Name ="tblNewBasicCable_New.[Special Instr Product1]"
    Expression ="\"NONE\""
    Name ="tblNewBasicCable_New.[zzSpecial Instr Product3]"
    Expression ="\"NONE\""
    Name ="tblNewBasicCable_New.[zzSpecial Instr Product4]"
    Expression ="\"NONE\""
    Name ="tblNewBasicCable_New.Base"
    Expression ="[Forms]![frmCreateItem]![Base]"
    Name ="tblNewBasicCable_New.[Color Chip ID]"
    Expression ="Null"
    Name ="tblNewBasicCable_New.FiberType2"
    Expression ="[Forms]![frmCreateitem]![Fiber2]"
    Name ="tblNewBasicCable_New.FiberType3"
    Expression ="\"NONE\""
    Name ="tblNewBasicCable_New.VendorSpecificFiber"
    Expression ="GetisVendorSpecificPrint()"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x5b7b7a89a5277941b5bba7269bbca0c9
End
Begin
    Begin
        dbText "Name" ="tblNewPrductConstruction.[Fiber Type]"
        dbInteger "ColumnWidth" ="3870"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[Fiber Type]"
        dbInteger "ColumnWidth" ="3870"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[Fiber Type A]"
        dbInteger "ColumnWidth" ="3870"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[Item No]"
        dbInteger "ColumnWidth" ="2040"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.Customer"
        dbInteger "ColumnWidth" ="1275"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[Customer Part Num]"
        dbInteger "ColumnWidth" ="1905"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[Revision History]"
        dbInteger "ColumnWidth" ="8400"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[Print Height]"
        dbInteger "ColumnWidth" ="2550"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[Print Type (base)]"
        dbInteger "ColumnWidth" ="3630"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[Special Instr Product3]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[New Oracle Part Number]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[Special Instr Product4]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[Special Instr Product1]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.ENumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[Special Instr Product2]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.PrintSpacing"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[Fiber Type C]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.Base"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[Print Line 3]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[Print Line 4]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.NewPrintLine4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[1st Min BandW - A]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[Unit Series]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[1st Max Atten - A]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[2nd Min BandW - A]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[1st Req Freq - A]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[2nd Max Atten - A]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[2nd Req Freq - A]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[Revision Date]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[Revision Letter]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[Jacket Color]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[Fiber Type B]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[1st Req Freq - B]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[2nd Max Atten - B]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.OracleStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[Jacket Material]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[2nd Req Freq - B]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[Print Reel No]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[Print Item No]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[Print Line 1]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[Print Line 2]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.Active"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[1st Min BandW - B]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.Reason"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[1st Max Atten - B]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[2nd Min BandW - B]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[UL-ETL-Section]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[UL-ETL-Const]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[UL-ETL-Listing]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[Listing Company]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.CustomerRev"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[Label Type]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.ColorChip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.PrintSpacing"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.Customer"
        dbInteger "ColumnWidth" ="1275"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[Revision History]"
        dbInteger "ColumnWidth" ="8400"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[Print Type (base)]"
        dbInteger "ColumnWidth" ="3630"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[Print Line 2]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[1st Min BandW - A]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[Unit Series]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[Fiber Type A]"
        dbInteger "ColumnWidth" ="3870"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[Customer Part Num]"
        dbInteger "ColumnWidth" ="1905"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.OracleStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[Print Reel No]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[Print Line 3]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[2nd Req Freq - A]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[1st Req Freq - A]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[New Oracle Part Number]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[Revision Date]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[Jacket Material]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[Print Item No]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[Print Line 4]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[2nd Max Atten - A]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[Item No]"
        dbInteger "ColumnWidth" ="2040"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[Revision Letter]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[Jacket Color]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[Print Line 1]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.NewPrintLine4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[1st Max Atten - A]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[2nd Min BandW - A]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.Active"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[2nd Max Atten]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.Reason"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[2nd Min BandW]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[Fiber Type B]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[zzSpecial Instr Product3]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[1st Req Freq - B]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[zzSpecial Instr Product4]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[1st Max Atten - B]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[Color Chip ID]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[1st Min BandW - B]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[2nd Req Freq - B]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[2nd Max Atten - B]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[2nd Min BandW - B]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.ENumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[Print Height]"
        dbInteger "ColumnWidth" ="2550"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[UL-ETL-Section]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[UL-ETL-Const]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[UL-ETL-Listing]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[Listing Company]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[Label Type]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.CustomerRev"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[Special Instr Product1]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[Special Instr Product2]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[Special Instr Product3]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[Special Instr Product4]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[Fiber Type C]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.Base"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.ColorChip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[New Oracle Part #]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[Customer Part#]"
        dbInteger "ColumnWidth" ="1905"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[Print Spacing]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[Fiber Oracle item]"
        dbInteger "ColumnWidth" ="3870"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[1st Req Freq]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[1st Max Atten]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[1st Min BandW]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.[2nd Req Freq]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.FiberType2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.FiberType3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable_New.VendorSpecificFiber"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1109
    Bottom =669
    Left =-1
    Top =-1
    Right =1077
    Bottom =322
    Left =49
    Top =17
    ColumnsShown =579
    Begin
        Left =578
        Top =12
        Right =840
        Bottom =250
        Top =0
        Name ="tblNewBasicCable_New"
        Name =""
    End
End

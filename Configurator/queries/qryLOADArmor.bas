Operation =4
Option =8
Where ="(((tblArmoredSetups.ArmorNo)=[Forms]![frmCreateItem]![ArmorNo]))"
Begin InputTables
    Name ="tblNewBasicCable"
    Name ="tblArmoredSetups"
End
Begin OutputColumns
    Name ="tblNewBasicCable.[CM Type]"
    Expression ="Left(Forms!frmCreateItem![NewOracle#],10)"
    Name ="tblNewBasicCable.[Unit Series]"
    Expression ="Null"
    Name ="tblNewBasicCable.OracleStatus"
    Expression ="\"Active\""
    Name ="tblNewBasicCable.Base"
    Expression ="[OracleBase]"
    Name ="tblNewBasicCable.[CM Matrl]"
    Expression ="Null"
    Name ="tblNewBasicCable.[CM OD]"
    Expression ="Null"
    Name ="tblNewBasicCable.[CM OD Tol +]"
    Expression ="Null"
    Name ="tblNewBasicCable.[CM OD Tol -]"
    Expression ="Null"
    Name ="tblNewBasicCable.[CM Min Wall]"
    Expression ="Null"
    Name ="tblNewBasicCable.[Binder #1]"
    Expression ="Null"
    Name ="tblNewBasicCable.[Qty Binder #1]"
    Expression ="Null"
    Name ="tblNewBasicCable.[Binder #2]"
    Expression ="Null"
    Name ="tblNewBasicCable.[Qty Binder #2]"
    Expression ="Null"
    Name ="tblNewBasicCable.[Lay Length]"
    Expression ="Null"
    Name ="tblNewBasicCable.[Helix Factor]"
    Expression ="Null"
    Name ="tblNewBasicCable.[Binder #3]"
    Expression ="Null"
    Name ="tblNewBasicCable.[Qty Binder #3]"
    Expression ="Null"
    Name ="tblNewBasicCable.Ripcord"
    Expression ="Null"
    Name ="tblNewBasicCable.[Special Instr Product2]"
    Expression ="Null"
    Name ="tblNewBasicCable.[Special Instr Product3]"
    Expression ="Null"
    Name ="tblNewBasicCable.[Special Instr Product4]"
    Expression ="Null"
    Name ="tblNewBasicCable.Aramid1Type"
    Expression ="Null"
    Name ="tblNewBasicCable.Aramid1Count"
    Expression ="Null"
    Name ="tblNewBasicCable.Aramid2Type"
    Expression ="Null"
    Name ="tblNewBasicCable.Aramid2Count"
    Expression ="Null"
    Name ="tblNewBasicCable.[New Oracle Part Number]"
    Expression ="Forms!frmCreateItem![NewOracle#]"
    Name ="tblNewBasicCable.[Item No]"
    Expression ="Forms!frmCreateItem!ItemNo"
    Name ="tblNewBasicCable.Customer"
    Expression ="Forms!frmCreateItem!Customer"
    Name ="tblNewBasicCable.[Customer Part Num]"
    Expression ="Forms!frmCreateItem!PID"
    Name ="tblNewBasicCable.[Revision Date]"
    Expression ="Date()"
    Name ="tblNewBasicCable.[Revision Letter]"
    Expression ="1"
    Name ="tblNewBasicCable.[Revision History]"
    Expression ="Date() & \" - INITIAL ISSUE \" & Forms!frmSwitchBoard!Initials"
    Name ="tblNewBasicCable.[Jacket Material]"
    Expression ="Forms!frmCreateItem!Jacket"
    Name ="tblNewBasicCable.[Jacket Color]"
    Expression ="Forms!frmCreateItem!Color"
    Name ="tblNewBasicCable.[Print Type (base)]"
    Expression ="IIf(Forms!frmCreateItem!Color Like \"*black*\",\"SPECIAL: WHITE PRINT\",\"Standa"
        "rd black print\")"
    Name ="tblNewBasicCable.[Print Reel No]"
    Expression ="IIf(Forms!frmCreateItem!PrintType=0,0,-1)"
    Name ="tblNewBasicCable.[Print Item No]"
    Expression ="IIf(Forms!frmCreateItem!PrintType=1,-1,0)"
    Name ="tblNewBasicCable.[Print Line 1]"
    Expression ="Forms!frmCreateItem!PrintLine1"
    Name ="tblNewBasicCable.[Print Line 2]"
    Expression ="Forms!frmCreateItem!PrintLine2"
    Name ="tblNewBasicCable.[Print Line 3]"
    Expression ="Forms!frmCreateItem!PrintLine3"
    Name ="tblNewBasicCable.[Print Line 4]"
    Expression ="Forms!frmCreateItem!PrintLine4"
    Name ="tblNewBasicCable.NewPrintLine4"
    Expression ="IIf(Forms!frmCreateItem!PrintLine4 Like \"*meter*\",\"[DATE2] [METER] METERS\",\""
        "[DATE2] [2FT] FEET\")"
    Name ="tblNewBasicCable.PrintSpacing"
    Expression ="IIf([Forms]![frmCreateItem]![PrintLine4] Like \"*meter*\",\"METER\",\"2FT\")"
    Name ="tblNewBasicCable.[Fiber Type A]"
    Expression ="Forms!frmFiberSpec!FiberType"
    Name ="tblNewBasicCable.[1st Req Freq - A]"
    Expression ="Forms!frmFiberSpec![1stReqFreq]"
    Name ="tblNewBasicCable.[1st Max Atten - A]"
    Expression ="Forms!frmFiberSpec![1stMaxAtten]"
    Name ="tblNewBasicCable.[1st Min BandW - A]"
    Expression ="IIf(Forms!frmFiberSpec![1stMinBandW]=\"\",Null,Forms!frmFiberSpec![1stMinBandW])"
    Name ="tblNewBasicCable.[2nd Req Freq - A]"
    Expression ="Forms!frmFiberSpec![2ndReqFreq]"
    Name ="tblNewBasicCable.[2nd Max Atten - A]"
    Expression ="Forms!frmFiberSpec![2ndMaxAtten]"
    Name ="tblNewBasicCable.[2nd Min BandW - A]"
    Expression ="IIf(Forms!frmFiberSpec![2ndMinBandW]=\"\",Null,Forms!frmFiberSpec![2ndMinBandW])"
    Name ="tblNewBasicCable.Active"
    Expression ="0"
    Name ="tblNewBasicCable.Reason"
    Expression ="\"IP \" & Forms!frmSwitchBoard!Initials"
    Name ="tblNewBasicCable.[Fiber Type B]"
    Expression ="\"NONE\""
    Name ="tblNewBasicCable.[1st Req Freq - B]"
    Expression ="\"0\""
    Name ="tblNewBasicCable.[1st Max Atten - B]"
    Expression ="0"
    Name ="tblNewBasicCable.[1st Min BandW - B]"
    Expression ="\"0\""
    Name ="tblNewBasicCable.[2nd Req Freq - B]"
    Expression ="\"0\""
    Name ="tblNewBasicCable.[2nd Max Atten - B]"
    Expression ="0"
    Name ="tblNewBasicCable.[2nd Min BandW - B]"
    Expression ="\"0\""
    Name ="tblNewBasicCable.[UL-ETL-Section]"
    Expression ="Forms!frmCreateItem!ListingSection"
    Name ="tblNewBasicCable.[UL-ETL-Const]"
    Expression ="Forms!frmCreateItem!ListingConstruction"
    Name ="tblNewBasicCable.[UL-ETL-Listing]"
    Expression ="Forms!frmCreateItem!Listing"
    Name ="tblNewBasicCable.[Listing Company]"
    Expression ="Forms!frmCreateItem!ListingCompany"
    Name ="tblNewBasicCable.[Label Type]"
    Expression ="\"(\" & Forms!frmCreateItem!ListingCompany & \")c(\" & Forms!frmCreateItem!Listi"
        "ngCompany & \")\""
    Name ="tblNewBasicCable.CustomerRev"
    Expression ="Forms!frmCreateItem!CustomerRev"
    Name ="tblNewBasicCable.ENumber"
    Expression ="Forms!frmCreateItem!ListingNumber"
    Name ="tblNewBasicCable.[Nom Wall]"
    Expression ="Null"
    Name ="tblNewBasicCable.[Min Ave Wall]"
    Expression ="Null"
    Name ="tblNewBasicCable.[Min Spot Wall]"
    Expression ="Null"
    Name ="tblNewBasicCable.[Max Ave Wall]"
    Expression ="Null"
    Name ="tblNewBasicCable.[Nominal OD]"
    Expression ="Null"
    Name ="tblNewBasicCable.[OD Tol  (+)]"
    Expression ="Null"
    Name ="tblNewBasicCable.[OD Tol  (-)]"
    Expression ="Null"
    Name ="tblNewBasicCable.[Special Instr Product1]"
    Expression ="Null"
    Name ="tblNewBasicCable.[Print Height]"
    Expression ="\"Print Height:  2.9 +/-0.1mm\""
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0xb809e3e3eaa9a1449e439c61c60063e1
End
Begin
    Begin
        dbText "Name" ="tblNewBasicCable.[CM Type]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.OracleStatus"
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
        dbText "Name" ="tblNewBasicCable.[Nominal OD]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[Min Ave Wall]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[Core Wrap]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[Nom Wall]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[OD Tol  (-)]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[OD Tol  (+)]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[PSS Document #]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[Special Instr Product1]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[Max Ave Wall]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[FRP Dia]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[Min Spot Wall]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[Core Dia]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.PrintSpacing"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[Special Instr Product2]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[Special Instr Product3]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.[Print Height]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.ColorChip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewBasicCable.CustomerRev"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1292
    Bottom =772
    Left =-1
    Top =-1
    Right =1260
    Bottom =77
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =249
        Top =-17
        Right =673
        Bottom =191
        Top =0
        Name ="tblNewBasicCable"
        Name =""
    End
    Begin
        Left =851
        Top =-9
        Right =1248
        Bottom =229
        Top =0
        Name ="tblArmoredSetups"
        Name =""
    End
End

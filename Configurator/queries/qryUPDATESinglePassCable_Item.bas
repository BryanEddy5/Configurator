Operation =4
Option =0
Begin InputTables
    Name ="Basic Product Construction"
    Name ="tblSinglePassCable"
End
Begin OutputColumns
    Name ="[Basic Product Construction].[New Oracle Part #]"
    Expression ="[tblSinglePassCable]![New Oracle Part #]"
    Name ="[Basic Product Construction].[Item No]"
    Expression ="[tblSinglePassCable]![Item No]"
    Name ="[Basic Product Construction].Customer"
    Expression ="[tblSinglePassCable]![Customer]"
    Name ="[Basic Product Construction].[Customer Part#]"
    Expression ="[tblSinglePassCable]![Customer Part#]"
    Name ="[Basic Product Construction].CustomerRev"
    Expression ="[tblSinglePassCable]![CustomerRev]"
    Name ="[Basic Product Construction].Active"
    Expression ="[tblsinglepasscable]![ActiveItem]"
    Name ="[Basic Product Construction].Reason"
    Expression ="[tblSinglePassCable]"
    Name ="[Basic Product Construction].Base"
    Expression ="[tblSinglePassCable]![Base]"
    Name ="[Basic Product Construction].[Revision Date]"
    Expression ="[tblSinglePassCable]![Revision Date]"
    Name ="[Basic Product Construction].[Revision Letter]"
    Expression ="[tblSinglePassCable]![Revision Letter]"
    Name ="[Basic Product Construction].[Revision History]"
    Expression ="[tblSinglePassCable]![Revision History]"
    Name ="[Basic Product Construction].[Jacket Material]"
    Expression ="[tblSinglePassCable]![JacketMaterial]"
    Name ="[Basic Product Construction].[Jacket Color]"
    Expression ="[tblSinglePassCable]![Jacket Color]"
    Name ="[Basic Product Construction].[Color Chip ID]"
    Expression ="[tblSinglePassCable]![Color Chip ID]"
    Name ="[Basic Product Construction].[Print Reel No]"
    Expression ="[tblSinglePassCable]![Print Reel No]"
    Name ="[Basic Product Construction].[Print Item No]"
    Expression ="[tblSinglePassCable]![Print Item No]"
    Name ="[Basic Product Construction].[Print Type (base)]"
    Expression ="[tblsinglepasscable]![Print Type (base)]"
    Name ="[Basic Product Construction].[Print Line 1]"
    Expression ="[tblsinglepasscable]![Print Line 1]"
    Name ="[Basic Product Construction].[Print Line 2]"
    Expression ="[tblsinglepasscable]![Print Line 2]"
    Name ="[Basic Product Construction].[Print Line 3]"
    Expression ="[tblsinglepasscable]![Print Line 3]"
    Name ="[Basic Product Construction].[Print Line 4]"
    Expression ="[tblsinglepasscable]![Print Line 4]"
    Name ="[Basic Product Construction].NewPrintLine4"
    Expression ="[tblsinglepasscable]![NewPrintLine4]"
    Name ="[Basic Product Construction].[Print Spacing]"
    Expression ="[tblsinglepasscable]![Print Spacing]"
    Name ="[Basic Product Construction].PrintNotes"
    Expression ="[tblsinglepasscable]![PrintNotes]"
    Name ="[Basic Product Construction].[Listing Company]"
    Expression ="[tblsinglepasscable]![Listing Company]"
    Name ="[Basic Product Construction].[UL-ETL-Listing]"
    Expression ="[tblsinglepasscable]![UL-ETL-Listing]"
    Name ="[Basic Product Construction].[UL-ETL-Const]"
    Expression ="[tblsinglepasscable]![UL-ETL-Const]"
    Name ="[Basic Product Construction].[UL-ETL-Section]"
    Expression ="[tblsinglepasscable]![UL-ETL-Section]"
    Name ="[Basic Product Construction].ENumber"
    Expression ="[tblsinglepasscable]![ENumber]"
    Name ="[Basic Product Construction].[Label Type]"
    Expression ="[tblsinglepasscable]![labe]"
    Name ="[Basic Product Construction].[EZ Strip]"
    Expression ="[tblsinglepasscable]![EZ Strip]"
    Name ="[Basic Product Construction].[TB Material]"
    Expression ="[tblsinglepasscable]![TB Material]"
    Name ="[Basic Product Construction].[TB Chips Type]"
    Expression ="[tblsinglepasscable]![TB Chips Type]"
    Name ="[Basic Product Construction].[TB Nominal OD]"
    Expression ="[tblsinglepasscable]![TB Nominal OD]"
    Name ="[Basic Product Construction].[TB OD Tol  (+)]"
    Expression ="[tblsinglepasscable]![TB OD Tol  (+)]"
    Name ="[Basic Product Construction].[TB OD Tol  (-)]"
    Expression ="[tblsinglepasscable]![TB OD Tol  (-)]"
    Name ="[Basic Product Construction].[TB Color Series]"
    Expression ="[tblsinglepasscable]![TB Color Series]"
    Name ="[Basic Product Construction].[Fiber Type]"
    Expression ="[tblsinglepasscable]![Fiber Type]"
    Name ="[Basic Product Construction].[Fiber Oracle item]"
    Expression ="[tblsinglepasscable]![Fiber Oracle item]"
    Name ="[Basic Product Construction].FiberType2"
    Expression ="[tblsinglepasscable]![FiberType2]"
    Name ="[Basic Product Construction].FiberType3"
    Expression ="[tblsinglepasscable]![FiberType3]"
    Name ="[Basic Product Construction].[1st Req Freq]"
    Expression ="[tblsinglepasscable]![1st Req Freq]"
    Name ="[Basic Product Construction].[1st Max Atten]"
    Expression ="[tblsinglepasscable]![1st Max Atten]"
    Name ="[Basic Product Construction].[1st Min BandW]"
    Expression ="[tblsinglepasscable]![1st Min BandW]"
    Name ="[Basic Product Construction].[2nd Req Freq]"
    Expression ="[tblsinglepasscable]![2nd Req Freq]"
    Name ="[Basic Product Construction].[2nd Max Atten]"
    Expression ="[tblsinglepasscable]![2nd Max Atten]"
    Name ="[Basic Product Construction].[2nd Min BandW]"
    Expression ="[tblsinglepasscable]![2nd Min BandW]"
    Name ="[Basic Product Construction].[SM 1300 Max Atten]"
    Expression ="[tblsinglepasscable]![SM 1300 Max Atten]"
    Name ="[Basic Product Construction].[SM 1550 Max Atten]"
    Expression ="[tblsinglepasscable]![SM 1550 Max Atten]"
    Name ="[Basic Product Construction].[Special Instr Product1]"
    Expression ="[tblsinglepasscable]![Special Instr Product1]"
End
Begin Joins
    LeftTable ="Basic Product Construction"
    RightTable ="tblSinglePassCable"
    Expression ="[Basic Product Construction].[New Oracle Part #] = tblSinglePassCable.[New Oracl"
        "e Part #]"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x0e87d40aa05cce42bbed4a7f436f439c
End
Begin
    Begin
        dbText "Name" ="[Basic Product Construction].RevisedDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[New Oracle Part #]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].CustomerRev"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].Base"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblSinglePassCable.[zzSpecial Instr Product4]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="zzSpecial Instr Product3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Item No]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].Active"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblSinglePassCable.[New Oracle Part #]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Special Instr Product1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].Customer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].Reason"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="zzSpecial Instr Product4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].BaseID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Customer Part#]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[calculate bill?]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[OD Tol  (-)]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Cable Type]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Standard Operation]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Nominal OD]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Height Tol (+)]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Min Ave Wall]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Std Subunit Print]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Revision Date]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Jacket Material]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[OD Tol  (+)]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Height Tol (-)]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Min Spot Wall]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Revision Letter]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Jacket Color]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Nominal Wall]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Print Reel No]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Revision History]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Color Chip ID]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].Height"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Max Ave Wall]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Print Item No]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Print Type (base)]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[2nd Min BandW]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Print Line 4]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[zzSpecial Instr Product3]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Listing Company]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[PSS Document #]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].ENumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[EZ Strip]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[TB OD Tol  (+)]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Fiber Oracle item]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Print Line 1]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[SM 1300 Max Atten]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].NewPrintLine4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[zzSpecial Instr Product4]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[UL-ETL-Listing]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].DateCreated"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Label Type]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[TB Material]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[TB OD Tol  (-)]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].FiberType2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Print Line 2]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[SM 1550 Max Atten]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Print Spacing]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[zzOracle Part#]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[UL-ETL-Const]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].CreatedBy"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Aramid Type]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[TB Chips Type]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[TB Color Series]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].FiberType3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Print Line 3]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Special Instr Product1]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].PrintNotes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].OracleStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[UL-ETL-Section]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].RevisedBy"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].Talc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[TB Nominal OD]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Fiber Type]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[1st Req Freq]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[1st Max Atten]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[1st Min BandW]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[2nd Req Freq]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[2nd Max Atten]"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =43
    Top =91
    Right =1152
    Bottom =720
    Left =-1
    Top =-1
    Right =1077
    Bottom =554
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =48
        Top =12
        Right =477
        Bottom =433
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
    Begin
        Left =525
        Top =12
        Right =929
        Bottom =433
        Top =0
        Name ="tblSinglePassCable"
        Name =""
    End
End

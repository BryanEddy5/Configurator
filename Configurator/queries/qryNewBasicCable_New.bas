Operation =3
Name ="tblNewBasicCable_New"
Option =0
Begin InputTables
    Name ="qryCopyItemMatch"
End
Begin OutputColumns
    Alias ="Oracle"
    Name ="New Oracle Part #"
    Expression ="qryCopyItemMatch.[New Oracle Part #]"
End
Begin OrderBy
    Expression ="qryCopyItemMatch.[New Oracle Part #]"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0xc600700a971a934391e1860bfd211850
End
Begin
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].BaseID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[New Oracle Part Number]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].Aramid2Count"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[Item No]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[CM OD Tol +]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].Base"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[PSS Document #]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[UL-ETL-Const]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[CM Type]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[UL-ETL-Section]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[CM OD Tol -]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[UL-ETL-Listing]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[Revision History]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[CM Matrl]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[Revision Letter]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].Customer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[Print Line 3]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].PrintNotes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].CustomerRev"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[Print Line 1]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].NewPrintLine4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].Active"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[Print Line 2]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].PrintSpacing"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[Revision Date]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[Print Line 4]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[Listing Company]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[CM OD]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[Print Reel No]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[Customer Part Num]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[Print Type (base)]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].Reason"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[Print Height]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[Print Item No]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[Unit Series]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[FRP Dia]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[CM Wall]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[Jacket Material]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[Nom Wall]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[CM Min Wall]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[Core Wrap]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[Binder #1]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[Binder #2]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[Helix Factor]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[2nd Min BandW - A]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[load into oracle 11?]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[Lay Length]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[1st Max Atten - C]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[Min Ave Wall]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[OD Tol  (+)]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[Min Spot Wall]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[OD Tol  (-)]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[1st Max Atten - A]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[2nd Min BandW - C]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[Max Ave Wall]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[Chip Type]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[Nominal OD]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[2nd Req Freq - B]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[Qty Binder #1]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].Binder2LayLength"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[1st Req Freq - C]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].Binder1LayLength"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[2nd Max Atten - A]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[Qty Binder #2]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[Fiber Type C]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[Fiber Type A]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[1st Req Freq - A]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[2nd Max Atten - C]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[Standard Operation]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[Binder #3]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[Core Dia]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[Qty Binder #3]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].Binder3LayLength"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[Jacket Color]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].Ripcord"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].ColorChip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[1st Min BandW - A]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[Special Instr Product1]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].Aramid1Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[2nd Req Freq - A]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[Fiber Type B]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[1st Req Freq - B]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[Label Type]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[1st Max Atten - B]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[1st Min BandW - B]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[2nd Max Atten - B]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].Aramid2Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[2nd Min BandW - B]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[1st Min BandW - C]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[2nd Req Freq - C]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[calculate bill?]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[Special Instr Product2]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[Special Instr Product3]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[Special Instr Product4]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].ENumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].OracleStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].Aramid1Count"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryCopyItemMatch.Oracle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Oracle"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x00000000000000000000000001000000
        End
    End
    Begin
        dbText "Name" ="Expr1"
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
    Bottom =281
    Left =5
    Top =0
    ColumnsShown =651
    Begin
        Left =96
        Top =19
        Right =744
        Bottom =176
        Top =0
        Name ="qryCopyItemMatch"
        Name =""
    End
End

Operation =1
Option =0
Begin InputTables
    Name ="tblCableConstructionReferences"
    Name ="tblCableConstructions"
    Name ="CableRunSpeeds"
End
Begin OutputColumns
    Expression ="tblCableConstructionReferences.Base"
    Expression ="tblCableConstructionReferences.PrefixID"
    Expression ="tblCableConstructions.NominalOD"
    Expression ="CableRunSpeeds.PrimaryBufferingLine"
    Expression ="tblCableConstructionReferences.CableFamily"
    Expression ="tblCableConstructionReferences.CableLevel1"
    Expression ="CableRunSpeeds.[Buff#11 Line Speed]"
    Expression ="CableRunSpeeds.[Buff#1 Line Speed]"
    Expression ="CableRunSpeeds.[Buff#4 Line Speed]"
    Expression ="CableRunSpeeds.[Buff#6 Line Speed]"
    Expression ="CableRunSpeeds.[Buff#2 Line Speed]"
    Expression ="CableRunSpeeds.[Buff#9 Line Speed]"
    Expression ="CableRunSpeeds.[Buff#10 Line Speed]"
End
Begin Joins
    LeftTable ="tblCableConstructionReferences"
    RightTable ="tblCableConstructions"
    Expression ="tblCableConstructionReferences.BaseID = tblCableConstructions.BaseID"
    Flag =1
    LeftTable ="tblCableConstructionReferences"
    RightTable ="CableRunSpeeds"
    Expression ="tblCableConstructionReferences.PrefixID = CableRunSpeeds.PrefixID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbMemo "Filter" ="([Query1].[Base] Like \"s?001*\")"
dbMemo "OrderBy" ="[Query1].[CableLevel1], [Query1].[PrimaryBufferingLine], [Query1].[NominalOD], ["
    "Query1].[Base], [Query1].[PrefixID]"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xe2dcf73f98c40f4da74259971a2b23bc
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblCableConstructionReferences.Base"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.CableFamily"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CableRunSpeeds.[Buff#9 Line Speed]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.PrefixID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.CableLevel1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CableRunSpeeds.[Buff#4 Line Speed]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CableRunSpeeds.[Buff#11 Line Speed]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CableRunSpeeds.[Buff#10 Line Speed]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.NominalOD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CableRunSpeeds.[Buff#6 Line Speed]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CableRunSpeeds.[Buff#1 Line Speed]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CableRunSpeeds.PrimaryBufferingLine"
        dbInteger "ColumnWidth" ="2400"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CableRunSpeeds.[Buff#2 Line Speed]"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-310
    Top =118
    Right =799
    Bottom =747
    Left =-1
    Top =-1
    Right =1077
    Bottom =476
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =145
        Top =109
        Right =383
        Bottom =380
        Top =0
        Name ="tblCableConstructionReferences"
        Name =""
    End
    Begin
        Left =492
        Top =38
        Right =837
        Bottom =288
        Top =0
        Name ="tblCableConstructions"
        Name =""
    End
    Begin
        Left =989
        Top =46
        Right =1208
        Bottom =422
        Top =0
        Name ="CableRunSpeeds"
        Name =""
    End
End

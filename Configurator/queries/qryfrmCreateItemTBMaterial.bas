Operation =1
Option =0
Begin InputTables
    Name ="qryDesignCodeBreakdown"
    Name ="Basic Product Construction"
End
Begin OutputColumns
    Expression ="qryDesignCodeBreakdown.CableType"
    Expression ="qryDesignCodeBreakdown.TBType"
    Expression ="[Basic Product Construction].[Jacket Material]"
    Alias ="CountOfTBType"
    Expression ="Count(qryDesignCodeBreakdown.TBType)"
    Expression ="[Basic Product Construction].[TB Material]"
    Expression ="[Basic Product Construction].[EZ Strip]"
End
Begin Joins
    LeftTable ="qryDesignCodeBreakdown"
    RightTable ="Basic Product Construction"
    Expression ="qryDesignCodeBreakdown.Oracle = [Basic Product Construction].[New Oracle Part #]"
    Flag =1
End
Begin OrderBy
    Expression ="qryDesignCodeBreakdown.CableType"
    Flag =0
    Expression ="qryDesignCodeBreakdown.TBType"
    Flag =0
    Expression ="[Basic Product Construction].[Jacket Material]"
    Flag =0
    Expression ="Count(qryDesignCodeBreakdown.TBType)"
    Flag =1
End
Begin Groups
    Expression ="qryDesignCodeBreakdown.CableType"
    GroupLevel =0
    Expression ="qryDesignCodeBreakdown.TBType"
    GroupLevel =0
    Expression ="[Basic Product Construction].[Jacket Material]"
    GroupLevel =0
    Expression ="[Basic Product Construction].[TB Material]"
    GroupLevel =0
    Expression ="[Basic Product Construction].[EZ Strip]"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x2898a1a646f69b4b96fcba4c457bc396
End
dbMemo "Filter" ="([qryfrmCreateItemTBMaterial].[CableType]=\"E\")"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="qryDesignCodeBreakdown.CableType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryDesignCodeBreakdown.TBType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Jacket Material]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CountOfTBType"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x618fb850fbde274e9230e9c81577fa30
        End
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[TB Material]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[EZ Strip]"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1579
    Bottom =881
    Left =-1
    Top =-1
    Right =1547
    Bottom =196
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =68
        Top =13
        Right =362
        Bottom =176
        Top =0
        Name ="qryDesignCodeBreakdown"
        Name =""
    End
    Begin
        Left =440
        Top =6
        Right =646
        Bottom =169
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
End

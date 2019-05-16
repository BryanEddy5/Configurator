Operation =1
Option =0
Where ="(((Mid([Item No],13,20)) Not Like \"*/*\") AND ((qryItemNoJoin.Active)=True) AND"
    " ((qryItemNoJoin.Oracle) Not Like \"*test*\"))"
Begin InputTables
    Name ="qryItemNoJoin"
End
Begin OutputColumns
    Alias ="CableType"
    Expression ="Left([Item No],1)"
    Alias ="Jacket"
    Expression ="Mid([Item No],2,1)"
    Alias ="FiberCount"
    Expression ="Mid([Item No],3,3)"
    Alias ="FiberType"
    Expression ="Mid([Item No],6,1)"
    Alias ="OD"
    Expression ="Mid([Item No],7,2)"
    Alias ="Print"
    Expression ="Mid([Item No],9,1)"
    Alias ="JacketColor"
    Expression ="Mid([Item No],10,1)"
    Alias ="TBColor"
    Expression ="Mid([Item No],11,1)"
    Alias ="TBType"
    Expression ="Mid([Item No],12,1)"
    Alias ="Ext"
    Expression ="Mid([Item No],13,20)"
    Expression ="qryItemNoJoin.Active"
    Expression ="qryItemNoJoin.Oracle"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x3303ee55b7b9934aa1e10c3ee004786a
End
dbMemo "Filter" ="([qryDesignCodeBreakdown1].[JacketColor]=\"S\")"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="CableType"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xaafc6d923cf4a7479423ea2662fa68a7
        End
    End
    Begin
        dbText "Name" ="Jacket"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x87470deee1232c4092de93b97b56d482
        End
    End
    Begin
        dbText "Name" ="FiberCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x0d0b53b9a164e344a92848001afac459
        End
    End
    Begin
        dbText "Name" ="FiberType"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x55602befabace84b8b772e0c228992e5
        End
    End
    Begin
        dbText "Name" ="OD"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x8eb1f195a273f849af40d93a17b66bb0
        End
    End
    Begin
        dbText "Name" ="Print"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd590959c5c01274192a20af055ad70a8
        End
    End
    Begin
        dbText "Name" ="JacketColor"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xcab71f09d1f4134796f936d79968ecd6
        End
    End
    Begin
        dbText "Name" ="TBColor"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf1dc1e8a2fdf614e9d6c08e7799d84bf
        End
    End
    Begin
        dbText "Name" ="TBType"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb2e4daedf0f0fc4c8b5b8d6cdb6c2117
        End
    End
    Begin
        dbText "Name" ="Ext"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9abae93d2df5834aa15aac3506dde4ab
        End
    End
    Begin
        dbText "Name" ="qryItemNoJoin.Active"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryItemNoJoin.Oracle"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-91
    Top =73
    Right =1521
    Bottom =515
    Left =-1
    Top =-1
    Right =1580
    Bottom =194
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =469
        Bottom =169
        Top =0
        Name ="qryItemNoJoin"
        Name =""
    End
End

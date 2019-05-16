Operation =1
Option =0
Begin InputTables
    Name ="qryDesignCodeSubFiberCount"
End
Begin OutputColumns
    Alias ="Expr1"
    Expression ="qryDesignCodeSubFiberCount.[Level 1 Product Compound]"
    Alias ="Expr2"
    Expression ="qryDesignCodeSubFiberCount.Product"
    Alias ="Expr3"
    Expression ="qryDesignCodeSubFiberCount.[Level 1 Desription]"
    Alias ="SubFiberCount"
    Expression ="Max(Int([Subcount]))"
    Alias ="Count"
    Expression ="Count(qryDesignCodeSubFiberCount.Subcount)"
End
Begin OrderBy
    Expression ="Count(qryDesignCodeSubFiberCount.Subcount)"
    Flag =1
End
Begin Groups
    Expression ="qryDesignCodeSubFiberCount.[Level 1 Product Compound]"
    GroupLevel =0
    Expression ="qryDesignCodeSubFiberCount.Product"
    GroupLevel =0
    Expression ="qryDesignCodeSubFiberCount.[Level 1 Desription]"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x5231aa05c33a1c448613dea35b9827e7
End
dbMemo "OrderBy" ="[qryDesignCodeSubCountAggregate].[Product]"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="SubFiberCount"
        dbInteger "ColumnWidth" ="2085"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x29a9d0641e6c76448123a1b8dc7b451a
        End
    End
    Begin
        dbText "Name" ="qryDesignCodeSubFiberCount.[Level 1 Product Compound]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Count"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb279eed4a48a994e91d0f95f9ad84410
        End
    End
    Begin
        dbText "Name" ="qryDesignCodeSubFiberCount.Product"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryDesignCodeSubFiberCount.[Level 1 Desription]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x56475089d0e6ff46894d842dbb04e1a9
        End
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc18c9a14e3daca438f6046c23fc98c6f
        End
    End
    Begin
        dbText "Name" ="Expr3"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x0ad3aca82bb5cd449a3b6530b9ea0316
        End
    End
End
Begin
    State =0
    Left =-376
    Top =111
    Right =1274
    Bottom =717
    Left =-1
    Top =-1
    Right =1618
    Bottom =107
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =202
        Bottom =139
        Top =0
        Name ="qryDesignCodeSubFiberCount"
        Name =""
    End
End

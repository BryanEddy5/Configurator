Operation =1
Option =0
Begin InputTables
    Name ="qryDesignCodeBase"
    Name ="Product Desrcriptions"
    Name ="tblCableConstructionReferences"
End
Begin OutputColumns
    Expression ="qryDesignCodeBase.Base"
    Expression ="qryDesignCodeBase.FiberCount"
    Expression ="qryDesignCodeBase.SubOD"
    Expression ="[Product Desrcriptions].Desrciprion"
    Expression ="qryDesignCodeBase.OD"
    Alias ="Count"
    Expression ="Sum(qryDesignCodeBase.CountOfBase)"
    Expression ="qryDesignCodeBase.Jacket"
    Alias ="Expr1"
    Expression ="qryDesignCodeBase.Location"
    Expression ="[Product Desrcriptions].[Level 1 Product]"
    Expression ="qryDesignCodeBase.Base"
    Expression ="qryDesignCodeBase.FiberCount"
End
Begin Joins
    LeftTable ="qryDesignCodeBase"
    RightTable ="tblCableConstructionReferences"
    Expression ="qryDesignCodeBase.Base = tblCableConstructionReferences.Base"
    Flag =1
    LeftTable ="tblCableConstructionReferences"
    RightTable ="Product Desrcriptions"
    Expression ="tblCableConstructionReferences.SetupID = [Product Desrcriptions].ProductID"
    Flag =2
End
Begin OrderBy
    Expression ="qryDesignCodeBase.Jacket"
    Flag =1
    Expression ="qryDesignCodeBase.SubOD"
    Flag =0
    Expression ="qryDesignCodeBase.OD"
    Flag =0
End
Begin Groups
    Expression ="qryDesignCodeBase.SubOD"
    GroupLevel =0
    Expression ="[Product Desrcriptions].Desrciprion"
    GroupLevel =0
    Expression ="qryDesignCodeBase.OD"
    GroupLevel =0
    Expression ="qryDesignCodeBase.Jacket"
    GroupLevel =0
    Expression ="qryDesignCodeBase.Location"
    GroupLevel =0
    Expression ="[Product Desrcriptions].[Level 1 Product]"
    GroupLevel =0
    Expression ="qryDesignCodeBase.Base"
    GroupLevel =0
    Expression ="qryDesignCodeBase.FiberCount"
    GroupLevel =0
    Expression ="qryDesignCodeBase.Base"
    GroupLevel =0
    Expression ="qryDesignCodeBase.FiberCount"
    GroupLevel =0
    Expression ="qryDesignCodeBase.Jacket"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x2fcca22346ad394b8712b1ee9f92c911
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Count"
        dbInteger "ColumnWidth" ="1935"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x74e8b79f09899244947cc16f878d7513
        End
    End
    Begin
        dbText "Name" ="Expr1000"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x8d30d9037f1bd1479f4259fc75793ada
        End
    End
    Begin
        dbText "Name" ="qryDesignCodeBase.OD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryDesignCodeBase.Location"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Product Desrcriptions].[Level 1 Product]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryDesignCodeBase.FiberCount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryDesignCodeBase.Jacket"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryDesignCodeBase.Base"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1001"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Product Desrcriptions].Desrciprion"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryDesignCodeBase.SubOD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xbd47b6e7ebf89f4287ce0d6999c53d73
        End
    End
End
Begin
    State =0
    Left =-240
    Top =48
    Right =1468
    Bottom =543
    Left =-1
    Top =-1
    Right =1676
    Bottom =213
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =375
        Bottom =184
        Top =0
        Name ="qryDesignCodeBase"
        Name =""
    End
    Begin
        Left =761
        Top =-7
        Right =1129
        Bottom =216
        Top =0
        Name ="Product Desrcriptions"
        Name =""
    End
    Begin
        Left =433
        Top =15
        Right =672
        Bottom =241
        Top =0
        Name ="tblCableConstructionReferences"
        Name =""
    End
End

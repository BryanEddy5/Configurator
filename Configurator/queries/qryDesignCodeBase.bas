Operation =1
Option =0
Begin InputTables
    Name ="qryDesignCodeOD"
End
Begin OutputColumns
    Expression ="qryDesignCodeOD.Base"
    Expression ="qryDesignCodeOD.OD"
    Alias ="CountOfBase"
    Expression ="Count(qryDesignCodeOD.Base)"
    Alias ="Expr3"
    Expression ="qryDesignCodeOD.Location"
    Expression ="qryDesignCodeOD.Jacket"
    Alias ="SubOD"
    Expression ="Mid([Base],InStr([Base],\"-\")+1,3)"
    Expression ="qryDesignCodeOD.TBLetter"
    Expression ="qryDesignCodeOD.FiberCount"
End
Begin OrderBy
    Expression ="Count(qryDesignCodeOD.Base)"
    Flag =1
End
Begin Groups
    Expression ="qryDesignCodeOD.Base"
    GroupLevel =0
    Expression ="qryDesignCodeOD.OD"
    GroupLevel =0
    Expression ="qryDesignCodeOD.Location"
    GroupLevel =0
    Expression ="qryDesignCodeOD.Jacket"
    GroupLevel =0
    Expression ="Mid([Base],InStr([Base],\"-\")+1,3)"
    GroupLevel =0
    Expression ="qryDesignCodeOD.TBLetter"
    GroupLevel =0
    Expression ="qryDesignCodeOD.FiberCount"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x18830a0a01ba344e99d04898d88a2850
End
dbMemo "OrderBy" ="qryDesignCodeBase.Base"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="CountOfBase"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x25c30afd26e95044b71f9277c38a1dbd
        End
    End
    Begin
        dbText "Name" ="SubOD"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3188c45a5d8b424b99986183a8f0f510
        End
    End
    Begin
        dbText "Name" ="qryDesignCodeOD.Jacket"
        dbInteger "ColumnWidth" ="4440"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Location"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc00bcd0cc8fc214bbbebfdbed1cdf5d1
        End
    End
    Begin
        dbText "Name" ="TBLetter"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xbb465a7e4cd0c54096dd94abe915aaba
        End
    End
    Begin
        dbText "Name" ="FiberCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x519033a852bd4946b49500282da6120b
        End
    End
    Begin
        dbText "Name" ="Expr3"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x301a4d86e8698c4f8d02e6fac9e09c23
        End
    End
End
Begin
    State =0
    Left =190
    Top =104
    Right =1802
    Bottom =585
    Left =-1
    Top =-1
    Right =1580
    Bottom =250
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =281
        Bottom =154
        Top =0
        Name ="qryDesignCodeOD"
        Name =""
    End
End

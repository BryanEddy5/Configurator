Operation =4
Option =0
Where ="((([BasicProductConstruction_LIVE].[Customer Part#]) Like \"*rev*\"))"
Begin InputTables
    Name ="BasicProductConstruction_LIVE"
End
Begin OutputColumns
    Alias ="Expr1"
    Name ="BasicProductConstruction_LIVE.[Customer Part#]"
    Expression ="Left([Customer Part#],Len([Customer Part#])-Len(InStr([Customer Part#],\"rev\"))"
        "-4)"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x60d2b5498830524cbe1275f3ce540984
End
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
Begin
    Begin
        dbText "Name" ="[Basic Product Construction].[Customer Part#]"
        dbInteger "ColumnWidth" ="3090"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BasicProductConstruction_LIVE.[Customer Part#]"
        dbInteger "ColumnWidth" ="3090"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x95513a5cd85dc94d957e6c3cb69905d3
        End
    End
End
Begin
    State =0
    Left =124
    Top =105
    Right =1875
    Bottom =676
    Left =-1
    Top =-1
    Right =1719
    Bottom =185
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =490
        Top =6
        Right =586
        Bottom =124
        Top =0
        Name ="BasicProductConstruction_LIVE"
        Name =""
    End
End

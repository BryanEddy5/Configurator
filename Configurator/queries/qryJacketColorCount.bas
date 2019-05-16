Operation =1
Option =0
Having ="((([qryJacketColor].[ColorChip]) Is Not Null And ([qryJacketColor].[ColorChip]) "
    "Not Like \"n/a\" And ([qryJacketColor].[ColorChip])<>\"\" And ([qryJacketColor]."
    "[ColorChip])<>\"NA\"))"
Begin InputTables
    Name ="qryJacketColor"
End
Begin OutputColumns
    Alias ="Expr1"
    Expression ="qryJacketColor.Color"
    Alias ="Expr2"
    Expression ="qryJacketColor.Jacket"
    Alias ="Expr3"
    Expression ="qryJacketColor.ColorChip"
    Alias ="Count"
    Expression ="Count(qryJacketColor.ColorChip)"
    Alias ="Expr4"
    Expression ="qryJacketColor.ColorLetter"
End
Begin Groups
    Expression ="qryJacketColor.Color"
    GroupLevel =0
    Expression ="qryJacketColor.Jacket"
    GroupLevel =0
    Expression ="qryJacketColor.ColorChip"
    GroupLevel =0
    Expression ="qryJacketColor.ColorLetter"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xfd55259b5fe6824e8d74d31ac093bfe9
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="qryJacketColor.ColorChip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryJacketColor.Color"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryJacketColor.Jacket"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Count"
        dbInteger "ColumnWidth" ="2115"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x00000000000000000000000001000000
        End
    End
    Begin
        dbText "Name" ="qryJacketColor.ColorLetter"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x492bdfd1aed8194e959587493c9b52ef
        End
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xdbf52e202e63654c9e5250d259572b05
        End
    End
    Begin
        dbText "Name" ="Expr3"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe980bc6aa690dd40a9fa281336288bf8
        End
    End
    Begin
        dbText "Name" ="Expr4"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xcf9556f58a59144cbedb8a39d1fc7d60
        End
    End
End
Begin
    State =0
    Left =-303
    Top =109
    Right =989
    Bottom =841
    Left =-1
    Top =-1
    Right =1260
    Bottom =364
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =364
        Bottom =277
        Top =0
        Name ="qryJacketColor"
        Name =""
    End
End

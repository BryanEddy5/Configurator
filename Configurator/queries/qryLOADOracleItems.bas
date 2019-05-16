Operation =3
Name ="tblOracleItems"
Option =0
Begin InputTables
    Name ="OraclePremiseItems"
End
Begin OutputColumns
    Alias ="Expr1"
    Name ="ItemNumber"
    Expression ="OraclePremiseItems.ItemNumber"
    Alias ="Expr1"
    Name ="ItemDescription"
    Expression ="OraclePremiseItems.CategoryName"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0xee6272b29081ea4d86ce04025c151942
End
Begin
    Begin
        dbText "Name" ="tblOracleItems.ItemDescription"
        dbInteger "ColumnWidth" ="9975"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OraclePremiseItems.ItemNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OraclePremiseItems.CategoryName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe4e43287e0ba9d4586d541d4d1930606
        End
    End
End
Begin
    State =0
    Left =123
    Top =61
    Right =1159
    Bottom =613
    Left =-1
    Top =-1
    Right =1004
    Bottom =192
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =147
        Top =22
        Right =345
        Bottom =192
        Top =0
        Name ="OraclePremiseItems"
        Name =""
    End
End

Operation =1
Option =0
Begin InputTables
    Name ="tblOracleItems"
End
Begin OutputColumns
    Expression ="tblOracleItems.ItemNumber"
    Alias ="Expr1"
    Expression ="tblOracleItems.ItemDescription"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x81864fc37dade04aac223e43b3f587f1
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x760271e665ff3d4f8ac31d4d936429cf
        End
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1109
    Bottom =669
    Left =-1
    Top =-1
    Right =1077
    Bottom =110
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =90
        Top =0
        Name ="tblOracleItems"
        Name =""
    End
End

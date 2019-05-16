Operation =1
Option =2
Begin InputTables
    Name ="tblOracleCatalogs"
End
Begin OutputColumns
    Alias ="Expr1"
    Expression ="tblOracleCatalogs.CatalogGroupName"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x888e59993df6714094b5a838d60b1b3a
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa942f875726d4743a62c2f2decefee8d
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
        Right =233
        Bottom =90
        Top =0
        Name ="tblOracleCatalogs"
        Name =""
    End
End

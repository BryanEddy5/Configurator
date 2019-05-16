Operation =1
Option =2
Begin InputTables
    Name ="qryPremiseTBReplaceTB"
End
Begin OutputColumns
    Alias ="NewBaseComponent"
    Expression ="IIf(Left([NewComponentItem],3)=\"rbn\",[NewComponentItem],Left([NewComponentItem"
        "],7))"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbBinary "GUID" = Begin
    0x99f42d45c297684d94b7d87ee95b59d8
End
Begin
    Begin
        dbText "Name" ="NewBaseComponent"
        dbInteger "ColumnWidth" ="1995"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xad7c6c48eef2ae4eba425276132a29a2
        End
    End
End
Begin
    State =0
    Left =110
    Top =430
    Right =1219
    Bottom =754
    Left =-1
    Top =-1
    Right =1077
    Bottom =93
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =270
        Top =6
        Right =519
        Bottom =124
        Top =0
        Name ="qryPremiseTBReplaceTB"
        Name =""
    End
End

Operation =1
Option =2
Begin InputTables
    Name ="qryPremiseTBReplaceRibbon"
End
Begin OutputColumns
    Alias ="NewBaseComponent"
    Expression ="Left([NewComponentitem],8)"
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
    0x600439e22f738243aa56362f8e7617ea
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
    Left =-418
    Top =129
    Right =691
    Bottom =611
    Left =-1
    Top =-1
    Right =1077
    Bottom =192
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =524
        Bottom =201
        Top =0
        Name ="qryPremiseTBReplaceRibbon"
        Name =""
    End
End

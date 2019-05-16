Operation =1
Option =0
Where ="(((tblBOMOpResourcesInterface.ResourceCode)=\"0\"))"
Begin InputTables
    Name ="tblBOMOpResourcesInterface"
End
Begin OutputColumns
    Alias ="ComponentItemNumber"
    Expression ="\"\""
    Alias ="ItemSeqNumber"
    Expression ="\"\""
    Alias ="CheckCode"
    Expression ="\"Routing - Unspecified Resource\""
    Expression ="tblBOMOpResourcesInterface.OperationSeqNumber"
    Expression ="tblBOMOpResourcesInterface.ResourceCode"
    Alias ="StopLoad"
    Expression ="-1"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xb71a8028974bee4e8470032f8f35a006
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="CheckCode"
        dbInteger "ColumnWidth" ="3090"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x018b264cae72464da8427c6abc7df13e
        End
    End
    Begin
        dbText "Name" ="tblBOMOpResourcesInterface.OperationSeqNumber"
        dbInteger "ColumnWidth" ="2040"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ComponentItemNumber"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xdcc0c16ca7adcf459c64c79b5c7ddd58
        End
    End
    Begin
        dbText "Name" ="ItemSeqNumber"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x82c6df8df765684a94259dfed42ecd09
        End
    End
    Begin
        dbText "Name" ="StopLoad"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x0d1ee468bfe9bf479b6fbd01dfb2290f
        End
    End
End
Begin
    State =0
    Left =18
    Top =40
    Right =1418
    Bottom =443
    Left =-1
    Top =-1
    Right =1368
    Bottom =193
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =294
        Top =6
        Right =605
        Bottom =195
        Top =0
        Name ="tblBOMOpResourcesInterface"
        Name =""
    End
End

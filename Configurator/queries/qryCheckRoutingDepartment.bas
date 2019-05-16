Operation =1
Option =0
Where ="(((tblBOMOpSequencesInterfaceTemp.DepartmentCode) Is Null))"
Begin InputTables
    Name ="tblBOMOpSequencesInterfaceTemp"
End
Begin OutputColumns
    Alias ="ComponentItemNumber"
    Expression ="\"\""
    Alias ="ItemSeqNumber"
    Expression ="\"\""
    Alias ="CheckCode"
    Expression ="\"Routing - Unspecified Department\""
    Expression ="tblBOMOpSequencesInterfaceTemp.OperationSeqNum"
    Expression ="tblBOMOpSequencesInterfaceTemp.DepartmentCode"
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
    0x5b0c913f3f50184398862d9b4b58159c
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
            0xc8804656ff8a62499d525b837fb237d8
        End
    End
    Begin
        dbText "Name" ="ComponentItemNumber"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x379929932f52b84fb60a51c5e0e19205
        End
    End
    Begin
        dbText "Name" ="ItemSeqNumber"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7ec0c899f956e340a8e5378cbf060390
        End
    End
    Begin
        dbText "Name" ="StopLoad"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xfb4adcc5c423294fb80ae3222e6a7700
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
        Right =636
        Bottom =120
        Top =0
        Name ="tblBOMOpSequencesInterfaceTemp"
        Name =""
    End
End

Operation =5
Option =0
Where ="(((tblBOMOpSequencesInterfaceTemp.OperationSeqNum)=[Forms]![frmItemInterface]![f"
    "rmItemInterfaceRouting].[Form]![OperationSeqNumber]))"
Begin InputTables
    Name ="tblBOMOpSequencesInterfaceTemp"
End
Begin OutputColumns
    Expression ="tblBOMOpSequencesInterfaceTemp.OperationSeqNum"
    Expression ="tblBOMOpSequencesInterfaceTemp.*"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x3dfc549042aaf648891d0a5808b38576
End
Begin
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.OperationSeqNumber"
        dbInteger "ColumnWidth" ="2040"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ComponentQuantity"
        dbInteger "ColumnWidth" ="1905"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemSeqNumberNew"
        dbInteger "ColumnWidth" ="1950"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ComponentItemNumber"
        dbInteger "ColumnWidth" ="2220"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OpSeqNumberNew"
        dbInteger "ColumnWidth" ="1950"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
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
    Bottom =222
    Left =0
    Top =0
    ColumnsShown =771
    Begin
        Left =283
        Top =6
        Right =576
        Bottom =120
        Top =0
        Name ="tblBOMOpSequencesInterfaceTemp"
        Name =""
    End
End

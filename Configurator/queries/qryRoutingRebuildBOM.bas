Operation =4
Option =0
Begin InputTables
    Name ="tblBOMInventoryCompsInterface"
    Name ="tblBOMOpSequencesInterfaceTemp"
End
Begin OutputColumns
    Name ="tblBOMInventoryCompsInterface.OperationSeqNumber"
    Expression ="tblBOMOpSequencesInterfaceTemp!OperationSeqNum"
End
Begin Joins
    LeftTable ="tblBOMInventoryCompsInterface"
    RightTable ="tblBOMOpSequencesInterfaceTemp"
    Expression ="tblBOMInventoryCompsInterface.OrgOpSeq = tblBOMOpSequencesInterfaceTemp.OrgOpSeq"
    Flag =1
    LeftTable ="tblBOMInventoryCompsInterface"
    RightTable ="tblBOMOpSequencesInterfaceTemp"
    Expression ="tblBOMInventoryCompsInterface.CopyCount = tblBOMOpSequencesInterfaceTemp.CopyCou"
        "nt"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0xddb0886eec942e449c87b5358eaf5f79
End
Begin
    Begin
        dbText "Name" ="tblBOMOpSequencesInterface.OperationSeqNum"
        dbInteger "ColumnWidth" ="1770"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMOpResourcesInterface.OperationSeqNumber"
        dbInteger "ColumnWidth" ="2040"
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
    Bottom =376
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =312
        Top =6
        Right =550
        Bottom =270
        Top =0
        Name ="tblBOMInventoryCompsInterface"
        Name =""
    End
    Begin
        Left =588
        Top =6
        Right =845
        Bottom =333
        Top =0
        Name ="tblBOMOpSequencesInterfaceTemp"
        Name =""
    End
End

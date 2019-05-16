Operation =4
Option =0
Begin InputTables
    Name ="tblBOMOpResourcesInterface"
    Name ="tblBOMOpSequencesInterfaceTemp"
End
Begin OutputColumns
    Name ="tblBOMOpResourcesInterface.OperationSeqNumber"
    Expression ="tblBOMOpSequencesInterfaceTemp!OperationSeqNum"
End
Begin Joins
    LeftTable ="tblBOMOpResourcesInterface"
    RightTable ="tblBOMOpSequencesInterfaceTemp"
    Expression ="tblBOMOpResourcesInterface.OrgOpSeq = tblBOMOpSequencesInterfaceTemp.OrgOpSeq"
    Flag =1
    LeftTable ="tblBOMOpResourcesInterface"
    RightTable ="tblBOMOpSequencesInterfaceTemp"
    Expression ="tblBOMOpResourcesInterface.CopyCount = tblBOMOpSequencesInterfaceTemp.CopyCount"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbBinary "GUID" = Begin
    0xe46424e42edf34469c897f1b26f63996
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
        Left =354
        Top =10
        Right =591
        Bottom =349
        Top =0
        Name ="tblBOMOpResourcesInterface"
        Name =""
    End
    Begin
        Left =629
        Top =6
        Right =893
        Bottom =316
        Top =0
        Name ="tblBOMOpSequencesInterfaceTemp"
        Name =""
    End
End

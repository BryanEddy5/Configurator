Operation =4
Option =0
Begin InputTables
    Name ="tblBOMOpSequencesInterfaceTemp"
    Name ="tblBOMRoutingsAttachments"
End
Begin OutputColumns
    Name ="tblBOMRoutingsAttachments.OperationSeqNumber"
    Expression ="tblBOMOpSequencesInterfaceTemp!OperationSeqNum"
End
Begin Joins
    LeftTable ="tblBOMOpSequencesInterfaceTemp"
    RightTable ="tblBOMRoutingsAttachments"
    Expression ="tblBOMOpSequencesInterfaceTemp.OrgOpSeq = tblBOMRoutingsAttachments.OrgOpSeq"
    Flag =1
    LeftTable ="tblBOMOpSequencesInterfaceTemp"
    RightTable ="tblBOMRoutingsAttachments"
    Expression ="tblBOMOpSequencesInterfaceTemp.CopyCount = tblBOMRoutingsAttachments.CopyCount"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x6c180097f796ee4fb58781f15bf453b3
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
        Right =569
        Bottom =280
        Top =0
        Name ="tblBOMOpSequencesInterfaceTemp"
        Name =""
    End
    Begin
        Left =629
        Top =6
        Right =857
        Bottom =180
        Top =0
        Name ="tblBOMRoutingsAttachments"
        Name =""
    End
End

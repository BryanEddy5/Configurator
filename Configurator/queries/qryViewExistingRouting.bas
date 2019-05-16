Operation =1
Option =0
Begin InputTables
    Name ="tblBOMOpSequencesInterfaceTemp"
    Name ="tblBOMRoutingsAttachments"
End
Begin OutputColumns
    Expression ="tblBOMOpSequencesInterfaceTemp.OperationSeqNum"
    Expression ="tblBOMOpSequencesInterfaceTemp.DepartmentCode"
    Expression ="tblBOMOpSequencesInterfaceTemp.OperationDescription"
    Expression ="tblBOMRoutingsAttachments.ShortText"
    Expression ="tblBOMRoutingsAttachments.SeqNumber"
    Expression ="tblBOMOpSequencesInterfaceTemp.AssemblyItemNumber"
End
Begin Joins
    LeftTable ="tblBOMOpSequencesInterfaceTemp"
    RightTable ="tblBOMRoutingsAttachments"
    Expression ="tblBOMOpSequencesInterfaceTemp.OperationSeqNum = tblBOMRoutingsAttachments.Opera"
        "tionSeqNumber"
    Flag =2
    LeftTable ="tblBOMOpSequencesInterfaceTemp"
    RightTable ="tblBOMRoutingsAttachments"
    Expression ="tblBOMOpSequencesInterfaceTemp.AssemblyItemNumber = tblBOMRoutingsAttachments.As"
        "semblyItemNumber"
    Flag =2
    LeftTable ="tblBOMOpSequencesInterfaceTemp"
    RightTable ="tblBOMRoutingsAttachments"
    Expression ="tblBOMOpSequencesInterfaceTemp.OrganizationCode = tblBOMRoutingsAttachments.Orga"
        "nizationCode"
    Flag =2
End
Begin OrderBy
    Expression ="tblBOMOpSequencesInterfaceTemp.OperationSeqNum"
    Flag =0
    Expression ="tblBOMRoutingsAttachments.SeqNumber"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x7156092c7e006244b9dfcf23f55db426
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblBOMOpSequencesInterfaceTemp.DepartmentCode"
        dbInteger "ColumnWidth" ="1635"
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
    Bottom =383
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =531
        Top =47
        Right =815
        Bottom =296
        Top =0
        Name ="tblBOMOpSequencesInterfaceTemp"
        Name =""
    End
    Begin
        Left =49
        Top =158
        Right =255
        Bottom =332
        Top =0
        Name ="tblBOMRoutingsAttachments"
        Name =""
    End
End

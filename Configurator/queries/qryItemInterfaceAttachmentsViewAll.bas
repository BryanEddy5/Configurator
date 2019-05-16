Operation =1
Option =0
Begin InputTables
    Name ="tblBOMRoutingsAttachments"
    Name ="tblBOMOpSequencesInterfaceTemp"
End
Begin OutputColumns
    Expression ="tblBOMRoutingsAttachments.*"
    Expression ="tblBOMOpSequencesInterfaceTemp.DepartmentCode"
    Expression ="tblBOMOpSequencesInterfaceTemp.OperationDescription"
End
Begin Joins
    LeftTable ="tblBOMRoutingsAttachments"
    RightTable ="tblBOMOpSequencesInterfaceTemp"
    Expression ="tblBOMRoutingsAttachments.OrgOpSeq = tblBOMOpSequencesInterfaceTemp.OrgOpSeq"
    Flag =3
    LeftTable ="tblBOMRoutingsAttachments"
    RightTable ="tblBOMOpSequencesInterfaceTemp"
    Expression ="tblBOMRoutingsAttachments.CopyCount = tblBOMOpSequencesInterfaceTemp.CopyCount"
    Flag =3
    LeftTable ="tblBOMRoutingsAttachments"
    RightTable ="tblBOMOpSequencesInterfaceTemp"
    Expression ="tblBOMRoutingsAttachments.OperationSeqNumber = tblBOMOpSequencesInterfaceTemp.Op"
        "erationSeqNum"
    Flag =3
    LeftTable ="tblBOMRoutingsAttachments"
    RightTable ="tblBOMOpSequencesInterfaceTemp"
    Expression ="tblBOMRoutingsAttachments.AssemblyItemNumber = tblBOMOpSequencesInterfaceTemp.As"
        "semblyItemNumber"
    Flag =3
    LeftTable ="tblBOMRoutingsAttachments"
    RightTable ="tblBOMOpSequencesInterfaceTemp"
    Expression ="tblBOMRoutingsAttachments.OrganizationCode = tblBOMOpSequencesInterfaceTemp.Orga"
        "nizationCode"
    Flag =3
End
Begin OrderBy
    Expression ="tblBOMOpSequencesInterfaceTemp.OperationSeqNum"
    Flag =0
    Expression ="tblBOMRoutingsAttachments.OperationSeqNumber"
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
    0x978842c6e78c684caf09124606a399f6
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
End
Begin
    State =0
    Left =0
    Top =40
    Right =1560
    Bottom =854
    Left =-1
    Top =-1
    Right =1528
    Bottom =380
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =39
        Right =254
        Bottom =286
        Top =0
        Name ="tblBOMRoutingsAttachments"
        Name =""
    End
    Begin
        Left =613
        Top =6
        Right =877
        Bottom =292
        Top =0
        Name ="tblBOMOpSequencesInterfaceTemp"
        Name =""
    End
End

Operation =4
Option =0
Where ="(((tblBOMOpSequencesInterfaceTemp.OrgOpSeq)=[Forms]![frmItemInterface]![frmItemI"
    "nterfaceRouting].[Form]![OrgOpSeq]))"
Begin InputTables
    Name ="tblBOMOpSequencesInterfaceTemp"
    Name ="tblBOMRoutingsAttachments"
End
Begin OutputColumns
    Name ="tblBOMRoutingsAttachments.OperationSeqNumber"
    Expression ="Forms!frmItemInterface!frmItemInterfaceRouting.Form!OperationSeqNumber"
End
Begin Joins
    LeftTable ="tblBOMOpSequencesInterfaceTemp"
    RightTable ="tblBOMRoutingsAttachments"
    Expression ="tblBOMOpSequencesInterfaceTemp.OrgOpSeq = tblBOMRoutingsAttachments.OrgOpSeq"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x548d25cd20cff649a6e249c67a0b2725
End
Begin
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
    Bottom =293
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =38
        Top =6
        Right =261
        Bottom =276
        Top =0
        Name ="tblBOMOpSequencesInterfaceTemp"
        Name =""
    End
    Begin
        Left =841
        Top =11
        Right =1060
        Bottom =185
        Top =0
        Name ="tblBOMRoutingsAttachments"
        Name =""
    End
End

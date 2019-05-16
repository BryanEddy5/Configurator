Operation =4
Option =0
Where ="(((tblBOMOpSequencesInterfaceTemp.OrgOpSeq)=[Forms]![frmItemInterface]![frmItemI"
    "nterfaceRouting].[Form]![OrgOpSeq]))"
Begin InputTables
    Name ="tblBOMOpSequencesInterfaceTemp"
    Name ="tblBOMOpResourcesInterface"
End
Begin OutputColumns
    Name ="tblBOMOpResourcesInterface.OperationSeqNumber"
    Expression ="Forms!frmItemInterface!frmItemInterfaceRouting.Form!OperationSeqNumber"
End
Begin Joins
    LeftTable ="tblBOMOpSequencesInterfaceTemp"
    RightTable ="tblBOMOpResourcesInterface"
    Expression ="tblBOMOpSequencesInterfaceTemp.OrgOpSeq = tblBOMOpResourcesInterface.OrgOpSeq"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x1e504bdc7d2fd646a622033613091423
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
        Right =304
        Bottom =262
        Top =0
        Name ="tblBOMOpSequencesInterfaceTemp"
        Name =""
    End
    Begin
        Left =820
        Top =12
        Right =1047
        Bottom =276
        Top =0
        Name ="tblBOMOpResourcesInterface"
        Name =""
    End
End

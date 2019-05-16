Operation =4
Option =0
Where ="(((tblBOMOpSequencesInterfaceTemp.OrgOpSeq)=[Forms]![frmItemInterface]![frmItemI"
    "nterfaceRouting].[Form]![OrgOpSeq]))"
Begin InputTables
    Name ="tblBOMOpSequencesInterfaceTemp"
    Name ="tblBOMInventoryCompsInterface"
End
Begin OutputColumns
    Name ="tblBOMInventoryCompsInterface.OperationSeqNumber"
    Expression ="Forms!frmItemInterface!frmItemInterfaceRouting.Form!OperationSeqNumber"
End
Begin Joins
    LeftTable ="tblBOMOpSequencesInterfaceTemp"
    RightTable ="tblBOMInventoryCompsInterface"
    Expression ="tblBOMOpSequencesInterfaceTemp.OrgOpSeq = tblBOMInventoryCompsInterface.OrgOpSeq"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0xceef2830903fc54cb68fda9a3a140b1f
End
Begin
End
Begin
    State =0
    Left =134
    Top =156
    Right =1243
    Bottom =659
    Left =-1
    Top =-1
    Right =1077
    Bottom =293
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =356
        Top =6
        Right =605
        Bottom =255
        Top =0
        Name ="tblBOMOpSequencesInterfaceTemp"
        Name =""
    End
    Begin
        Left =38
        Top =6
        Right =318
        Bottom =270
        Top =0
        Name ="tblBOMInventoryCompsInterface"
        Name =""
    End
End

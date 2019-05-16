Operation =4
Option =0
Where ="(((tblBOMOpSequencesInterface.OperationCode) Like \"J*\"))"
Begin InputTables
    Name ="tblBOMOpSequencesInterface"
End
Begin OutputColumns
    Name ="tblBOMOpSequencesInterface.OperationCode"
    Expression ="[Forms]![frmItemInterface]![frmPremiseItemDetails].[Form]![Routing1]"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x95223527ea90ca4d896f589c81c46490
End
Begin
    Begin
        dbText "Name" ="tblBOMOpSequencesInterface.OperationCode"
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
    Bottom =149
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =124
        Top =0
        Name ="tblBOMOpSequencesInterface"
        Name =""
    End
End

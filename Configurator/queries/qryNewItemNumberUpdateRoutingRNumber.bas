Operation =4
Option =0
Where ="(((tblBOMOpSequencesInterface.OperationCode) Like \"r*\"))"
Begin InputTables
    Name ="tblBOMOpSequencesInterface"
End
Begin OutputColumns
    Name ="tblBOMOpSequencesInterface.OperationCode"
    Expression ="Forms!frmItemInterface!frmPremiseItemDetails.Form!Routing1"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0xf2b7a010cc35974d9df14e68f3cf37d7
End
Begin
End
Begin
    State =0
    Left =154
    Top =394
    Right =1263
    Bottom =774
    Left =-1
    Top =-1
    Right =1077
    Bottom =166
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =38
        Top =6
        Right =302
        Bottom =178
        Top =0
        Name ="tblBOMOpSequencesInterface"
        Name =""
    End
End

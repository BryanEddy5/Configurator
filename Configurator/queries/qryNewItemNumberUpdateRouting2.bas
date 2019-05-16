Operation =4
Option =0
Where ="(((tblBOMOpSequencesInterface.OperationCode)=[Forms]![frmItemInterface]![frmPrem"
    "iseItemDetailsCopied].[Form]![Routing2]))"
Begin InputTables
    Name ="tblBOMOpSequencesInterface"
End
Begin OutputColumns
    Name ="tblBOMOpSequencesInterface.OperationCode"
    Expression ="Forms!frmItemInterface!frmPremiseItemDetails.Form!Routing2"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x19b639d9778aff47b9608575e8a14400
End
Begin
End
Begin
    State =0
    Left =119
    Top =31
    Right =1228
    Bottom =411
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
        Right =134
        Bottom =124
        Top =0
        Name ="tblBOMOpSequencesInterface"
        Name =""
    End
End

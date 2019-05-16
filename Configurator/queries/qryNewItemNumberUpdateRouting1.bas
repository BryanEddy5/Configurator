Operation =4
Option =0
Where ="(((tblBOMOpSequencesInterface.OperationCode)=[Forms]![frmItemInterface]![frmPrem"
    "iseItemDetailsCopied].[Form]![Routing1]))"
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
    0xe063c5ab4f1f1c438d88ae69d1d8d1d9
End
Begin
End
Begin
    State =0
    Left =154
    Top =420
    Right =1263
    Bottom =800
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

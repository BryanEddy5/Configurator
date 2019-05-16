Operation =4
Option =0
Where ="(((tblBOMInventoryCompsInterface.ComponentItemNumber) Like Left([Forms]![frmItem"
    "Interface]![frmPremiseItemDetailsCopied].[Form]![OracleItemNumber],2) & \"*\"))"
Begin InputTables
    Name ="tblBOMInventoryCompsInterface"
End
Begin OutputColumns
    Name ="tblBOMInventoryCompsInterface.ComponentItemNumber"
    Expression ="Left(Forms!frmItemInterface!frmPremiseItemDetails.Form!OracleItemNumber,10)"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x0f7e38a3648e214aaf975afcab41330b
End
Begin
End
Begin
    State =0
    Left =249
    Top =448
    Right =1358
    Bottom =772
    Left =-1
    Top =-1
    Right =1077
    Bottom =110
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =308
        Top =6
        Right =526
        Bottom =124
        Top =0
        Name ="tblBOMInventoryCompsInterface"
        Name =""
    End
End

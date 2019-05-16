Operation =4
Option =0
Where ="(((tblBOMInventoryCompsInterface.ComponentItemNumber) Like Left([Forms]![frmItem"
    "Interface]![frmPremiseItemDetailsCopied].[Form]![OracleItemNumber],2) & \"*\"))"
Begin InputTables
    Name ="tblBOMInventoryCompsInterface"
End
Begin OutputColumns
    Name ="tblBOMInventoryCompsInterface.ComponentItemNumber"
    Expression ="Left([Forms]![frmItemInterface]![frmPremiseItemDetails].[Form]![OracleItemNumber"
        "],7) & \"-00\""
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0xf7ccfa06fc862340a6a41fdc315f0d2d
End
Begin
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.ComponentItemNumber"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =127
    Top =401
    Right =1236
    Bottom =725
    Left =-1
    Top =-1
    Right =1077
    Bottom =93
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

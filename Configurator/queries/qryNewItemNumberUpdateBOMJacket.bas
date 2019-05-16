Operation =4
Option =0
Where ="(((tblBOMInventoryCompsInterface.ComponentItemNumber)=[Forms]![frmItemInterface]"
    "![frmPremiseItemDetailsCopied].[Form]![JacketMaterial]))"
Begin InputTables
    Name ="tblBOMInventoryCompsInterface"
End
Begin OutputColumns
    Name ="tblBOMInventoryCompsInterface.ComponentItemNumber"
    Expression ="Forms!frmItemInterface!frmPremiseItemDetails.Form!JacketMaterial"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x32000df0cb71ba4b9affaa8f122ef83f
End
Begin
End
Begin
    State =0
    Left =119
    Top =228
    Right =1228
    Bottom =552
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

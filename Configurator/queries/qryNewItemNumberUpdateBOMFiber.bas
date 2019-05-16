Operation =4
Option =0
Where ="(((tblBOMInventoryCompsInterface.ComponentItemNumber) Like [Forms]![frmItemInter"
    "face]![frmPremiseItemDetailsCopied].[Form]![FiberType1] & \"*\" Or (tblBOMInvent"
    "oryCompsInterface.ComponentItemNumber) Like [Forms]![frmItemInterface]![frmPremi"
    "seItemDetailsCopied].[Form]![FiberType2] & \"*\" Or (tblBOMInventoryCompsInterfa"
    "ce.ComponentItemNumber) Like [Forms]![frmItemInterface]![frmPremiseItemDetailsCo"
    "pied].[Form]![FiberType3] & \"*\"))"
Begin InputTables
    Name ="tblBOMInventoryCompsInterface"
End
Begin OutputColumns
    Name ="tblBOMInventoryCompsInterface.ComponentItemNumber"
    Expression ="Forms!frmItemInterface!frmPremiseItemDetails.Form!FiberType1 & Mid([ComponentIte"
        "mNumber],9,7)"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x30d53e97c3a5a24fb0295b9cc7f29d66
End
Begin
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.ComponentItemNumber"
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
    Bottom =77
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =154
        Top =22
        Right =372
        Bottom =140
        Top =0
        Name ="tblBOMInventoryCompsInterface"
        Name =""
    End
End

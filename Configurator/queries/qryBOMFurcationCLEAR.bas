Operation =5
Option =1
Where ="(((tblBOMInventoryCompsInterface.ComponentItemNumber) Like \"pt*\" Or (tblBOMInv"
    "entoryCompsInterface.ComponentItemNumber) Like \"frc*\"))"
Begin InputTables
    Name ="tblBOMInventoryCompsInterface"
End
Begin OutputColumns
    Expression ="tblBOMInventoryCompsInterface.ComponentItemNumber"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x4ffacd7ed9b5574fb826faca9802e6a3
End
Begin
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.[ComponentItemNumber]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.AssemblyItemNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.OperationSeqNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.AlternateBOMDesignator"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.ItemCopied"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.OrgOpSeq"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.OrganizationCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.ComponentQuantity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.CountPerUOM"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.ItemSeqNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.ComponentItemNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.EffectivityDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.LayerID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.CopyCount"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =51
    Top =257
    Right =1552
    Bottom =944
    Left =-1
    Top =-1
    Right =1469
    Bottom =257
    Left =0
    Top =0
    ColumnsShown =771
    Begin
        Left =43
        Top =3
        Right =365
        Bottom =303
        Top =0
        Name ="tblBOMInventoryCompsInterface"
        Name =""
    End
End

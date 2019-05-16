Operation =1
Option =0
Begin InputTables
    Name ="tblBOMInventoryCompsInterface"
End
Begin OutputColumns
    Expression ="tblBOMInventoryCompsInterface.*"
    Alias ="StoresPrefix"
    Expression ="Left(tblBOMInventoryCompsInterface!ComponentItemNumber,4)"
    Alias ="StoresSuffix"
    Expression ="Right(tblBOMInventoryCompsInterface!ComponentItemNumber,2)"
    Alias ="PTCheck"
    Expression ="IIf(Left([ComponentItemNumber],2)=\"PT\",1,0)"
End
Begin OrderBy
    Expression ="tblBOMInventoryCompsInterface.OperationSeqNumber"
    Flag =0
    Expression ="tblBOMInventoryCompsInterface.ItemSeqNumber"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x87883a5217ceea45b378e72184489bfc
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="StoresPrefix"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf854fdd382a2db47a5885930714386a8
        End
    End
    Begin
        dbText "Name" ="StoresSuffix"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x23ecb5f30057e04095c72ce303869dc3
        End
    End
    Begin
        dbText "Name" ="PTCheck"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6f71f3ced3a83d47829d5ca26f3f7be9
        End
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1560
    Bottom =854
    Left =-1
    Top =-1
    Right =1528
    Bottom =264
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =59
        Top =18
        Right =346
        Bottom =226
        Top =0
        Name ="tblBOMInventoryCompsInterface"
        Name =""
    End
End

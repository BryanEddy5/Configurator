Operation =1
Option =0
Where ="(((tblBOMInventoryCompsInterface.OperationSeqNumber)=[Forms]![frmMainMenu]![Oper"
    "ationSequence]))"
Begin InputTables
    Name ="tblBOMInventoryCompsInterface"
End
Begin OutputColumns
    Expression ="tblBOMInventoryCompsInterface.*"
    Alias ="StoresPrefix"
    Expression ="Left(tblBOMInventoryCompsInterface!ComponentItemNumber,4)"
    Alias ="StoresSuffix"
    Expression ="Right(tblBOMInventoryCompsInterface!ComponentItemNumber,2)"
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
    0x0622cbdefad901408ea3acfa253275a1
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="StoresPrefix"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x00be745e8f02e6419c22b0dddbc3838b
        End
    End
    Begin
        dbText "Name" ="StoresSuffix"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x10c53afaf0f9cb4e9903cfcc220450f4
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
        Bottom =237
        Top =0
        Name ="tblBOMInventoryCompsInterface"
        Name =""
    End
End

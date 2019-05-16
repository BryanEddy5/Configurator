Operation =1
Option =0
Begin InputTables
    Name ="tblBOMInventoryCompsInterface"
    Name ="tblOracleItems"
End
Begin OutputColumns
    Expression ="tblBOMInventoryCompsInterface.*"
    Alias ="Expr1"
    Expression ="tblOracleItems.ItemDescription"
End
Begin Joins
    LeftTable ="tblBOMInventoryCompsInterface"
    RightTable ="tblOracleItems"
    Expression ="tblBOMInventoryCompsInterface.ComponentItemNumber = tblOracleItems.ItemNumber"
    Flag =2
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
    0x3feb7b48a1d10949b4c95571748097cb
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x521d949b38a9c044ab7536c5659b4b82
        End
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
    Begin
        Left =384
        Top =6
        Right =480
        Bottom =90
        Top =0
        Name ="tblOracleItems"
        Name =""
    End
End

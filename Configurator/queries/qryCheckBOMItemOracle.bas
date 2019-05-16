Operation =1
Option =0
Where ="(((tblOracleItems.ItemNumber) Is Null))"
Begin InputTables
    Name ="tblBOMInventoryCompsInterface"
    Name ="tblOracleItems"
End
Begin OutputColumns
    Expression ="tblBOMInventoryCompsInterface.ItemSeqNumber"
    Expression ="tblBOMInventoryCompsInterface.ComponentItemNumber"
    Expression ="tblBOMInventoryCompsInterface.OperationSeqNumber"
    Expression ="tblOracleItems.ItemNumber"
    Alias ="CheckCode"
    Expression ="IIf([ComponentItemNumber]=\"0\",\"BOM - Missing Component Item\",\"BOM - Check C"
        "omponent: \" & [ComponentItemNumber])"
    Alias ="StopLoad"
    Expression ="0"
End
Begin Joins
    LeftTable ="tblBOMInventoryCompsInterface"
    RightTable ="tblOracleItems"
    Expression ="tblBOMInventoryCompsInterface.ComponentItemNumber = tblOracleItems.ItemNumber"
    Flag =2
End
Begin OrderBy
    Expression ="tblBOMInventoryCompsInterface.ItemSeqNumber"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbBinary "GUID" = Begin
    0x277ca1f221c3e84c8f161638b66cbc47
End
Begin
    Begin
        dbText "Name" ="CheckCode"
        dbInteger "ColumnWidth" ="4245"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc3bb73539d90de4dbfdd88ab0ec38d6e
        End
    End
    Begin
        dbText "Name" ="StopLoad"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x724f864039017f4487496fdb272ec9ff
        End
    End
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.ItemSeqNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.ComponentItemNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.OperationSeqNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblOracleItems.ItemNumber"
        dbLong "AggregateType" ="-1"
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
    Bottom =93
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =270
        Bottom =120
        Top =0
        Name ="tblBOMInventoryCompsInterface"
        Name =""
    End
    Begin
        Left =465
        Top =10
        Right =561
        Bottom =94
        Top =0
        Name ="tblOracleItems"
        Name =""
    End
End

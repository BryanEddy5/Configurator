Operation =1
Option =0
Begin InputTables
    Name ="tblOracleBOMInventoryCompsInterface"
End
Begin OutputColumns
    Expression ="tblOracleBOMInventoryCompsInterface.*"
End
Begin OrderBy
    Expression ="tblOracleBOMInventoryCompsInterface.OperationSeqNumber"
    Flag =0
    Expression ="tblOracleBOMInventoryCompsInterface.ItemSeqNumber"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x4b7ccb268278224ea08a79fa7d318c21
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
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
    Bottom =110
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =325
        Bottom =120
        Top =0
        Name ="tblOracleBOMInventoryCompsInterface"
        Name =""
    End
End

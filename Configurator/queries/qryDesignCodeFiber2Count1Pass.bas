Operation =1
Option =16
RowCount ="1"
Where ="(((tblFiberCount.FiberCount)<=Int([Forms]![frmDesignCode]![FiberCount])-Int([For"
    "ms]![frmDesignCode]![Fiber1])))"
Begin InputTables
    Name ="tblFiberCount"
End
Begin OutputColumns
    Expression ="tblFiberCount.FiberCountString"
    Expression ="tblFiberCount.FiberCount"
End
Begin OrderBy
    Expression ="tblFiberCount.FiberCount"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x5876620ced7af54e88ebc8b6e8bb778b
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblFiberCount.FiberCountString"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =169
    Top =177
    Right =1877
    Bottom =501
    Left =-1
    Top =-1
    Right =1676
    Bottom =76
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =275
        Top =6
        Right =496
        Bottom =115
        Top =0
        Name ="tblFiberCount"
        Name =""
    End
End

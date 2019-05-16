Operation =1
Option =2
Where ="(((tblFiberCount.FiberCount)<=Int([Forms]![frmDesignCode]![FiberCount])))"
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
    0x820dd106643e164a86875338191e0aa6
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
End
Begin
    State =0
    Left =10
    Top =33
    Right =1709
    Bottom =363
    Left =-1
    Top =-1
    Right =1667
    Bottom =99
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =206
        Bottom =94
        Top =0
        Name ="tblFiberCount"
        Name =""
    End
End

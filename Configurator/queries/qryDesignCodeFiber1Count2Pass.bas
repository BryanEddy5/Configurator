Operation =1
Option =2
Where ="(((tblFiberCount.FiberCount)<=Int(Forms!frmDesignCode!FiberCount)) And (([FiberC"
    "ount] Mod Forms!frmdesigncode!SubFiberCount)=0))"
Begin InputTables
    Name ="tblFiberCount"
End
Begin OutputColumns
    Expression ="tblFiberCount.FiberCountString"
    Expression ="tblFiberCount.FiberCount"
    Alias ="Divisible"
    Expression ="[FiberCount] Mod [Forms]![frmdesigncode]![SubFiberCount]"
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
    0xa958106471e65b49a4b82fa35b04a3b8
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Divisible"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x1008df217bdb854dbce68c8614b8a5af
        End
    End
    Begin
        dbText "Name" ="tblFiberCount.FiberCountString"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblFiberCount.FiberCount"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =251
    Top =52
    Right =1950
    Bottom =523
    Left =-1
    Top =-1
    Right =1667
    Bottom =108
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =229
        Bottom =137
        Top =0
        Name ="tblFiberCount"
        Name =""
    End
End

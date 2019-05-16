Operation =1
Option =18
RowCount ="1"
Where ="(((tblFiberCount.FiberCount)<=Int(Forms!frmDesignCode!FiberCount)-Int(Forms!frmD"
    "esignCode!Fiber1)) And (([FiberCount] Mod Forms!frmdesigncode!SubFiberCount)=0))"
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
    0x8ba6e4add825ab45900c8cfc483b89f9
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Divisible"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x05b05c6444591a42b04cde596767ee29
        End
    End
    Begin
        dbText "Name" ="tblFiberCount.FiberCountString"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =75
    Top =-11
    Right =1783
    Bottom =313
    Left =-1
    Top =-1
    Right =1676
    Bottom =59
    Left =1677
    Top =0
    ColumnsShown =539
    Begin
        Left =-1402
        Top =6
        Right =-1306
        Bottom =94
        Top =0
        Name ="tblFiberCount"
        Name =""
    End
End

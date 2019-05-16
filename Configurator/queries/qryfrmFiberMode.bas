Operation =1
Option =0
Where ="(((tbl_Fibers.Fiber)=[Forms]![frmCreateItem]![Fiber]))"
Begin InputTables
    Name ="tbl_Fibers"
End
Begin OutputColumns
    Expression ="tbl_Fibers.Fiber"
    Expression ="tbl_Fibers.[OM#]"
    Expression ="tbl_Fibers.[SM/MM]"
    Expression ="tbl_Fibers.BIF"
    Expression ="tbl_Fibers.ITU"
    Expression ="tbl_Fibers.Vendor"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x8a41d4fe78d3044bb57769c43bc5b415
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
End
Begin
    State =0
    Left =388
    Top =275
    Right =1819
    Bottom =684
    Left =-1
    Top =-1
    Right =1399
    Bottom =178
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =296
        Bottom =94
        Top =0
        Name ="tbl_Fibers"
        Name =""
    End
End

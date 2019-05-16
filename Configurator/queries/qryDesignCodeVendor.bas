Operation =1
Option =0
Having ="(((tbl_Fibers.ITU)=[Forms]![frmDesignCodeSpecialFiber]![ITU]))"
Begin InputTables
    Name ="tbl_Fibers"
End
Begin OutputColumns
    Expression ="tbl_Fibers.Vendor"
    Expression ="tbl_Fibers.Fiber_Type"
    Expression ="tbl_Fibers.ITU"
End
Begin Groups
    Expression ="tbl_Fibers.Vendor"
    GroupLevel =0
    Expression ="tbl_Fibers.Fiber_Type"
    GroupLevel =0
    Expression ="tbl_Fibers.ITU"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xd86fa56d1e578246830dd929f5bbeacc
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
End
Begin
    State =0
    Left =125
    Top =300
    Right =1781
    Bottom =696
    Left =-1
    Top =-1
    Right =1624
    Bottom =165
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =230
        Bottom =184
        Top =0
        Name ="tbl_Fibers"
        Name =""
    End
End

Operation =1
Option =0
Where ="(((tbl_Fibers.Fiber)=[Forms]![frmCreateItem]![Fiber2]))"
Begin InputTables
    Name ="tbl_Fibers"
End
Begin OutputColumns
    Expression ="tbl_Fibers.BIF"
    Expression ="tbl_Fibers.Fiber"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xb1cb0169867bda4bb7d8790f3bf54b70
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
End
Begin
    State =0
    Left =199
    Top =156
    Right =1811
    Bottom =480
    Left =-1
    Top =-1
    Right =1580
    Bottom =93
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =124
        Top =0
        Name ="tbl_Fibers"
        Name =""
    End
End

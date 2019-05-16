Operation =1
Option =2
Where ="(((qry_SearchByFiberType.Fiber)=[Forms]![frmCreateItem]![Fiber]))"
Begin InputTables
    Name ="qry_SearchByFiberType"
End
Begin OutputColumns
    Expression ="qry_SearchByFiberType.[OM#]"
    Expression ="qry_SearchByFiberType.BIF"
    Expression ="qry_SearchByFiberType.Fiber"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x94d716562b8df14fbda96527b8bbd0b1
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
End
Begin
    State =0
    Left =147
    Top =153
    Right =1578
    Bottom =547
    Left =-1
    Top =-1
    Right =1399
    Bottom =163
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =261
        Bottom =154
        Top =0
        Name ="qry_SearchByFiberType"
        Name =""
    End
End

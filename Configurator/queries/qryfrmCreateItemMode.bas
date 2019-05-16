Operation =1
Option =8
Where ="(((qry_SearchByFiberType.Fiber)=[Forms]![frmCreateItem]![Fiber]))"
Begin InputTables
    Name ="qry_SearchByFiberType"
End
Begin OutputColumns
    Expression ="qry_SearchByFiberType.[OM#]"
    Expression ="qry_SearchByFiberType.Fiber"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xf06813346d872f4289323cadf9c39e39
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
    Right =1579
    Bottom =881
    Left =-1
    Top =-1
    Right =1547
    Bottom =235
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =412
        Bottom =244
        Top =0
        Name ="qry_SearchByFiberType"
        Name =""
    End
End

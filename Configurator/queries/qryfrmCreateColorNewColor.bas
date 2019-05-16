Operation =1
Option =0
Where ="(((qryfrmCreateColorNew.[Jacket Material])=[Forms]![frmCreateItem]![Jacket]))"
Begin InputTables
    Name ="qryfrmCreateColorNew"
End
Begin OutputColumns
    Expression ="qryfrmCreateColorNew.[Color Chip ID]"
    Expression ="qryfrmCreateColorNew.[Jacket Color]"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xe7129d0e24d40742b04823204057bd40
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
End
Begin
    State =0
    Left =38
    Top =193
    Right =1875
    Bottom =656
    Left =-1
    Top =-1
    Right =1805
    Bottom =232
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =385
        Bottom =154
        Top =0
        Name ="qryfrmCreateColorNew"
        Name =""
    End
End

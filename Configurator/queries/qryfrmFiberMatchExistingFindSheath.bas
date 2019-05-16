Operation =1
Option =0
Where ="((([Basic (Cable/Sheath)].[New Oracle Part Number]) Like Left([Forms]![frmFiberM"
    "atchExisting]![Oracle],7) & \"*\"))"
Begin InputTables
    Name ="Basic (Cable/Sheath)"
End
Begin OutputColumns
    Expression ="[Basic (Cable/Sheath)].*"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x8695ca6fe3424c4786e5297d4e02633d
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
    Bottom =93
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =124
        Top =0
        Name ="Basic (Cable/Sheath)"
        Name =""
    End
End

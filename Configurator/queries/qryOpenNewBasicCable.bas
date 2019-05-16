Operation =1
Option =0
Where ="(((Left([New Oracle Part Number],7))=Left([Forms]![frmCreateItem]![ItemNo],7)))"
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
    0x3f9c4ddfc1ffc6429f5525bc82170b49
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
End
Begin
    State =0
    Left =147
    Top =75
    Right =1256
    Bottom =492
    Left =-1
    Top =-1
    Right =1077
    Bottom =186
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =344
        Bottom =199
        Top =0
        Name ="Basic (Cable/Sheath)"
        Name =""
    End
End

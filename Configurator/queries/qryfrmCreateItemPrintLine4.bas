Operation =1
Option =0
Where ="(((tblLine4PrintCode.Priority)=1 Or (tblLine4PrintCode.Priority)=2))"
Begin InputTables
    Name ="tblLine4PrintCode"
End
Begin OutputColumns
    Expression ="tblLine4PrintCode.Line4Print"
    Expression ="tblLine4PrintCode.Priority"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x6725506c79cb1c4fa551102b218c4ef3
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblLine4PrintCode.Line4Print"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblLine4PrintCode.Priority"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =124
    Top =76
    Right =1618
    Bottom =884
    Left =-1
    Top =-1
    Right =1462
    Bottom =457
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =344
        Top =72
        Right =768
        Bottom =331
        Top =0
        Name ="tblLine4PrintCode"
        Name =""
    End
End

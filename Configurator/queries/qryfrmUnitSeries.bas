Operation =1
Option =0
Where ="((([Unit Series].[Unit Color Series])=[Forms]![frmUnitSeries]![NewUnit]))"
Begin InputTables
    Name ="Unit Series"
End
Begin OutputColumns
    Alias ="UnitSeries"
    Expression ="[Unit Series].[Unit Color Series]"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x4d92ac91310f7d469322e510309cb274
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="UnitSeries"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x52f53c7e4b47114cac14adddc9314195
        End
    End
End
Begin
    State =0
    Left =282
    Top =190
    Right =1527
    Bottom =592
    Left =-1
    Top =-1
    Right =1213
    Bottom =175
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =234
        Bottom =184
        Top =0
        Name ="Unit Series"
        Name =""
    End
End

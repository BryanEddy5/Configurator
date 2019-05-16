Operation =3
Name ="tblUnitSeriesCopy"
Option =0
Where ="(((tblUnitSeriesTemplate.UnitNumber)=0) AND ((tblUnitSeriesTemplate.UnitIDTypeNu"
    "mber)=1))"
Begin InputTables
    Name ="tblUnitSeriesTemplate"
End
Begin OutputColumns
    Name ="Unit ID"
    Expression ="tblUnitSeriesTemplate.UnitID"
    Name ="Print Spacing"
    Expression ="tblUnitSeriesTemplate.PrintSpacing"
    Name ="Print type"
    Expression ="tblUnitSeriesTemplate.PrintType"
    Name ="TB Color Series"
    Expression ="tblUnitSeriesTemplate.UnitIDType"
    Alias ="ColorSeries"
    Name ="Unit Color Series"
    Expression ="[Forms]![frmCreateItem]![UnitSeries]"
    Name ="UnitNumber"
    Expression ="tblUnitSeriesTemplate.UnitNumber"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x5ec0b64385ec6649a52e90ee07912bf1
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbBoolean "UseTransaction" ="-1"
Begin
    Begin
        dbText "Name" ="tblUnitSeriesTemplate.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblUnitSeriesTemplate.PrintSpacing"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblUnitSeriesTemplate.PrintType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblUnitSeriesTemplate.UnitIDType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblUnitSeriesTemplate.UnitNumber"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2370"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tblUnitSeriesTemplate.UnitIDTypeNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1005"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ColorSeries"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x00000000000000000000000001000000
        End
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-38
    Top =214
    Right =1071
    Bottom =843
    Left =-1
    Top =-1
    Right =1077
    Bottom =401
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =48
        Top =12
        Right =526
        Bottom =334
        Top =0
        Name ="tblUnitSeriesTemplate"
        Name =""
    End
End

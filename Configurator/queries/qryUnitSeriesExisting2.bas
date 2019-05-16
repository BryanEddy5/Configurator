Operation =1
Option =2
Begin InputTables
    Name ="Unit Series"
    Name ="qryfrmUnitSeries"
End
Begin OutputColumns
    Alias ="Existing"
    Expression ="IIf([UnitSeries] Is Null,0,-1)"
End
Begin Joins
    LeftTable ="Unit Series"
    RightTable ="qryfrmUnitSeries"
    Expression ="[Unit Series].[Unit Color Series] = qryfrmUnitSeries.UnitSeries"
    Flag =2
End
Begin OrderBy
    Expression ="IIf([UnitSeries] Is Null,0,-1)"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xf2130464a9b8fe4d81c0128f9387ad8d
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Existing"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x20ee3ac8ee823c49b198c20056304ee0
        End
    End
End
Begin
    State =0
    Left =141
    Top =306
    Right =1250
    Bottom =628
    Left =-1
    Top =-1
    Right =1077
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
        Name ="Unit Series"
        Name =""
    End
    Begin
        Left =325
        Top =21
        Right =421
        Bottom =94
        Top =0
        Name ="qryfrmUnitSeries"
        Name =""
    End
End

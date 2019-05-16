Operation =1
Option =0
Where ="((([Basic Product Construction].Active)<>0))"
Begin InputTables
    Name ="qryPremiseItemDetails"
    Name ="tblUnitSeries"
    Name ="Basic Product Construction"
End
Begin OutputColumns
    Alias ="Oracle"
    Expression ="[Basic Product Construction].[New Oracle Part #]"
    Expression ="tblUnitSeries.[Unit ID]"
    Expression ="tblUnitSeries.[Unit Type]"
    Expression ="[Basic Product Construction].[Fiber Oracle item]"
End
Begin Joins
    LeftTable ="qryPremiseItemDetails"
    RightTable ="tblUnitSeries"
    Expression ="qryPremiseItemDetails.UnitSeries = tblUnitSeries.[Unit Color Series]"
    Flag =1
    LeftTable ="tblUnitSeries"
    RightTable ="Basic Product Construction"
    Expression ="tblUnitSeries.OracleItem = [Basic Product Construction].[New Oracle Part #]"
    Flag =1
End
Begin OrderBy
    Expression ="tblUnitSeries.[Unit ID]"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbBinary "GUID" = Begin
    0xaff69b8016d0bb4fb35eeba51094d107
End
Begin
    Begin
        dbText "Name" ="Oracle"
        dbInteger "ColumnWidth" ="2025"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4d50c8cddf5ad347b108c871e5660ca2
        End
    End
    Begin
        dbText "Name" ="tblUnitSeries.[Unit ID]"
        dbInteger "ColumnWidth" ="2220"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblUnitSeries.[Unit Type]"
        dbInteger "ColumnWidth" ="3375"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Fiber Oracle item]"
        dbInteger "ColumnWidth" ="3270"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1109
    Bottom =669
    Left =-1
    Top =-1
    Right =1077
    Bottom =388
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =315
        Bottom =371
        Top =0
        Name ="qryPremiseItemDetails"
        Name =""
    End
    Begin
        Left =456
        Top =27
        Right =698
        Bottom =394
        Top =0
        Name ="tblUnitSeries"
        Name =""
    End
    Begin
        Left =1040
        Top =34
        Right =1283
        Bottom =379
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
End

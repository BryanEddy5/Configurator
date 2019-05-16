Operation =2
Name ="tblBOMSubunitComponents"
Option =0
Where ="((([Basic Product Construction].[Fiber Oracle item])=[Forms]![frmItemInterface]!"
    "[frmPremiseItemDetails].[Form]![FiberType1] Or ([Basic Product Construction].[Fi"
    "ber Oracle item])=[Forms]![frmItemInterface]![frmPremiseItemDetails].[Form]![Fib"
    "erType2] Or ([Basic Product Construction].[Fiber Oracle item])=[Forms]![frmItemI"
    "nterface]![frmPremiseItemDetails].[Form]![FiberType3] Or ([Basic Product Constru"
    "ction].[Fiber Oracle item])=\"\" Or ([Basic Product Construction].[Fiber Oracle "
    "item]) Is Null))"
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
    Expression ="tblUnitSeries.Description"
    Expression ="tblUnitSeries.[TB Color Series]"
    Expression ="tblUnitSeries.UnitNumber"
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
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x8b2a6e92c9ffa14eb8f8548bb3addeb0
End
Begin
    Begin
        dbText "Name" ="Oracle"
        dbInteger "ColumnWidth" ="2025"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x41b784e32f1a064ca0f5482bc1ab464c
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
    Begin
        dbText "Name" ="tblUnitSeries.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblUnitSeries.[TB Color Series]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblUnitSeries.UnitNumber"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =43
    Top =91
    Right =1335
    Bottom =823
    Left =-1
    Top =-1
    Right =1260
    Bottom =413
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =145
        Top =17
        Right =387
        Bottom =385
        Top =0
        Name ="qryPremiseItemDetails"
        Name =""
    End
    Begin
        Left =604
        Top =66
        Right =906
        Bottom =448
        Top =0
        Name ="tblUnitSeries"
        Name =""
    End
    Begin
        Left =1010
        Top =51
        Right =1438
        Bottom =407
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
End

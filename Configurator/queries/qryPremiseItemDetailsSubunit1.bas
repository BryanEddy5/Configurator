Operation =1
Option =0
Where ="((([Basic Product Construction].[Fiber Oracle item])=Forms!frmItemInterface!frmP"
    "remiseItemDetails.Form!FiberType1 Or ([Basic Product Construction].[Fiber Oracle"
    " item])=Forms!frmItemInterface!frmPremiseItemDetails.Form!FiberType2 Or ([Basic "
    "Product Construction].[Fiber Oracle item])=Forms!frmItemInterface!frmPremiseItem"
    "Details.Form!FiberType3 Or ([Basic Product Construction].[Fiber Oracle item])=\""
    "\" Or ([Basic Product Construction].[Fiber Oracle item]) Is Null) And ((tblUnitS"
    "eries.[Unit Color Series])=Forms!frmItemInterface!frmPremiseItemDetails.Form!Uni"
    "tSeries) And (([Product Desrcriptions].Product)=Forms!frmItemInterface!frmPremis"
    "eItemDetails.Form!Base) And (([Basic Product Construction].Active)<>0))"
Begin InputTables
    Name ="tblUnitSeries"
    Name ="Basic Product Construction"
    Name ="Product Desrcriptions"
End
Begin OutputColumns
    Alias ="Oracle"
    Expression ="[Basic Product Construction].[New Oracle Part #]"
    Expression ="tblUnitSeries.[Unit ID]"
    Expression ="tblUnitSeries.[Unit Type]"
    Expression ="[Basic Product Construction].[Fiber Oracle item]"
    Expression ="tblUnitSeries.[Unit Color Series]"
    Expression ="[Product Desrcriptions].Product"
End
Begin Joins
    LeftTable ="tblUnitSeries"
    RightTable ="Basic Product Construction"
    Expression ="tblUnitSeries.[Unit Type] = [Basic Product Construction].[Item No]"
    Flag =1
    LeftTable ="Basic Product Construction"
    RightTable ="Product Desrcriptions"
    Expression ="[Basic Product Construction].Base = [Product Desrcriptions].[Level 1 Product]"
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
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbBinary "GUID" = Begin
    0x21628ad83093834d893a84f7c9a5a626
End
Begin
    Begin
        dbText "Name" ="tblUnitSeries.[Unit Color Series]"
        dbInteger "ColumnWidth" ="2445"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblUnitSeries.[Unit Type]"
        dbInteger "ColumnWidth" ="2895"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Oracle"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6020936bf2ff654089ed6d81a7ca74bd
        End
    End
End
Begin
    State =0
    Left =43
    Top =91
    Right =1152
    Bottom =720
    Left =-1
    Top =-1
    Right =1077
    Bottom =502
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="tblUnitSeries"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =156
        Top =0
        Name ="Product Desrcriptions"
        Name =""
    End
End

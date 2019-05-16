Operation =1
Option =2
Where ="(((tblUnitSeries.[Unit Type]) Not Like \"*@*\") AND (([Basic Product Constructio"
    "n].Active)<>0))"
Begin InputTables
    Name ="qryPremiseItemCopied"
    Name ="tblUnitSeries"
    Name ="Basic Product Construction"
End
Begin OutputColumns
    Expression ="tblUnitSeries.OracleItem"
    Expression ="tblUnitSeries.[Unit Type]"
End
Begin Joins
    LeftTable ="qryPremiseItemCopied"
    RightTable ="tblUnitSeries"
    Expression ="qryPremiseItemCopied.UnitSeries = tblUnitSeries.[Unit Color Series]"
    Flag =1
    LeftTable ="tblUnitSeries"
    RightTable ="Basic Product Construction"
    Expression ="tblUnitSeries.OracleItem = [Basic Product Construction].[New Oracle Part #]"
    Flag =1
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
    0x29b8096cbc8cc047a63532f0fdef1500
End
Begin
    Begin
        dbText "Name" ="tblUnitSeries.[Unit Type]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblUnitSeries.OracleItem"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OracleItem"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc90767cbca19d14694812e6abced20cb
        End
    End
    Begin
        dbText "Name" ="Unit Type"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa6e3093edec6384fa7de6b9e2d4888fc
        End
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
    Bottom =491
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =21
        Right =348
        Bottom =327
        Top =0
        Name ="qryPremiseItemCopied"
        Name =""
    End
    Begin
        Left =387
        Top =36
        Right =611
        Bottom =313
        Top =0
        Name ="tblUnitSeries"
        Name =""
    End
    Begin
        Left =667
        Top =36
        Right =923
        Bottom =349
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
End

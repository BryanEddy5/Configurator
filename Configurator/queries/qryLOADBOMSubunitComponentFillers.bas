Operation =3
Name ="tblBOMSubunitComponents"
Option =0
Where ="(((tblUnitSeries.[Unit Type]) Like \"*@*\") AND (([Basic Product Construction].A"
    "ctive)<>0))"
Begin InputTables
    Name ="qryPremiseItemDetails"
    Name ="tblUnitSeries"
    Name ="Basic Product Construction"
    Name ="Product Desrcriptions"
    Name ="tblCableConstructionReferences"
End
Begin OutputColumns
    Alias ="Oracle"
    Name ="Oracle"
    Expression ="[Basic Product Construction].[New Oracle Part #]"
    Name ="Unit ID"
    Expression ="tblUnitSeries.[Unit ID]"
    Name ="Unit Type"
    Expression ="tblUnitSeries.[Unit Type]"
    Name ="Fiber Oracle item"
    Expression ="[Basic Product Construction].[Fiber Oracle item]"
End
Begin Joins
    LeftTable ="qryPremiseItemDetails"
    RightTable ="tblUnitSeries"
    Expression ="qryPremiseItemDetails.UnitSeries = tblUnitSeries.[Unit Color Series]"
    Flag =1
    LeftTable ="tblUnitSeries"
    RightTable ="Basic Product Construction"
    Expression ="tblUnitSeries.[Unit Type] = [Basic Product Construction].[Item No]"
    Flag =1
    LeftTable ="Basic Product Construction"
    RightTable ="Product Desrcriptions"
    Expression ="[Basic Product Construction].Base = [Product Desrcriptions].[Level 1 Product]"
    Flag =1
    LeftTable ="qryPremiseItemDetails"
    RightTable ="tblCableConstructionReferences"
    Expression ="qryPremiseItemDetails.Base = tblCableConstructionReferences.Base"
    Flag =1
    LeftTable ="tblCableConstructionReferences"
    RightTable ="Product Desrcriptions"
    Expression ="tblCableConstructionReferences.SetupID = [Product Desrcriptions].ProductID"
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
    0xf95b70015614cd45bddfde61f8d975e9
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
End
Begin
    State =0
    Left =0
    Top =40
    Right =1292
    Bottom =772
    Left =-1
    Top =-1
    Right =1260
    Bottom =212
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =47
        Top =11
        Right =262
        Bottom =249
        Top =0
        Name ="qryPremiseItemDetails"
        Name =""
    End
    Begin
        Left =783
        Top =26
        Right =1090
        Bottom =204
        Top =0
        Name ="tblUnitSeries"
        Name =""
    End
    Begin
        Left =1281
        Top =35
        Right =1453
        Bottom =183
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
    Begin
        Left =511
        Top =14
        Right =692
        Bottom =207
        Top =0
        Name ="Product Desrcriptions"
        Name =""
    End
    Begin
        Left =294
        Top =24
        Right =448
        Bottom =234
        Top =0
        Name ="tblCableConstructionReferences"
        Name =""
    End
End

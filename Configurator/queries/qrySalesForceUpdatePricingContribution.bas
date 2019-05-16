Operation =4
Option =0
Begin InputTables
    Name ="dbo_tblSalesForce_Pricing"
    Name ="tblSalesForce_Catalog_Pricing"
End
Begin OutputColumns
    Name ="dbo_tblSalesForce_Pricing.Contribution"
    Expression ="[tblSalesForce_Catalog_Pricing].[Contribution]"
End
Begin Joins
    LeftTable ="dbo_tblSalesForce_Pricing"
    RightTable ="tblSalesForce_Catalog_Pricing"
    Expression ="dbo_tblSalesForce_Pricing.Product_Category = tblSalesForce_Catalog_Pricing.Produ"
        "ct_Category"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x08dad2ad7b44a847a3c260379404dfb9
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
Begin
    Begin
        dbText "Name" ="tblCableConstructions.StandardOperation"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.JacketMaterial"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.NominalOD"
        dbInteger "ColumnWidth" ="1665"
        dbInteger "ColumnOrder" ="6"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.Base"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.AramidType1"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.AramidEnds1"
        dbInteger "ColumnOrder" ="8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.DesignCode"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.FibersPerBundle"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_tblSalesForce_Pricing.Product_Category"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblSalesForce_Catalog_Pricing.Product_Category"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_tblSalesForce_Pricing.Contribution"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =176
    Top =223
    Right =1285
    Bottom =852
    Left =-1
    Top =-1
    Right =1077
    Bottom =179
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =235
        Top =0
        Name ="dbo_tblSalesForce_Pricing"
        Name =""
    End
    Begin
        Left =332
        Top =10
        Right =1059
        Bottom =180
        Top =0
        Name ="tblSalesForce_Catalog_Pricing"
        Name =""
    End
End

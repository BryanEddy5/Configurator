Operation =4
Option =0
Begin InputTables
    Name ="tblDesignArmor"
    Name ="tblSalesForce_Catalog_Pricing"
End
Begin OutputColumns
    Name ="tblDesignArmor.ArmorCostAdder"
    Expression ="[tblSalesForce_Catalog_Pricing].[ArmorCostAdder]"
End
Begin Joins
    LeftTable ="tblDesignArmor"
    RightTable ="tblSalesForce_Catalog_Pricing"
    Expression ="tblDesignArmor.Armor = tblSalesForce_Catalog_Pricing.Armor"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x08dad2ad7b44a847a3c260379404dfb9
End
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
        dbText "Name" ="tblCableConstructionReferences.DesignCode"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_tblSalesForce_Pricing.Contribution"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.Base"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.FibersPerBundle"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.AramidType1"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_tblSalesForce_Pricing.Product_Category"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.AramidEnds1"
        dbInteger "ColumnOrder" ="8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblSalesForce_Catalog_Pricing.Product_Category"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblDesignArmor.Armor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblDesignArmor.ArmorCostAdder"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =180
    Top =187
    Right =1289
    Bottom =816
    Left =-1
    Top =-1
    Right =1077
    Bottom =162
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =211
        Top =13
        Right =542
        Bottom =222
        Top =0
        Name ="tblDesignArmor"
        Name =""
    End
    Begin
        Left =643
        Top =25
        Right =1023
        Bottom =206
        Top =0
        Name ="tblSalesForce_Catalog_Pricing"
        Name =""
    End
End

Operation =4
Option =0
Begin InputTables
    Name ="tblSalesForce_Catalog_Pricing"
    Name ="tblSalesForce_Pricing_Multiplier"
End
Begin OutputColumns
    Name ="tblSalesForce_Pricing_Multiplier.Multiplier"
    Expression ="[tblSalesForce_Catalog_Pricing].[Multiplier]"
End
Begin Joins
    LeftTable ="tblSalesForce_Catalog_Pricing"
    RightTable ="tblSalesForce_Pricing_Multiplier"
    Expression ="tblSalesForce_Catalog_Pricing.Product_Category = tblSalesForce_Pricing_Multiplie"
        "r.Product_Category"
    Flag =1
    LeftTable ="tblSalesForce_Catalog_Pricing"
    RightTable ="tblSalesForce_Pricing_Multiplier"
    Expression ="tblSalesForce_Catalog_Pricing.FiberID = tblSalesForce_Pricing_Multiplier.FiberID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x2747bacdeaab744a9932b2fe69c37e31
End
Begin
    Begin
        dbText "Name" ="tblSalesForce_Pricing_Multiplier.Multiplier"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblSalesForce_Catalog_Pricing.Multiplier"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1000"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =57
    Top =142
    Right =1166
    Bottom =771
    Left =-1
    Top =-1
    Right =1077
    Bottom =435
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =180
        Top =55
        Right =464
        Bottom =393
        Top =0
        Name ="tblSalesForce_Catalog_Pricing"
        Name =""
    End
    Begin
        Left =692
        Top =50
        Right =896
        Bottom =434
        Top =0
        Name ="tblSalesForce_Pricing_Multiplier"
        Name =""
    End
End

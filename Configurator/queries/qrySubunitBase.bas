Operation =1
Option =2
Begin InputTables
    Name ="Product Desrcriptions"
    Name ="tblCableConstructionReferences"
    Name ="tblCableConstructions"
    Name ="tblCableConstructionReferences"
    Alias ="tblCableConstructionReferences_1"
    Name ="tblCableConstructions"
    Alias ="tblCableConstructions_1"
End
Begin OutputColumns
    Alias ="SubBase"
    Expression ="[Product Desrcriptions].[Level 1 Product]"
    Expression ="tblCableConstructions_1.JacketMaterial"
    Expression ="[Product Desrcriptions].NumSubs"
    Expression ="tblCableConstructionReferences.Base"
    Expression ="[Product Desrcriptions].SubFiller"
End
Begin Joins
    LeftTable ="tblCableConstructionReferences"
    RightTable ="Product Desrcriptions"
    Expression ="tblCableConstructionReferences.ProductID = [Product Desrcriptions].ProductID"
    Flag =1
    LeftTable ="tblCableConstructions"
    RightTable ="tblCableConstructionReferences"
    Expression ="tblCableConstructions.BaseID = tblCableConstructionReferences.BaseID"
    Flag =1
    LeftTable ="tblCableConstructions_1"
    RightTable ="tblCableConstructionReferences_1"
    Expression ="tblCableConstructions_1.BaseID = tblCableConstructionReferences_1.BaseID"
    Flag =1
    LeftTable ="tblCableConstructionReferences_1"
    RightTable ="Product Desrcriptions"
    Expression ="tblCableConstructionReferences_1.Base = [Product Desrcriptions].[Level 1 Product"
        "]"
    Flag =1
    LeftTable ="Product Desrcriptions"
    RightTable ="tblCableConstructionReferences"
    Expression ="[Product Desrcriptions].ProductID = tblCableConstructionReferences.SetupID"
    Flag =1
End
Begin OrderBy
    Expression ="[Product Desrcriptions].[Level 1 Product]"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbByte "RecordsetType" ="0"
dbBinary "GUID" = Begin
    0xaad1e9de6bec2e4a83bc1ccceff75dab
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblCableConstructions_1.JacketMaterial"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2745"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="[Product Desrcriptions].NumSubs"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.Base"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SubBase"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7ba1c1a1271e48489a44d8ae848ee1f8
        End
    End
    Begin
        dbText "Name" ="[Product Desrcriptions].SubFiller"
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
    Bottom =281
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =797
        Top =9
        Right =1031
        Bottom =310
        Top =0
        Name ="Product Desrcriptions"
        Name =""
    End
    Begin
        Left =1057
        Top =2
        Right =1312
        Bottom =216
        Top =0
        Name ="tblCableConstructionReferences"
        Name =""
    End
    Begin
        Left =1330
        Top =22
        Right =1521
        Bottom =166
        Top =0
        Name ="tblCableConstructions"
        Name =""
    End
    Begin
        Left =84
        Top =15
        Right =319
        Bottom =159
        Top =0
        Name ="tblCableConstructionReferences_1"
        Name =""
    End
    Begin
        Left =391
        Top =103
        Right =621
        Bottom =252
        Top =0
        Name ="tblCableConstructions_1"
        Name =""
    End
End

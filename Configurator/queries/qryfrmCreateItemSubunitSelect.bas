Operation =1
Option =0
Where ="(((tblCableTBType.TBLetter)=Mid(Forms!frmCreateItem!ItemNo,12,1)) And ((tblCable"
    "ConstructionReferences.Base)=Forms!frmCreateItem!Base) And (([Product Desrcripti"
    "ons].SubFiller)=False))"
Begin InputTables
    Name ="tblCableTBType"
    Name ="tblCableConstructionReferences"
    Name ="tblCableTightBufferReference"
    Name ="Product Desrcriptions"
    Name ="tblCableConstructions"
    Alias ="tblCableConstructions_1"
    Name ="tblCableConstructionReferences"
    Alias ="tblCableConstructionReferences_1"
End
Begin OutputColumns
    Expression ="[Product Desrcriptions].[Level 1 Product]"
    Expression ="tblCableConstructions_1.JacketMaterial"
    Expression ="tblCableConstructionReferences_1.FibersPerBundle"
    Expression ="tblCableTightBufferReference.TBTypeID"
    Expression ="tblCableTBType.TBLetter"
End
Begin Joins
    LeftTable ="tblCableConstructions_1"
    RightTable ="tblCableConstructionReferences_1"
    Expression ="tblCableConstructions_1.BaseID = tblCableConstructionReferences_1.BaseID"
    Flag =1
    LeftTable ="Product Desrcriptions"
    RightTable ="tblCableConstructionReferences_1"
    Expression ="[Product Desrcriptions].[Level 1 Product] = tblCableConstructionReferences_1.Bas"
        "e"
    Flag =1
    LeftTable ="tblCableTightBufferReference"
    RightTable ="tblCableConstructionReferences_1"
    Expression ="tblCableTightBufferReference.TBType = tblCableConstructionReferences_1.TBType"
    Flag =1
    LeftTable ="tblCableConstructionReferences"
    RightTable ="Product Desrcriptions"
    Expression ="tblCableConstructionReferences.SetupID = [Product Desrcriptions].ProductID"
    Flag =1
    LeftTable ="tblCableTBType"
    RightTable ="tblCableTightBufferReference"
    Expression ="tblCableTBType.TBIndicatorID = tblCableTightBufferReference.TBTypeID"
    Flag =1
End
Begin OrderBy
    Expression ="[Product Desrcriptions].[Level 1 Product]"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x353adf2cfd3da242b98ece428aa91c12
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="[Product Desrcriptions].[Level 1 Product]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableTightBufferReference.TBTypeID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableTBType.TBLetter"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions_1.JacketMaterial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences_1.FibersPerBundle"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =43
    Top =91
    Right =1228
    Bottom =950
    Left =-1
    Top =-1
    Right =1153
    Bottom =243
    Left =673
    Top =0
    ColumnsShown =539
    Begin
        Left =925
        Top =45
        Right =1133
        Bottom =232
        Top =0
        Name ="tblCableTBType"
        Name =""
    End
    Begin
        Left =-598
        Top =9
        Right =-351
        Bottom =327
        Top =0
        Name ="tblCableConstructionReferences"
        Name =""
    End
    Begin
        Left =623
        Top =49
        Right =859
        Bottom =287
        Top =0
        Name ="tblCableTightBufferReference"
        Name =""
    End
    Begin
        Left =-300
        Top =21
        Right =-25
        Bottom =225
        Top =0
        Name ="Product Desrcriptions"
        Name =""
    End
    Begin
        Left =412
        Top =45
        Right =583
        Bottom =250
        Top =0
        Name ="tblCableConstructions_1"
        Name =""
    End
    Begin
        Left =36
        Top =20
        Right =352
        Bottom =300
        Top =0
        Name ="tblCableConstructionReferences_1"
        Name =""
    End
End

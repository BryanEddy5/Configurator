Operation =1
Option =0
Where ="((([Basic Product Construction].Active)<>0) And ((tblCableConstructionReferences"
    ".Base)=Forms!frmCreateItem!Base) And (([Product Desrcriptions].SubFiller)=CountS"
    "ubBases()) And (([Basic Product Construction].[Item No]) Like \"??000@*\") And ("
    "(tblUnitSeriesCopy.UnitNumber)=0))"
Begin InputTables
    Name ="tblUnitSeriesCopy"
    Name ="tblCableConstructionReferences"
    Name ="Basic Product Construction"
    Name ="Product Desrcriptions"
End
Begin OutputColumns
    Expression ="tblUnitSeriesCopy.[Unit Type]"
    Expression ="tblUnitSeriesCopy.OracleItem"
End
Begin Joins
    LeftTable ="Basic Product Construction"
    RightTable ="Product Desrcriptions"
    Expression ="[Basic Product Construction].Base = [Product Desrcriptions].[Level 1 Product]"
    Flag =1
    LeftTable ="tblCableConstructionReferences"
    RightTable ="Product Desrcriptions"
    Expression ="tblCableConstructionReferences.ProductID = [Product Desrcriptions].ProductID"
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
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbBinary "GUID" = Begin
    0x97215a6d087ffe45a6bd64859bf7ad91
End
Begin
    Begin
        dbText "Name" ="tblUnitSeriesCopy.[Unit Type]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblUnitSeriesCopy.OracleItem"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =82
    Top =156
    Right =1191
    Bottom =785
    Left =-1
    Top =-1
    Right =1077
    Bottom =339
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =81
        Top =4
        Right =283
        Bottom =308
        Top =0
        Name ="tblUnitSeriesCopy"
        Name =""
    End
    Begin
        Left =1159
        Top =14
        Right =1397
        Bottom =271
        Top =0
        Name ="tblCableConstructionReferences"
        Name =""
    End
    Begin
        Left =348
        Top =15
        Right =890
        Bottom =331
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
    Begin
        Left =922
        Top =18
        Right =1126
        Bottom =307
        Top =0
        Name ="Product Desrcriptions"
        Name =""
    End
End

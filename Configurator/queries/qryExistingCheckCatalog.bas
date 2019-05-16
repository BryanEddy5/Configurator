Operation =1
Option =0
Where ="(((tblOracleItemCatalogInterface.ItemNumber)=Forms!frmMainMenu!NewItemNumber) An"
    "d ((tblOracleItemCatalogInterface.ElementName) Not Like \"New Catalog\"))"
Begin InputTables
    Name ="tblOracleItemCatalogInterface"
End
Begin OutputColumns
    Alias ="ExistingCatalogCheck"
    Expression ="Count(tblOracleItemCatalogInterface.ElementName)"
End
Begin Parameters
    Name ="[Forms]![frmMainMenu]![NewItemNumber]"
    Flag =10
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xabdeec449dccfb408fbcff4b962ce5ef
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="ExistingCatalogCheck"
        dbInteger "ColumnWidth" ="2910"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x883d2d387450184a95416adc5f3f1e15
        End
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1560
    Bottom =854
    Left =-1
    Top =-1
    Right =1528
    Bottom =110
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =281
        Bottom =120
        Top =0
        Name ="tblOracleItemCatalogInterface"
        Name =""
    End
End

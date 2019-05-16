Operation =1
Option =2
Where ="(((qryPremiseBase_1.OracleItemNumber)=[Forms]![frmNewItemNumber]![NewItemNumber]"
    "))"
Begin InputTables
    Name ="qryPremiseBase"
    Alias ="qryPremiseBase_1"
    Name ="tblOracleItems"
End
Begin OutputColumns
    Expression ="tblOracleItems.ItemNumber"
    Expression ="tblOracleItems.Description"
    Expression ="tblOracleItems.CatalogNumber"
    Expression ="qryPremiseBase_1.OracleItemNumber"
End
Begin Joins
    LeftTable ="tblOracleItems"
    RightTable ="qryPremiseBase_1"
    Expression ="tblOracleItems.Base = qryPremiseBase_1.Base"
    Flag =1
End
Begin OrderBy
    Expression ="tblOracleItems.ItemNumber"
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
    0xc8fbf7490517f745a65f4804fbf7a0d7
End
Begin
    Begin
        dbText "Name" ="tblOracleItems.ItemNumber"
        dbInteger "ColumnWidth" ="1920"
        dbBoolean "ColumnHidden" ="0"
        dbMemo "Caption" ="Oracle Item Number"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPremiseBase_1.OracleItemNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblOracleItems.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblOracleItems.CatalogNumber"
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
    Bottom =240
    Left =1000
    Top =0
    ColumnsShown =539
    Begin
        Left =-1012
        Top =36
        Right =-757
        Bottom =263
        Top =0
        Name ="qryPremiseBase_1"
        Name =""
    End
    Begin
        Left =-1405
        Top =46
        Right =-1170
        Bottom =243
        Top =0
        Name ="tblOracleItems"
        Name =""
    End
End

Operation =1
Option =2
Where ="(((tblOracleItems.CatalogNumber) Not Like \"*box*\") And ((qryPremiseItems_2.Ora"
    "cleItemNumber)=Forms!frmNewItemNumber!NewItemNumber))"
Begin InputTables
    Name ="tblOracleItems"
    Name ="qryPremiseItems"
    Alias ="qryPremiseItems_2"
    Name ="tblCableConstructionReferences"
    Alias ="tblCableConstructionReferences_1"
    Name ="tblCableConstructionReferences"
End
Begin OutputColumns
    Expression ="tblOracleItems.ItemNumber"
    Expression ="tblOracleItems.Description"
    Expression ="tblOracleItems.CatalogNumber"
End
Begin Joins
    LeftTable ="tblCableConstructionReferences_1"
    RightTable ="tblCableConstructionReferences"
    Expression ="tblCableConstructionReferences_1.BaseID = tblCableConstructionReferences.BaseID"
    Flag =1
    LeftTable ="qryPremiseItems_2"
    RightTable ="tblCableConstructionReferences_1"
    Expression ="qryPremiseItems_2.base = tblCableConstructionReferences_1.Base"
    Flag =1
    LeftTable ="tblOracleItems"
    RightTable ="tblCableConstructionReferences"
    Expression ="tblOracleItems.Base = tblCableConstructionReferences.Base"
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
    0x485d873aaccc5841ba19696b8529cc24
End
Begin
    Begin
        dbText "Name" ="tblOracleItems.ItemNumber"
        dbMemo "Caption" ="Oracle Item Number"
        dbInteger "ColumnWidth" ="1920"
        dbBoolean "ColumnHidden" ="0"
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
    Bottom =164
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =1253
        Top =14
        Right =1558
        Bottom =181
        Top =0
        Name ="qryPremiseItems_2"
        Name =""
    End
    Begin
        Left =863
        Top =9
        Right =1150
        Bottom =215
        Top =0
        Name ="tblCableConstructionReferences_1"
        Name =""
    End
    Begin
        Left =482
        Top =16
        Right =717
        Bottom =200
        Top =0
        Name ="tblCableConstructionReferences"
        Name =""
    End
    Begin
        Left =107
        Top =17
        Right =335
        Bottom =213
        Top =0
        Name ="tblOracleItems"
        Name =""
    End
End

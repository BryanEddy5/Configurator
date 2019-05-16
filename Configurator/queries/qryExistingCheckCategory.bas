Operation =1
Option =0
Where ="((([tblOracleItemCategoryInterface].[ItemNumber])=[Forms]![frmMainMenu]![NewItem"
    "Number]))"
Begin InputTables
    Name ="tblOracleItemCategoryInterface"
End
Begin OutputColumns
    Alias ="ExistingCategoryCheck"
    Expression ="Count(tblOracleItemCategoryInterface.ItemNumber)"
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
    0xf23d5e41c3cc454f8ad166b9eee62065
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="ExistingCategoryCheck"
        dbInteger "ColumnWidth" ="2910"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x06feb604fb5f9a4b93ffa7091a4a3470
        End
    End
End
Begin
    State =0
    Left =134
    Top =182
    Right =1418
    Bottom =502
    Left =-1
    Top =-1
    Right =1252
    Bottom =110
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =366
        Top =6
        Right =617
        Bottom =120
        Top =0
        Name ="tblOracleItemCategoryInterface"
        Name =""
    End
End

Operation =1
Option =18
RowCount ="1,"
Where ="((([Basic Product Construction_1].[Fiber Oracle item])=[Forms]![frmCreateItem]!["
    "Fiber]) AND (([Basic Product Construction_1].[New Oracle Part #]) Not Like \"*te"
    "st*\") AND (([Basic Product Construction_1].Active)<>0) AND (([Product Desrcript"
    "ions].Product)=[Forms]![frmCreateItem]![Base]) AND ((Mid([Basic Product Construc"
    "tion_1]![Item No],10,1))=Mid([Forms]![frmCreateItem]![ItemNo],11,1)) AND ((Mid(["
    "Basic Product Construction_1]![Item No],11,2))=[Forms]![frmCreateItem]![TBColor]"
    " & [Forms]![frmCreateItem]![TBType]))"
Begin InputTables
    Name ="qryUnitSeriesCopy"
    Name ="Basic Product Construction"
    Name ="Basic Product Construction"
    Alias ="Basic Product Construction_1"
    Name ="Product Desrcriptions"
End
Begin OutputColumns
    Expression ="[Basic Product Construction_1].[Fiber Oracle item]"
    Expression ="[Basic Product Construction_1].OracleStatus"
    Expression ="[Basic Product Construction_1].[New Oracle Part #]"
    Expression ="[Basic Product Construction_1].Active"
    Expression ="[Basic Product Construction_1].[Item No]"
    Expression ="[Basic Product Construction_1].[Jacket Material]"
    Expression ="[Product Desrcriptions].Product"
End
Begin Joins
    LeftTable ="qryUnitSeriesCopy"
    RightTable ="Basic Product Construction"
    Expression ="qryUnitSeriesCopy.[Unit Type] = [Basic Product Construction].[Item No]"
    Flag =1
    LeftTable ="Basic Product Construction"
    RightTable ="Basic Product Construction_1"
    Expression ="[Basic Product Construction].Base = [Basic Product Construction_1].Base"
    Flag =1
    LeftTable ="Basic Product Construction_1"
    RightTable ="Basic Product Construction"
    Expression ="[Basic Product Construction_1].[TB Material] = [Basic Product Construction].[TB "
        "Material]"
    Flag =1
    LeftTable ="Product Desrcriptions"
    RightTable ="Basic Product Construction_1"
    Expression ="[Product Desrcriptions].[Level 1 Product] = [Basic Product Construction_1].Base"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FailOnError" ="0"
dbBinary "GUID" = Begin
    0xaf6e35de7122b44fa8e72dbae07f4902
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Basic Product Construction_1.New Oracle Part #"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.Item No"
        dbInteger "ColumnOrder" ="78"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.Active"
        dbInteger "ColumnWidth" ="3360"
        dbInteger "ColumnOrder" ="4"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.Fiber Oracle item"
        dbInteger "ColumnOrder" ="60"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.OracleStatus"
        dbInteger "ColumnOrder" ="73"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction_1].[Fiber Oracle item]"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction_1].[Item No]"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction_1].OracleStatus"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction_1].[New Oracle Part #]"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =36
    Top =0
    Right =1145
    Bottom =511
    Left =-1
    Top =-1
    Right =1077
    Bottom =263
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =285
        Bottom =109
        Top =0
        Name ="qryUnitSeriesCopy"
        Name =""
    End
    Begin
        Left =420
        Top =0
        Right =633
        Bottom =283
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
    Begin
        Left =770
        Top =15
        Right =1057
        Bottom =283
        Top =0
        Name ="Basic Product Construction_1"
        Name =""
    End
    Begin
        Left =1166
        Top =42
        Right =1333
        Bottom =251
        Top =0
        Name ="Product Desrcriptions"
        Name =""
    End
End

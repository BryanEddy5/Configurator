Operation =1
Option =2
Where ="(((Trim([Customer])) Like Trim([Forms]![frmCreateItem]![Customer]) & \"*\"))"
Begin InputTables
    Name ="Basic Product Construction"
End
Begin OutputColumns
    Expression ="[Basic Product Construction].[Print Line 1]"
    Expression ="[Basic Product Construction].[Print Line 2]"
    Expression ="[Basic Product Construction].[Print Line 3]"
    Expression ="[Basic Product Construction].[Print Line 4]"
    Expression ="[Basic Product Construction].[Listing Company]"
    Alias ="Expr1"
    Expression ="Trim([Customer])"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x17a8809eecf579498478dee2b7e5fb2d
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Expr1"
        dbInteger "ColumnWidth" ="3195"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x5b078398ad44d24fbac3e9ddb43e7c30
        End
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Print Line 2]"
        dbInteger "ColumnWidth" ="5865"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =96
    Top =342
    Right =1627
    Bottom =850
    Left =-1
    Top =-1
    Right =1499
    Bottom =277
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =437
        Bottom =259
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
End

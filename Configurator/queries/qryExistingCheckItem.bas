Operation =1
Option =0
Where ="((([tblOracleItemInterface].[ItemNumber])=[Forms]![frmMainMenu]![NewItemNumber])"
    ")"
Begin InputTables
    Name ="tblOracleItemInterface"
End
Begin OutputColumns
    Alias ="ExistingItemCheck"
    Expression ="Count(tblOracleItemInterface.ItemNumber)"
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
    0x1ce6f2397ad62c489bae81c749c1763a
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="ExistingItemCheck"
        dbInteger "ColumnWidth" ="2910"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x402296f2c1e08748be465dde38345ac6
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
        Left =366
        Top =6
        Right =617
        Bottom =120
        Top =0
        Name ="tblOracleItemInterface"
        Name =""
    End
End

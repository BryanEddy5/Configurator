Operation =1
Option =2
Where ="(((qryBasePrintLine1.Expr1001)=[Forms]![frmCreateItem]![Customer]))"
Begin InputTables
    Name ="qryBasePrintLine1"
End
Begin OutputColumns
    Expression ="qryBasePrintLine1.PrintLine1"
    Alias ="Count"
    Expression ="qryBasePrintLine1.PrintLine1"
    Expression ="qryBasePrintLine1.Expr1001"
End
Begin OrderBy
    Expression ="qryBasePrintLine1.PrintLine1"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xff04fbcc0312184ab2ccf24990e84ecc
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="qryBasePrintLine1.PrintLine1"
        dbInteger "ColumnWidth" ="6240"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Count"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x44fa2295c664d44d9a1b416058fe20d8
        End
        dbInteger "ColumnWidth" ="6570"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qryBasePrintLine1.Expr1001"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1001"
        dbBinary "GUID" = Begin
            0x152703e120c6f04f8c2c05afde328aaf
        End
    End
End
Begin
    State =0
    Left =49
    Top =175
    Right =1843
    Bottom =676
    Left =-1
    Top =-1
    Right =1762
    Bottom =219
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =172
        Top =6
        Right =477
        Bottom =244
        Top =0
        Name ="qryBasePrintLine1"
        Name =""
    End
End

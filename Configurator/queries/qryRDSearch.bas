Operation =1
Option =0
Begin InputTables
    Name ="qryItemNoSearch"
End
Begin OutputColumns
    Expression ="qryItemNoSearch.Oracle"
    Expression ="qryItemNoSearch.[Item No]"
    Alias ="Expr1"
    Expression ="qryItemNoSearch.Location"
    Expression ="qryItemNoSearch.Base"
End
Begin OrderBy
    Expression ="qryItemNoSearch.Oracle"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xc6497513f198044da3fbf03e878d577a
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x305aacb515a3c045ab06ff0e6edd5068
        End
    End
End
Begin
    State =0
    Left =253
    Top =208
    Right =1362
    Bottom =642
    Left =-1
    Top =-1
    Right =1077
    Bottom =203
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =387
        Bottom =156
        Top =0
        Name ="qryItemNoSearch"
        Name =""
    End
End

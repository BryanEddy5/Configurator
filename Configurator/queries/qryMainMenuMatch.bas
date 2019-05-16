Operation =1
Option =2
Where ="(((qryItemNoSearch.[Item No]) Like Left(Forms!frmMainMenu!ItemNo,8) & \"?\" & Mi"
    "d(Forms!frmMainMenu!ItemNo,10,3) & \"*\") And ((qryItemNoSearch.Active)=True))"
Begin InputTables
    Name ="qryItemNoSearch"
End
Begin OutputColumns
    Expression ="qryItemNoSearch.[Item No]"
    Expression ="qryItemNoSearch.Oracle"
    Expression ="qryItemNoSearch.[Item No]"
    Expression ="qryItemNoSearch.Reason"
    Expression ="qryItemNoSearch.Active"
End
Begin OrderBy
    Expression ="qryItemNoSearch.Oracle"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x5e3f0869b4812946818df293fea5e460
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="qryItemNoSearch.[Item No]"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryItemNoSearch.Oracle"
        dbInteger "ColumnWidth" ="2265"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1000"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xbf4c19a852b9fa489712a7acb624a075
        End
    End
    Begin
        dbText "Name" ="qryItemNoSearch.Reason"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryItemNoSearch.Active"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =112
    Top =166
    Right =1404
    Bottom =696
    Left =-1
    Top =-1
    Right =1260
    Bottom =193
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =338
        Bottom =169
        Top =0
        Name ="qryItemNoSearch"
        Name =""
    End
End

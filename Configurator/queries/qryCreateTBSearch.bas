Operation =1
Option =0
Having ="(((qryItemNoSearch.Active)=True) And ((Left(qryItemNoSearch.Oracle,7)) Like \"pt"
    "*\"))"
Begin InputTables
    Name ="qryFiberOracle"
    Name ="qryItemNoSearch"
End
Begin OutputColumns
    Alias ="TBBase"
    Expression ="Left(qryItemNoSearch.Oracle,7)"
    Expression ="qryItemNoSearch.Base"
    Expression ="qryFiberOracle.Fiber"
    Alias ="Expr1"
    Expression ="qryItemNoSearch.Location"
    Alias ="Count"
    Expression ="Count(Left(qryItemNoSearch.Oracle,7))"
End
Begin Joins
    LeftTable ="qryFiberOracle"
    RightTable ="qryItemNoSearch"
    Expression ="qryFiberOracle.Oracle = qryItemNoSearch.Oracle"
    Flag =1
End
Begin OrderBy
    Expression ="Left(qryItemNoSearch.Oracle,7)"
    Flag =0
End
Begin Groups
    Expression ="Left(qryItemNoSearch.Oracle,7)"
    GroupLevel =0
    Expression ="qryItemNoSearch.Base"
    GroupLevel =0
    Expression ="qryFiberOracle.Fiber"
    GroupLevel =0
    Expression ="qryItemNoSearch.Location"
    GroupLevel =0
    Expression ="Left(qryItemNoSearch.[Item No],5) & \"_\" & Mid(qryItemNoSearch.[Item No],7,2) &"
        " \"_\" & Mid(qryItemNoSearch.[Item No],11,10)"
    GroupLevel =0
    Expression ="qryItemNoSearch.Active"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xe5cda30561503d46b3437b904ff91874
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="TBBase"
        dbInteger "ColumnWidth" ="1575"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9420dbd214a3a0438cbf636c00c588a8
        End
    End
    Begin
        dbText "Name" ="qryItemNoSearch.Location"
        dbInteger "ColumnWidth" ="1500"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryItemNoSearch.Base"
        dbInteger "ColumnWidth" ="1560"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Count"
        dbInteger "ColumnWidth" ="1635"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xef9cbfe5bac11c49a1a25196ba18a2bc
        End
    End
    Begin
        dbText "Name" ="qryFiberOracle.Fiber"
        dbInteger "ColumnWidth" ="1485"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Base"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9773a58165b69a4b91b41bf9f599edc3
        End
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4587253a0638a249ab37523b926510f1
        End
    End
End
Begin
    State =0
    Left =107
    Top =254
    Right =1806
    Bottom =688
    Left =-1
    Top =-1
    Right =1667
    Bottom =203
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =474
        Top =6
        Right =685
        Bottom =154
        Top =0
        Name ="qryFiberOracle"
        Name =""
    End
    Begin
        Left =38
        Top =6
        Right =436
        Bottom =169
        Top =0
        Name ="qryItemNoSearch"
        Name =""
    End
End

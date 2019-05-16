Operation =1
Option =0
Where ="(((qryFiberOracle.Customer) Like \"belden*\") AND ((qryFiberOracle.[Item No]) Li"
    "ke \"kq*\" Or (qryFiberOracle.[Item No]) Like \"kp*\") AND ((qryFiberOracle.Base"
    ") Like \"k*\" Or (qryFiberOracle.Base) Like \"ia*\"))"
Begin InputTables
    Name ="qryFiberOracle"
    Name ="qryItemNoJoin"
End
Begin OutputColumns
    Expression ="qryFiberOracle.Oracle"
    Expression ="qryFiberOracle.Customer"
    Expression ="qryFiberOracle.[Item No]"
    Expression ="qryFiberOracle.Base"
    Expression ="qryFiberOracle.PID"
    Alias ="Active1"
    Expression ="IIf([qryFiberOracle].[Active]=-1,\"Yes\",\"No\")"
    Expression ="qryFiberOracle.Reason"
End
Begin Joins
    LeftTable ="qryItemNoJoin"
    RightTable ="qryFiberOracle"
    Expression ="qryItemNoJoin.Oracle = qryFiberOracle.Oracle"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbMemo "OrderBy" ="Query1.Base"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xe22ed7beb20da346a0fad243a5304b39
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="qryFiberOracle.Customer"
        dbInteger "ColumnWidth" ="3750"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.[Item No]"
        dbInteger "ColumnWidth" ="2610"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.Base"
        dbInteger "ColumnWidth" ="2580"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.PID"
        dbInteger "ColumnWidth" ="2895"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.Reason"
        dbInteger "ColumnWidth" ="9330"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Active1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3be1f889474e1c46a80ca214d1ef9659
        End
    End
End
Begin
    State =0
    Left =157
    Top =128
    Right =1779
    Bottom =829
    Left =-1
    Top =-1
    Right =1590
    Bottom =199
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =275
        Bottom =214
        Top =0
        Name ="qryFiberOracle"
        Name =""
    End
    Begin
        Left =382
        Top =9
        Right =584
        Bottom =200
        Top =0
        Name ="qryItemNoJoin"
        Name =""
    End
End

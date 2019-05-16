Operation =1
Option =16
RowCount ="1"
Where ="(((qryItemNoSearch.Oracle) Like Left([Forms]![frmFiberMatchExisting]![Oracle],7)"
    " & \"*\" And (qryItemNoSearch.Oracle) Not Like \"*-99*\"))"
Begin InputTables
    Name ="qryItemNoSearch"
End
Begin OutputColumns
    Expression ="qryItemNoSearch.Oracle"
    Alias ="Expr1"
    Expression ="Left([Forms]![frmFiberMatchExisting]![Oracle],7)"
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
    0x9adacd9160acd44390f91af7d5262a78
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="qryItemNoSearch.Oracle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe8641a47b8bd2542bbe71e1294b92d4a
        End
    End
End
Begin
    State =0
    Left =17
    Top =369
    Right =1811
    Bottom =781
    Left =-1
    Top =-1
    Right =1762
    Bottom =147
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =392
        Bottom =169
        Top =0
        Name ="qryItemNoSearch"
        Name =""
    End
End

Operation =1
Option =16
RowCount ="1,"
Having ="(((qryItemNoSearch.[Item No]) Not Like \"*test\") AND ((qryItemNoSearch.Base)=[F"
    "orms]![frmCreateItem]![Base]) AND (([qryItemNoSearch].[Location])=[Forms]![frmCr"
    "eateItem]![Location]))"
Begin InputTables
    Name ="qryItemNoSearch"
End
Begin OutputColumns
    Expression ="qryItemNoSearch.Base"
    Expression ="qryItemNoSearch.Label"
    Expression ="qryItemNoSearch.Oracle"
    Expression ="qryItemNoSearch.[Item No]"
    Alias ="Expr1"
    Expression ="qryItemNoSearch.Location"
    Expression ="qryItemNoSearch.Active"
End
Begin OrderBy
    Expression ="qryItemNoSearch.Oracle"
    Flag =1
    Expression ="Count(qryItemNoSearch.[Item No])"
    Flag =0
End
Begin Groups
    Expression ="qryItemNoSearch.Base"
    GroupLevel =0
    Expression ="qryItemNoSearch.Label"
    GroupLevel =0
    Expression ="qryItemNoSearch.Oracle"
    GroupLevel =0
    Expression ="qryItemNoSearch.[Item No]"
    GroupLevel =0
    Expression ="qryItemNoSearch.Location"
    GroupLevel =0
    Expression ="qryItemNoSearch.Active"
    GroupLevel =0
    Expression ="qryItemNoSearch.Base"
    GroupLevel =0
    Expression ="Mid([Item No],9,1)"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbBinary "GUID" = Begin
    0x1cfe5089edc60d428c4afa4802e26709
End
Begin
    Begin
        dbText "Name" ="qryItemNoSearch.Base"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryItemNoSearch.Label"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryItemNoSearch.Oracle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryItemNoSearch.[Item No]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryItemNoSearch.Location"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryItemNoSearch.Active"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x32ab49aa6d9dca4f97b5a0642df095b5
        End
    End
End
Begin
    State =0
    Left =22
    Top =138
    Right =1730
    Bottom =757
    Left =-1
    Top =-1
    Right =1676
    Bottom =177
    Left =18
    Top =0
    ColumnsShown =543
    Begin
        Left =20
        Top =6
        Right =335
        Bottom =154
        Top =0
        Name ="qryItemNoSearch"
        Name =""
    End
End

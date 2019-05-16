Operation =1
Option =0
Where ="(((qryItemNoJoin.Active)=True))"
Begin InputTables
    Name ="qryItemNoJoin"
End
Begin OutputColumns
    Alias ="ItemNo"
    Expression ="Left([Item No],5) & Mid([Item No],7,2) & Mid(IIf([Item No] Like \"*-sh*\" Or [It"
        "em No] Like \"*-rh*\" Or [Item No] Like \"*[1-9]\",Left([Item No],15),Left([Item"
        " No],12)),13,30)"
    Expression ="qryItemNoJoin.Base"
    Expression ="qryItemNoJoin.Active"
End
Begin OrderBy
    Expression ="qryItemNoJoin.Base"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xaf6eff65abbae143a515854b3a0a2e07
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="ItemNo "
        dbInteger "ColumnWidth" ="2070"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemNo"
        dbInteger "ColumnWidth" ="3570"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd1abe0c9357eb3439e193768f5e7a28b
        End
    End
    Begin
        dbText "Name" ="qryItemNoJoin.Base"
        dbInteger "ColumnWidth" ="3210"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =102
    Top =201
    Right =1394
    Bottom =525
    Left =-1
    Top =-1
    Right =1260
    Bottom =93
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =124
        Top =0
        Name ="qryItemNoJoin"
        Name =""
    End
End

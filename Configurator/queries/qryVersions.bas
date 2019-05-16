Operation =1
Option =16
RowCount ="1"
Begin InputTables
    Name ="tblVersion"
    Name ="qryApplicationSystem"
End
Begin OutputColumns
    Expression ="tblVersion.*"
    Alias ="Expr1"
    Expression ="qryApplicationSystem.Hyperlink"
End
Begin OrderBy
    Expression ="tblVersion.RevDate"
    Flag =1
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
    0x833f11a07cd5bf40b1e62dc3a6e8bcec
End
Begin
    Begin
        dbText "Name" ="tblVersion.AppVersionNo"
        dbInteger "ColumnWidth" ="1500"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblVersion.AppVersionSub"
        dbInteger "ColumnWidth" ="1620"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblVersion.RevDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryApplicationSystem.Hyperlink"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x53a2a7eb1a9a6446b4a54aa516cc585e
        End
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1109
    Bottom =669
    Left =-1
    Top =-1
    Right =1077
    Bottom =160
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =261
        Bottom =105
        Top =0
        Name ="tblVersion"
        Name =""
    End
    Begin
        Left =456
        Top =43
        Right =703
        Bottom =127
        Top =0
        Name ="qryApplicationSystem"
        Name =""
    End
End

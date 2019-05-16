Operation =1
Option =0
Begin InputTables
    Name ="qrySalesForce_DaysToCloseItemSetup"
End
Begin OutputColumns
    Alias ="Month"
    Expression ="Format([TaskOpened],\"yyyy/mm\")"
    Alias ="Average Days Open"
    Expression ="Avg(qrySalesForce_DaysToCloseItemSetup.DaysToCloseTask)"
    Expression ="qrySalesForce_DaysToCloseItemSetup.Assigned"
End
Begin OrderBy
    Expression ="Format([TaskOpened],\"yyyy/mm\")"
    Flag =0
End
Begin Groups
    Expression ="Format([TaskOpened],\"yyyy/mm\")"
    GroupLevel =0
    Expression ="qrySalesForce_DaysToCloseItemSetup.Assigned"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xe10de4306094bb47878cfec04b9d8cf1
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Month"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xbf2cdeefa215b44ea30b6b1eb71063b1
        End
    End
    Begin
        dbText "Name" ="Average Days Open"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2220"
        dbBoolean "ColumnHidden" ="0"
        dbBinary "GUID" = Begin
            0x9072cba9e81ad04ca88d793657d34118
        End
    End
    Begin
        dbText "Name" ="qrySalesForce_DaysToCloseItemSetup.Assigned"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =52
    Top =200
    Right =1161
    Bottom =829
    Left =-1
    Top =-1
    Right =1077
    Bottom =290
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =328
        Bottom =331
        Top =0
        Name ="qrySalesForce_DaysToCloseItemSetup"
        Name =""
    End
End

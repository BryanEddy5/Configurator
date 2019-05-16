dbMemo "SQL" ="SELECT tblTestItemInterface.ItemNumber, tblTestItemInterface.Description, tblTes"
    "tItemInterface.PrimaryUOMCode, tblTestItemInterface.TemplateName, tblTestItemInt"
    "erface.StdLotSize, tblNetworkItemInterface.ItemNumber, tblNetworkItemInterface.D"
    "escription, tblNetworkItemInterface.PrimaryUOMCode, tblNetworkItemInterface.Temp"
    "lateName, tblNetworkItemInterface.StdLotSize\015\012FROM tblTestItemInterface LE"
    "FT JOIN tblNetworkItemInterface ON (tblTestItemInterface.ItemNumber = tblNetwork"
    "ItemInterface.ItemNumber) AND (tblTestItemInterface.OrganizationCode = tblNetwor"
    "kItemInterface.OrganizationCode);\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xb9831fc4391dca42b02962f2b69e94b4
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="tblTestItemInterface.Description"
        dbInteger "ColumnWidth" ="5070"
        dbInteger "ColumnOrder" ="3"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblTestItemInterface.ItemNumber"
        dbInteger "ColumnWidth" ="2985"
        dbInteger "ColumnOrder" ="1"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblTestItemInterface.PrimaryUOMCode"
        dbInteger "ColumnWidth" ="570"
        dbInteger "ColumnOrder" ="6"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblTestItemInterface.TemplateName"
        dbInteger "ColumnWidth" ="1995"
        dbInteger "ColumnOrder" ="7"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblTestItemInterface.StdLotSize"
        dbInteger "ColumnWidth" ="960"
        dbInteger "ColumnOrder" ="9"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNetworkItemInterface.ItemNumber"
        dbInteger "ColumnWidth" ="3315"
        dbInteger "ColumnOrder" ="2"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNetworkItemInterface.Description"
        dbInteger "ColumnWidth" ="5070"
        dbInteger "ColumnOrder" ="4"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNetworkItemInterface.PrimaryUOMCode"
        dbInteger "ColumnWidth" ="645"
        dbInteger "ColumnOrder" ="5"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNetworkItemInterface.TemplateName"
        dbInteger "ColumnWidth" ="1965"
        dbInteger "ColumnOrder" ="8"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNetworkItemInterface.StdLotSize"
        dbInteger "ColumnWidth" ="735"
        dbInteger "ColumnOrder" ="10"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End

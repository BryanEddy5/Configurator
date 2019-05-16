dbMemo "SQL" ="SELECT tblTestBOMBillOfMtlsInterface.ItemNumber, tblTestBOMBillOfMtlsInterface.A"
    "ssemblyType, tblTestBOMBillOfMtlsInterface.CommonOrgCode, tblTestBOMBillOfMtlsIn"
    "terface.CommonItemNumber, tblNetworkBOMBillOfMtlsInterface.ItemNumber, tblNetwor"
    "kBOMBillOfMtlsInterface.AssemblyType, tblNetworkBOMBillOfMtlsInterface.CommonOrg"
    "Code, tblNetworkBOMBillOfMtlsInterface.CommonItemNumber\015\012FROM tblTestBOMBi"
    "llOfMtlsInterface LEFT JOIN tblNetworkBOMBillOfMtlsInterface ON (tblTestBOMBillO"
    "fMtlsInterface.OrganizationCode=tblNetworkBOMBillOfMtlsInterface.OrganizationCod"
    "e) AND (tblTestBOMBillOfMtlsInterface.ItemNumber=tblNetworkBOMBillOfMtlsInterfac"
    "e.ItemNumber);\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x9ea1232028a86046bf2c933986630f04
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="tblTestBOMBillOfMtlsInterface.ItemNumber"
        dbInteger "ColumnWidth" ="3900"
        dbInteger "ColumnOrder" ="1"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblTestBOMBillOfMtlsInterface.AssemblyType"
        dbInteger "ColumnWidth" ="4140"
        dbInteger "ColumnOrder" ="3"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblTestBOMBillOfMtlsInterface.CommonOrgCode"
        dbInteger "ColumnWidth" ="4395"
        dbInteger "ColumnOrder" ="5"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblTestBOMBillOfMtlsInterface.CommonItemNumber"
        dbInteger "ColumnWidth" ="4680"
        dbInteger "ColumnOrder" ="7"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNetworkBOMBillOfMtlsInterface.ItemNumber"
        dbInteger "ColumnWidth" ="4230"
        dbInteger "ColumnOrder" ="2"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNetworkBOMBillOfMtlsInterface.AssemblyType"
        dbInteger "ColumnWidth" ="4470"
        dbInteger "ColumnOrder" ="4"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNetworkBOMBillOfMtlsInterface.CommonOrgCode"
        dbInteger "ColumnWidth" ="4725"
        dbInteger "ColumnOrder" ="6"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNetworkBOMBillOfMtlsInterface.CommonItemNumber"
        dbInteger "ColumnWidth" ="5010"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End

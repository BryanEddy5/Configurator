dbMemo "SQL" ="SELECT DISTINCT tblOracleItemInterface.ItemNumber, tblOracleItemInterface.Descri"
    "ption, IIf([AssemblyItemNumber] Is Null,\"No\",\"Yes\") AS Routing, IIf(tblOracl"
    "eBOMBillOfMtlsInterface!ItemNumber Is Null,\"No\",\"Yes\") AS BOM\015\012FROM (("
    "tblOracleItemInterface LEFT JOIN tblOracleBOMBillOfMtlsInterface ON tblOracleIte"
    "mInterface.ItemNumber=tblOracleBOMBillOfMtlsInterface.ItemNumber) LEFT JOIN tblO"
    "racleBOMRoutingsAttachments ON tblOracleItemInterface.ItemNumber=tblOracleBOMRou"
    "tingsAttachments.AssemblyItemNumber) INNER JOIN qryDossertRoutings ON tblOracleI"
    "temInterface.ItemNumber=qryDossertRoutings.DossertItem\015\012WHERE (((IIf([Asse"
    "mblyItemNumber] Is Null,\"No\",\"Yes\"))=\"Yes\") AND ((IIf(tblOracleBOMBillOfMt"
    "lsInterface!ItemNumber Is Null,\"No\",\"Yes\"))=\"Yes\"));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xc2afb6b7f219ec48b09e837a41ab09b8
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="tblOracleItemInterface.ItemNumber"
        dbInteger "ColumnWidth" ="1620"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblOracleItemInterface.Description"
        dbInteger "ColumnWidth" ="4335"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Routing"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3b5cd70004b65542ab0230f4411fdaac
        End
    End
    Begin
        dbText "Name" ="BOM"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xbe6d88113dc0a34d83e4648369d782a7
        End
    End
End

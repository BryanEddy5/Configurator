dbMemo "SQL" ="SELECT DISTINCT tblOCAObsoleteSaddie.Oracle, tblOCAObsoleteSaddie.Description, t"
    "blShipmentsCable.bill_to_name\015\012FROM tblOCAObsoleteSaddie INNER JOIN tblShi"
    "pmentsCable ON tblOCAObsoleteSaddie.Oracle=tblShipmentsCable.item_number\015\012"
    "WHERE (((tblShipmentsCable.bill_to_name) Not Like \"*belden*\"));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xa03459235478db47b92bcad95d545a36
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="tblOCAObsoleteSaddie.Description"
        dbInteger "ColumnWidth" ="9780"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblShipmentsCable.bill_to_name"
        dbInteger "ColumnWidth" ="4290"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End

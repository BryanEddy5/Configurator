dbMemo "SQL" ="SELECT DISTINCT qryBurnListingDoublePass.*, [Basic Product Construction].Base, ["
    "Basic Product Construction].[Jacket Material], [Basic Product Construction].[TB "
    "Material], [Basic Product Construction].Active\015\012FROM (([Basic (Cable/Sheat"
    "h)] INNER JOIN (tblTake2PassBurn INNER JOIN qryBurnListingDoublePass ON tblTake2"
    "PassBurn.Base = qryBurnListingDoublePass.Base) ON [Basic (Cable/Sheath)].Base = "
    "qryBurnListingDoublePass.Base) INNER JOIN [Unit Series] ON [Basic (Cable/Sheath)"
    "].[Unit Series] = [Unit Series].[Unit Color Series]) INNER JOIN [Basic Product C"
    "onstruction] ON [Unit Series].[Unit Type] = [Basic Product Construction].[Item N"
    "o]\015\012WHERE ((([Basic Product Construction].Active)=-1));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xd31d86ae47cf41489108a611677decaa
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="qryBurnListingDoublePass.[Basic (Cable/Sheath)].Active"
        dbInteger "ColumnWidth" ="945"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryBurnListingDoublePass.tblBurnListing.ListingConstruction"
        dbInteger "ColumnWidth" ="810"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryBurnListingDoublePass.OD"
        dbInteger "ColumnWidth" ="1005"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryBurnListingDoublePass.NomWall"
        dbInteger "ColumnWidth" ="780"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End

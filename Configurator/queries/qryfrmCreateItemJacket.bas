dbMemo "SQL" ="SELECT DISTINCT qryItemNoSearchMatch.Base, qry_Jackets_Subs.[Jacket Material]\015"
    "\012FROM (qry_Match_Jacket_Base INNER JOIN qry_Jackets_Subs ON qry_Match_Jacket_"
    "Base.Expr1002=qry_Jackets_Subs.JacketBase) INNER JOIN qryItemNoSearchMatch ON qr"
    "y_Match_Jacket_Base.Base=qryItemNoSearchMatch.Base;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x49ee540a587088428e79b19e13d0c2c8
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="qry_Jackets_Subs.[Jacket Material]"
        dbInteger "ColumnWidth" ="7140"
        dbInteger "ColumnOrder" ="1"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End

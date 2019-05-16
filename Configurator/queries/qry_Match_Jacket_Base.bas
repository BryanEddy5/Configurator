dbMemo "SQL" ="SELECT DISTINCT [Basic Product Construction].Base, [Basic Product Construction]."
    "[Jacket Material], left([Basic Product Construction].[Jacket Material],8), [Prin"
    "t Line 3] AS PrintLine3\015\012FROM [Basic Product Construction]\015\012WHERE [B"
    "asic Product Construction].[Active] <> 0 AND  (([Basic Product Construction].[Ne"
    "w Oracle Part #])  not like \"*test*\");\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xe5531ca747062d4e9eab53aae0364779
End
dbMemo "Filter" ="((qry_Match_Jacket_Base.Base=\"BU004-2.0M\"))"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="[Basic Product Construction].[Jacket Material]"
        dbInteger "ColumnWidth" ="3405"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1002"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x63dd65c850269d40a3b7078ed9e5026c
        End
    End
    Begin
        dbText "Name" ="PrintLine3"
        dbInteger "ColumnWidth" ="3705"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x5313fad513119f4d8eaf084447731f3e
        End
    End
End

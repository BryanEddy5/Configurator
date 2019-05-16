dbMemo "SQL" ="SELECT DISTINCT qryFiberOracle.Oracle, tbl_Fibers.Fiber AS Expr1, tbl_Fibers.Mod"
    "e AS Expr2, qryFiberOracle.[Item No], tbl_Fibers.BIF AS Expr3, tbl_Fibers.[OM#] "
    "AS Expr4, tbl_Fibers.SpecialNotes AS Expr5, tblAFLPRD_INVSysItemCost_CAB.Item_St"
    "atus, tbl_Fibers.Description_Oracle AS Expr6, [Premise Sales History].Customer\015"
    "\012FROM tbl_Fibers, (qryFiberOracle INNER JOIN tblAFLPRD_INVSysItemCost_CAB ON "
    "qryFiberOracle.Oracle = tblAFLPRD_INVSysItemCost_CAB.ItemNumber) INNER JOIN [Pre"
    "mise Sales History] ON tblAFLPRD_INVSysItemCost_CAB.ItemNumber = [Premise Sales "
    "History].[Oracle Item Number]\015\012WHERE (((qryFiberOracle.Oracle) Not Like \""
    "fb*\") And ((qryFiberOracle.[Item No]) Not Like \"t*\" And (qryFiberOracle.[Item"
    " No]) Not Like \"e*\") And ((tblAFLPRD_INVSysItemCost_CAB.Item_Status) Not Like "
    "\"*obs*\" And (tblAFLPRD_INVSysItemCost_CAB.Item_Status) Not Like \"*disc*\") An"
    "d ((tbl_Fibers.Fiber) Like \"*161*\" Or (tbl_Fibers.Fiber) Like \"*119*\" Or (tb"
    "l_Fibers.Fiber) Like \"*131*\") And ((tbl_Fibers.BIF)=\"no\") And ((tbl_Fibers.["
    "OM#]) Not Like \"*om1*\") And ((tbl_Fibers.SpecialNotes) Not Like \"*500*\") And"
    " ((qryFiberOracle.Active)=-1) And (([Premise Sales History].[Shipped date])>#9/1"
    "/2013#))\015\012ORDER BY tbl_Fibers.Fiber, qryFiberOracle.[Item No];\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x6a1fb94c7a920c41b1ad1d8b6893fb02
End
dbMemo "OrderBy" ="qryMMnonBIFMike.Customer"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="qryFiberOracle.Oracle"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFiberOracle.[Item No]"
        dbInteger "ColumnWidth" ="3885"
        dbInteger "ColumnOrder" ="2"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblAFLPRD_INVSysItemCost_CAB.Item_Status"
        dbInteger "ColumnWidth" ="1980"
        dbInteger "ColumnOrder" ="8"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x577ef3af3adf2a44bfdc7fe0cf7bbfe6
        End
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9e71d5baa8457a4db737f8489370135a
        End
    End
    Begin
        dbText "Name" ="Expr3"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9f454b21ea936f47a6c7f024340c7187
        End
    End
    Begin
        dbText "Name" ="Expr4"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf2fcfd2e5bd83a45acd1af7d4bd48e22
        End
    End
    Begin
        dbText "Name" ="Expr5"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xda42c31c9c4b0d418b9d4998e8dc6aeb
        End
    End
    Begin
        dbText "Name" ="Expr6"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa26bb480f7392b40ad22116192c56046
        End
    End
End

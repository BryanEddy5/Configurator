dbMemo "SQL" ="SELECT DISTINCT qryFiberOracle.Oracle, tbl_Fibers.Fiber, tbl_Fibers.Mode, qryFib"
    "erOracle.[Item No], tbl_Fibers.BIF, tbl_Fibers.[OM#], tbl_Fibers.SpecialNotes, t"
    "blAFLPRD_INVSysItemCost_CAB.Item_Status, tbl_Fibers.Description_Oracle, tblShipm"
    "entsCable.bill_to_name, qryFiberOracle.PID\015\012FROM tblShipmentsCable INNER J"
    "OIN ((qryFiberOracle INNER JOIN tbl_Fibers ON qryFiberOracle.Fiber = tbl_Fibers."
    "Fiber) INNER JOIN tblAFLPRD_INVSysItemCost_CAB ON qryFiberOracle.Oracle = tblAFL"
    "PRD_INVSysItemCost_CAB.ItemNumber) ON tblShipmentsCable.item_number = tblAFLPRD_"
    "INVSysItemCost_CAB.ItemNumber\015\012WHERE (((qryFiberOracle.Oracle) Not Like \""
    "fb*\" And (qryFiberOracle.Oracle) Not Like \"*-01*\") AND ((tbl_Fibers.Mode)=\"m"
    "m\") AND ((qryFiberOracle.[Item No]) Not Like \"t*\" And (qryFiberOracle.[Item N"
    "o]) Not Like \"e*\") AND ((tbl_Fibers.BIF)=\"no\") AND ((tbl_Fibers.[OM#]) Not L"
    "ike \"*om1*\") AND ((tbl_Fibers.SpecialNotes) Not Like \"*500*\") AND ((tblAFLPR"
    "D_INVSysItemCost_CAB.Item_Status) Not Like \"*obs*\" And (tblAFLPRD_INVSysItemCo"
    "st_CAB.Item_Status) Not Like \"*disc*\") AND ((qryFiberOracle.PID) Is Not Null A"
    "nd (qryFiberOracle.PID)<>\"\") AND ((qryFiberOracle.Active)=-1) AND ((tblShipmen"
    "tsCable.transaction_date)>#9/1/2013#))\015\012ORDER BY tbl_Fibers.Fiber, tbl_Fib"
    "ers.Mode, qryFiberOracle.[Item No];\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x1ed8a812639a8c46a60336fed59a833c
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="qryFiberOracle.Oracle"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Fibers.BIF"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Fibers.Fiber"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Fibers.Mode"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Fibers.[OM#]"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Fibers.SpecialNotes"
        dbInteger "ColumnWidth" ="1695"
        dbInteger "ColumnOrder" ="7"
        dbBoolean "ColumnHidden" ="0"
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
        dbText "Name" ="qryFiberOracle.PID"
        dbInteger "ColumnWidth" ="1965"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Fibers.Description_Oracle"
        dbInteger "ColumnWidth" ="4530"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblShipmentsCable.bill_to_name"
        dbInteger "ColumnWidth" ="4365"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End

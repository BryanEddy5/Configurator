dbMemo "SQL" ="SELECT tblAFLPRD_INVSysItemCost_CAB.ItemNumber, tblAFLPRD_INVSysItemCost_CAB.Des"
    "cription, AFL_INV_Customer_Item_Cross_Re_091015.[Customer Name], AFL_INV_Custome"
    "r_Item_Cross_Re_091015.[Customer Num], AFL_INV_Customer_Item_Cross_Re_091015.[Cu"
    "st Item], tblAFLPRD_INVSysItemCost_CAB.Item_Cost, tblAFLPRD_INVSysItemCost_CAB.I"
    "tem_Status\015\012FROM tblAFLPRD_INVSysItemCost_CAB INNER JOIN AFL_INV_Customer_"
    "Item_Cross_Re_091015 ON tblAFLPRD_INVSysItemCost_CAB.ItemNumber = AFL_INV_Custom"
    "er_Item_Cross_Re_091015.Item\015\012WHERE (((AFL_INV_Customer_Item_Cross_Re_0910"
    "15.[Customer Name]) Like \"te con*\" Or (AFL_INV_Customer_Item_Cross_Re_091015.["
    "Customer Name]) Like \"tyco*\"));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x417668580452574ca6680686dc671e91
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblAFLPRD_INVSysItemCost_CAB.Item_Cost"
        dbInteger "ColumnWidth" ="2250"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblAFLPRD_INVSysItemCost_CAB.ItemNumber"
        dbInteger "ColumnWidth" ="3495"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AFL_INV_Customer_Item_Cross_Re_091015.[Customer Name]"
        dbInteger "ColumnWidth" ="4035"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AFL_INV_Customer_Item_Cross_Re_091015.[Cust Item]"
        dbInteger "ColumnWidth" ="2595"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblAFLPRD_INVSysItemCost_CAB.Description"
        dbInteger "ColumnWidth" ="7935"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End

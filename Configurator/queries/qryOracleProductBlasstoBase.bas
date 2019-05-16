dbMemo "SQL" ="SELECT qryFiberOracle.Base, tblAFLPRD_INVItmCatg_CAB.CategoryName, Count(tblAFLP"
    "RD_INVItmCatg_CAB.CategoryName) AS CountOfCategoryName\015\012FROM qryFiberOracl"
    "e INNER JOIN tblAFLPRD_INVItmCatg_CAB ON qryFiberOracle.Oracle=tblAFLPRD_INVItmC"
    "atg_CAB.ItemNumber\015\012GROUP BY qryFiberOracle.Base, tblAFLPRD_INVItmCatg_CAB"
    ".CategoryName, tblAFLPRD_INVItmCatg_CAB.CategorySetName\015\012HAVING (((tblAFLP"
    "RD_INVItmCatg_CAB.CategorySetName)=\"new product type\"))\015\012ORDER BY qryFib"
    "erOracle.Base, Count(tblAFLPRD_INVItmCatg_CAB.CategoryName) DESC;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xbc6af7b68f41ef49bfa8b48568ff0cfa
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="tblAFLPRD_INVItmCatg_CAB.CategoryName"
        dbInteger "ColumnWidth" ="1950"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CountOfCategoryName"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xabad843ba41dee47a5471a19050be6bb
        End
    End
End

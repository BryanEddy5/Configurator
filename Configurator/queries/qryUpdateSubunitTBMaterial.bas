dbMemo "SQL" ="SELECT qryfrmCreateItemTBType.TBType, qryfrmCreateItemTBType.Family, qryfrmCreat"
    "eItemTBType.[Jacket Material], tblNewPrductConstruction.[TB Material], tblNewPrd"
    "uctConstruction.[TB Chips Type]\015\012FROM tblNewPrductConstruction, qryfrmCrea"
    "teItemTBTypeMax INNER JOIN qryfrmCreateItemTBType ON (qryfrmCreateItemTBTypeMax."
    "[MaxOfCountOfTB Material] = qryfrmCreateItemTBType.[CountOfTB Material]) AND (qr"
    "yfrmCreateItemTBTypeMax.[Jacket Material] = qryfrmCreateItemTBType.[Jacket Mater"
    "ial]) AND (qryfrmCreateItemTBType.TBType = qryfrmCreateItemTBTypeMax.TBType) AND"
    " (qryfrmCreateItemTBTypeMax.Family = qryfrmCreateItemTBType.Family)\015\012WHERE"
    " (((qryfrmCreateItemTBType.TBType)=Mid(tblNewPrductConstruction.[Item No],12,1))"
    " And ((qryfrmCreateItemTBType.Family)=Mid(tblNewPrductConstruction.[Item No],1,2"
    ")) And ((qryfrmCreateItemTBType.[Jacket Material])=tblNewPrductConstruction.[Jac"
    "ket Material]));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x3e6db82e33c9c54297e48bc4e2b806c5
End
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
End

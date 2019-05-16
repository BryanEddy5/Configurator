dbMemo "SQL" ="SELECT DISTINCTROW [Unit Series].*\015\012FROM tblBasicCableCopy INNER JOIN [Uni"
    "t Series] ON tblBasicCableCopy.[Unit Series] = [Unit Series].[Unit Color Series]"
    "\015\012WHERE (((tblBasicCableCopy.Base)=Forms!frmCreateItem!Base) And ((tblBasi"
    "cCableCopy.Active)=-1))\015\012ORDER BY [Unit Series].[Unit Color Series], [Unit"
    " Series].[TB Color Series];\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x446514d9704fa34698ac9b3088e17827
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
End

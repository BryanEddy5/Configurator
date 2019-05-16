dbMemo "SQL" ="SELECT DISTINCT tblOracleDepartmentCodes.SetupRequired, tblOracleDepartmentCodes"
    ".DepartmentCode, tblOracleResourceCodes.ResourceCode\015\012FROM tblOracleDepart"
    "mentCodes INNER JOIN tblOracleResourceCodes ON tblOracleDepartmentCodes.Departme"
    "ntCode = tblOracleResourceCodes.DepartmentCode\015\012WHERE (((tblOracleResource"
    "Codes.ResourceCode)=\"setup\"));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x5a1bb29e268b6841a1b675413917da38
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
End

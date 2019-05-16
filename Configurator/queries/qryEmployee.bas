dbMemo "SQL" ="SELECT tblEmployees.*, [FirstName] & \" \" & [LastName] AS Name, tblEmployeeResp"
    "onsibility.ApplicationResponsibility\015\012FROM tblEmployees INNER JOIN tblEmpl"
    "oyeeResponsibility ON tblEmployees.EMPLOYEE_ID=tblEmployeeResponsibility.Employe"
    "eID\015\012WHERE (((tblEmployeeResponsibility.ApplicationName)=Forms!frmBegin!Ap"
    "plicationName));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x59ff03747f849642aae0f902e387850e
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="Name"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x229481560eaf1242aab20fe58a6409e1
        End
    End
End

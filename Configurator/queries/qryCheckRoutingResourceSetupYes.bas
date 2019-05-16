dbMemo "SQL" ="SELECT \"\" AS ComponentItemNumber, \"\" AS ItemSeqNumber, \"Routing - Missing S"
    "etup Resource for \" & [qryCheckDepartments]![DepartmentCode] AS CheckCode, qryC"
    "heckDepartments.OperationSeqNum, qryCheckRoutingResourceSetup.ResourceCode, qryC"
    "heckDepartments.SetupRequired, 0 AS StopLoad, qryCheckDepartments.DepartmentCode"
    "\015\012FROM qryCheckResourceSetup INNER JOIN (qryCheckDepartments LEFT JOIN qry"
    "CheckRoutingResourceSetup ON qryCheckDepartments.OperationSeqNum = qryCheckRouti"
    "ngResourceSetup.OperationSeqNumber) ON qryCheckResourceSetup.DepartmentCode = qr"
    "yCheckDepartments.DepartmentCode\015\012WHERE (((qryCheckDepartments.SetupRequir"
    "ed)=-1) AND ((qryCheckRoutingResourceSetup.DepartmentCode) Is Null));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x4b2d0ae57225324c913849ad59311766
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="CheckCode"
        dbInteger "ColumnWidth" ="3960"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x42453cfac6d52b46847ac47ec799221e
        End
    End
    Begin
        dbText "Name" ="qryCheckDepartments.DepartmentCode"
        dbInteger "ColumnWidth" ="1635"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ComponentItemNumber"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x5e00547b690f21499f0d02e1e1d4582c
        End
    End
    Begin
        dbText "Name" ="ItemSeqNumber"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4c301a2417c2944da99082167a41d3c4
        End
    End
    Begin
        dbText "Name" ="StopLoad"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6e3af9d2cbaf9f4e95ad78268121a799
        End
    End
End

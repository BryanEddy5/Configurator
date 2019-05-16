dbMemo "SQL" ="SELECT \"\" AS ComponentItemNumber, \"\" AS ItemSeqNumber, \"Routing - Unneccess"
    "ary Setup Resource\" AS CheckCode, qryCheckDepartments.OperationSeqNum, qryCheck"
    "RoutingResourceSetup.ResourceCode, qryCheckDepartments.SetupRequired, 0 AS StopL"
    "oad\015\012FROM qryCheckDepartments LEFT JOIN qryCheckRoutingResourceSetup ON qr"
    "yCheckDepartments.OperationSeqNum = qryCheckRoutingResourceSetup.OperationSeqNum"
    "ber\015\012WHERE (((qryCheckDepartments.SetupRequired)=0) AND ((qryCheckRoutingR"
    "esourceSetup.DepartmentCode) Is Not Null));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x38775aeeeb6e954aa3537fc66a722d32
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="CheckCode"
        dbInteger "ColumnWidth" ="3720"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x0ab440c70082f0449c376fc9bcf06567
        End
    End
    Begin
        dbText "Name" ="ComponentItemNumber"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xdd180bf40335f341940ca8bd17699e2b
        End
    End
    Begin
        dbText "Name" ="ItemSeqNumber"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x023e0ddc10d3da4cbf1c293545fa5244
        End
    End
    Begin
        dbText "Name" ="StopLoad"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa59997bfae11024d944acfeae6bf2eaa
        End
    End
End

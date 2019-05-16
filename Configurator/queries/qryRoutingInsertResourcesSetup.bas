dbMemo "SQL" ="INSERT INTO tblBOMOpResourcesInterface ( OperationSeqNumber, OrganizationCode, A"
    "ssemblyItemNumber, AlternateRoutingDesignator, ResourceSeqNumber, ResourceCode, "
    "AssignedUnits, OrgOpSeq )\015\012SELECT tblBOMOpSequencesInterfaceTemp.Operation"
    "SeqNum, tblBOMOpSequencesInterfaceTemp.OrganizationCode, tblBOMOpSequencesInterf"
    "aceTemp.AssemblyItemNumber, tblBOMOpSequencesInterfaceTemp.AlternateRoutingDesig"
    "nator, 2 AS ResourceSeqNumber, \"SETUP\" AS ResourceCode, 1 AS AssignedUnits, tb"
    "lBOMOpSequencesInterfaceTemp.OrgOpSeq\015\012FROM tblBOMOpSequencesInterfaceTemp"
    " INNER JOIN tblOracleDepartmentCodes ON tblBOMOpSequencesInterfaceTemp.Departmen"
    "tCode=tblOracleDepartmentCodes.DepartmentCode\015\012WHERE (((tblOracleDepartmen"
    "tCodes.SetupRequired)=-1));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x378d25c73f0321449b1fe0d4ad4678d5
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="tblBOMInventoryCompsInterface.OperationSeqNumber"
        dbInteger "ColumnWidth" ="2040"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ComponentQuantity"
        dbInteger "ColumnWidth" ="1905"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ItemSeqNumberNew"
        dbInteger "ColumnWidth" ="1950"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ComponentItemNumber"
        dbInteger "ColumnWidth" ="2220"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OpSeqNumberNew"
        dbInteger "ColumnWidth" ="1950"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbInteger "ColumnWidth" ="1950"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ResourceSeqNumber"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xbb5f1fbd1a8b0e458f8c0f984b66c64f
        End
    End
    Begin
        dbText "Name" ="ResourceCode"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xcc60f66b9ddc3e48adb6c602695dd4ae
        End
    End
    Begin
        dbText "Name" ="AssignedUnits"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc6d6ee8184397f418f1a9ade1d57b766
        End
    End
End

dbMemo "SQL" ="SELECT tblTestBOMOpResourcesInterface.AssemblyItemNumber, tblNetworkBOMOpResourc"
    "esInterface.AssemblyItemNumber, tblTestBOMOpResourcesInterface.OperationSeqNumbe"
    "r, tblNetworkBOMOpResourcesInterface.OperationSeqNumber, tblTestBOMOpResourcesIn"
    "terface.ResourceSeqNumber, tblNetworkBOMOpResourcesInterface.ResourceSeqNumber, "
    "tblTestBOMOpResourcesInterface.ResourceCode, tblNetworkBOMOpResourcesInterface.R"
    "esourceCode, tblTestBOMOpResourcesInterface.UsageRateOrAmount, tblNetworkBOMOpRe"
    "sourcesInterface.UsageRateOrAmount, tblTestBOMOpResourcesInterface.UsageRateOrAm"
    "ountInverse, tblNetworkBOMOpResourcesInterface.UsageRateOrAmountInverse, IIf(tbl"
    "TestBOMOpResourcesInterface!ResourceCode<>tblNetworkBOMOpResourcesInterface!Reso"
    "urceCode,-1,0) AS Expr1, tblTestBOMOpResourcesInterface!UsageRateOrAmount-tblNet"
    "workBOMOpResourcesInterface!UsageRateOrAmount AS Expr2, tblTestBOMOpResourcesInt"
    "erface!UsageRateOrAmountInverse-tblNetworkBOMOpResourcesInterface!UsageRateOrAmo"
    "untInverse AS Expr3\015\012FROM tblTestBOMOpResourcesInterface LEFT JOIN tblNetw"
    "orkBOMOpResourcesInterface ON (tblTestBOMOpResourcesInterface.OrganizationCode=t"
    "blNetworkBOMOpResourcesInterface.OrganizationCode) AND (tblTestBOMOpResourcesInt"
    "erface.AssemblyItemNumber=tblNetworkBOMOpResourcesInterface.AssemblyItemNumber) "
    "AND (tblTestBOMOpResourcesInterface.OperationSeqNumber=tblNetworkBOMOpResourcesI"
    "nterface.OperationSeqNumber) AND (tblTestBOMOpResourcesInterface.ResourceSeqNumb"
    "er=tblNetworkBOMOpResourcesInterface.ResourceSeqNumber);\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x689aa246e6552f47ab441ab8bd91ded7
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc8b674cc043c7142a5a8bea6e65047c6
        End
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x74be29f9f6311a4eade5d654441b2de3
        End
    End
    Begin
        dbText "Name" ="Expr3"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc37d6aa1643de646ae76a7ccb16a5579
        End
    End
End

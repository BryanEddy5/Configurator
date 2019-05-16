dbMemo "SQL" ="SELECT tblTestBOMOpSequencesInterface.AssemblyItemNumber, tblNetworkBOMOpSequenc"
    "esInterface.AssemblyItemNumber, tblTestBOMOpSequencesInterface.OperationSeqNum, "
    "tblNetworkBOMOpSequencesInterface.OperationSeqNum, tblTestBOMOpSequencesInterfac"
    "e.OperationCode, tblNetworkBOMOpSequencesInterface.OperationCode, tblTestBOMOpSe"
    "quencesInterface.DepartmentCode, tblNetworkBOMOpSequencesInterface.DepartmentCod"
    "e, tblTestBOMOpSequencesInterface.OperationDescription, tblNetworkBOMOpSequences"
    "Interface.OperationDescription\015\012FROM tblNetworkBOMOpSequencesInterface RIG"
    "HT JOIN tblTestBOMOpSequencesInterface ON (tblNetworkBOMOpSequencesInterface.Org"
    "anizationCode=tblTestBOMOpSequencesInterface.OrganizationCode) AND (tblNetworkBO"
    "MOpSequencesInterface.AssemblyItemNumber=tblTestBOMOpSequencesInterface.Assembly"
    "ItemNumber) AND (tblNetworkBOMOpSequencesInterface.OperationSeqNum=tblTestBOMOpS"
    "equencesInterface.OperationSeqNum);\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xf57915fb02c889498302482fd85ab936
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
End

dbMemo "SQL" ="SELECT tblOracleBOMOpSequencesInterface.OperationSeqNum, tblOracleBOMRoutingsAtt"
    "achments.SeqNumber, tblOracleBOMOpSequencesInterface.DepartmentCode, tblOracleBO"
    "MOpSequencesInterface.OperationDescription, tblOracleBOMRoutingsAttachments.Shor"
    "tText\015\012FROM tblOracleBOMOpSequencesInterface INNER JOIN tblOracleBOMRoutin"
    "gsAttachments ON (tblOracleBOMOpSequencesInterface.OrganizationCode=tblOracleBOM"
    "RoutingsAttachments.OrganizationCode) AND (tblOracleBOMOpSequencesInterface.Asse"
    "mblyItemNumber=tblOracleBOMRoutingsAttachments.AssemblyItemNumber) AND (tblOracl"
    "eBOMOpSequencesInterface.OperationSeqNum=tblOracleBOMRoutingsAttachments.Operati"
    "onSeqNumber)\015\012ORDER BY tblOracleBOMOpSequencesInterface.OperationSeqNum, t"
    "blOracleBOMRoutingsAttachments.SeqNumber;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xb1d1f17122aefb49baeb76a7e3654526
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
End

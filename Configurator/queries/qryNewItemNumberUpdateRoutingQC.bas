dbMemo "SQL" ="UPDATE tblBOMOpSequencesInterface SET tblBOMOpSequencesInterface.OperationCode ="
    " \"Q\" & Format(Val(Mid(Forms!frmItemInterface!frmPremiseItemDetails.Form!Design"
    "Code,3,3)),\"000\")\015\012WHERE (((tblBOMOpSequencesInterface.OperationCode) Li"
    "ke \"Q*\"));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xb57b29e5c8387a4da3facff5353d6713
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
End

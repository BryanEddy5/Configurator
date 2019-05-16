dbMemo "SQL" ="SELECT tblCableConstructionReferences.Base, tblCableConstructionReferences.Setup"
    "ID, tblBaseItemsCheckBox_local.Selected\015\012FROM tblCableConstructionReferenc"
    "es INNER JOIN tblBaseItemsCheckBox_local ON tblCableConstructionReferences.Setup"
    "ID = tblBaseItemsCheckBox_local.SetupID;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xa826a3dd8ec55f478d17d165cd30113d
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblCableConstructionReferences.Base"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.SetupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblBaseItemsCheckBox_local.Selected"
        dbLong "AggregateType" ="-1"
    End
End

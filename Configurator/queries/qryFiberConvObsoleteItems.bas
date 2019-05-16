dbMemo "SQL" ="UPDATE (tblFiberConversion_CopyingOldItemsToNew INNER JOIN qry_SubsBase ON tblFi"
    "berConversion_CopyingOldItemsToNew.[Old BaseItem] = qry_SubsBase.[OracleBase#]) "
    "INNER JOIN BasicProductConstruction_LIVE ON qry_SubsBase.[New Oracle Part #] = B"
    "asicProductConstruction_LIVE.[New Oracle Part #] SET BasicProductConstruction_LI"
    "VE.Active = No, BasicProductConstruction_LIVE.OracleStatus = \"Obsolete\", Basic"
    "ProductConstruction_LIVE.Reason = Date() & \" - \" & Forms!frmFiberReplace!Reaso"
    "nDeactivate & \" \" & Forms!frmSwitchBoard!Initials & \"; \" & BasicProductConst"
    "ruction_LIVE!Reason\015\012WHERE (((BasicProductConstruction_LIVE.[New Oracle Pa"
    "rt #]) Not Like \"*series*\" And (BasicProductConstruction_LIVE.[New Oracle Part"
    " #]) Not Like \"*test*\"));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x271a3f77bb8dc74fa037684f10dea3dd
End
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="[Basic Product Construction].Reason"
        dbInteger "ColumnWidth" ="9435"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BasicProductConstruction_LIVE.Reason"
        dbInteger "ColumnWidth" ="9345"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End

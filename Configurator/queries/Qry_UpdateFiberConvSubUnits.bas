dbMemo "SQL" ="UPDATE DISTINCTROW tblFibersToConvert, FiberConversionFinalAssemblyList INNER JO"
    "IN [Basic Product Construction] ON FiberConversionFinalAssemblyList.AssemblyItem"
    "Number=[Basic Product Construction].[New Oracle Part #] SET [Basic Product Const"
    "ruction].[Revision Date] = Date(), [Basic Product Construction].[Revision Letter"
    "] = IIf([Revision Letter]='A','B',IIf([Revision Letter]='B','C',IIf([Revision Le"
    "tter]='C','D',IIf([Revision Letter]='D','E',IIf([Revision Letter]='E','F',IIf([R"
    "evision Letter]='F','G',IIf([Revision Letter]='G','H',IIf([Revision Letter]='H',"
    "'I',IIf([Revision Letter]='I','J',\"\"))))))))), [Basic Product Construction].[F"
    "iber Type] = tblFibersToConvert!Description, [Basic Product Construction].[Revis"
    "ion History] = Date() & \"\" & \" -\" & tblFibersToConvert!Reason & \"-\" & [Ini"
    "tials] & \" \" & [Revision History], [Basic Product Construction].[Fiber Oracle "
    "item] = IIf([Fiber Oracle item]=Null,tblFibersToConvert!NewFiberNumber,IIf([Fibe"
    "r Oracle item]=[OldFiberNumber],Replace([Fiber Oracle item],tblFibersToConvert!O"
    "ldFiberNumber,tblFibersToConvert!NewFiberNumber),[NewFiberNumber]))\015\012WHERE"
    " (((tblFibersToConvert.Include)=-1));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x96fc4c29bff6824bb09b6ed6ee425a02
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[New Oracle Part Number]"
        dbInteger "ColumnWidth" ="2340"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[Revision History]"
        dbInteger "ColumnWidth" ="7140"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic (Cable/Sheath)].[Fiber Type A]"
        dbInteger "ColumnWidth" ="12150"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Fiber Type]"
        dbInteger "ColumnWidth" ="4545"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Revision History]"
        dbInteger "ColumnWidth" ="20835"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End

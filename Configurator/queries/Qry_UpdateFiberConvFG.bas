dbMemo "SQL" ="UPDATE DISTINCTROW tblFibersToConvert, FiberConversionFinalAssemblyList INNER JO"
    "IN [Basic (Cable/Sheath)] ON FiberConversionFinalAssemblyList.AssemblyItemNumber"
    " = [Basic (Cable/Sheath)].[Item No] SET [Basic (Cable/Sheath)].[Revision Date] ="
    " Date(), [Basic (Cable/Sheath)].[Revision Letter] = IIf([Basic (Cable/Sheath)].["
    "Revision Letter]='A','B',IIf([Basic (Cable/Sheath)].[Revision Letter]='B','C',II"
    "f([Basic (Cable/Sheath)].[Revision Letter]='C','D',IIf([Basic (Cable/Sheath)].[R"
    "evision Letter]='D','E',IIf([Basic (Cable/Sheath)].[Revision Letter]='E','F',IIf"
    "([Basic (Cable/Sheath)].[Revision Letter]='F','G',IIf([Basic (Cable/Sheath)].[Re"
    "vision Letter]='G','H',IIf([Basic (Cable/Sheath)].[Revision Letter]='H','I',IIf("
    "[Basic (Cable/Sheath)].[Revision Letter]='I','J',IIf([Basic (Cable/Sheath)].[Rev"
    "ision Letter]='J','K',\"\")))))))))), [Basic (Cable/Sheath)].[Fiber Type A] = tb"
    "lFibersToConvert!Description, [Basic (Cable/Sheath)].[Revision History] = Date()"
    " & \"\" & \" -\" & tblFibersToConvert!Reason & \"-\" & [Initials] & \" \" & [Rev"
    "ision History]\015\012WHERE (((tblFibersToConvert.Include)=-1));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xefe99dd2fcb5274ba884bb38d6a89e0a
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
End

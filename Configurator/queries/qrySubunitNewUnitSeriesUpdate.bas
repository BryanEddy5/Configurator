Operation =4
Option =0
Where ="(((tblUnitSeriesCopy.[TB Color Series])<>\"FILLER\"))"
Begin InputTables
    Name ="tblUnitSeriesCopy"
    Name ="tblNewPrductConstruction"
End
Begin OutputColumns
    Name ="tblUnitSeriesCopy.[Unit Type]"
    Expression ="[Item No]"
    Name ="tblUnitSeriesCopy.[Fiber Code]"
    Expression ="IIf(Forms!frmfibermode![OM#] Like \"*SM*\",\"SM\" & Mid(Forms!frmCreateItem!Item"
        "No,6,1),\"MM\" & Mid(Forms!frmCreateItem!ItemNo,6,1))"
    Name ="tblUnitSeriesCopy.[Print type]"
    Expression ="IIf(Forms!frmCreateItem!StripePrint=1,\"SPECIAL: One (1) BLACK STRIPE\",IIf(Mid("
        "[Item No],10,1)=\"8\",\"SPECIAL WHITE PRINT\",\"Standard Black Print\"))"
    Name ="tblUnitSeriesCopy.OracleItem"
    Expression ="[New Oracle Part #]"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0xb5a25e1126b989488e9133833aec0283
End
Begin
    Begin
        dbText "Name" ="tblUnitSeriesCopy.[Unit Type]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblUnitSeriesCopy.[Fiber Code]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblUnitSeriesCopy.[Print type]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblUnitSeriesCopy.[TB Color Series]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblUnitSeriesCopy.OracleItem"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =65
    Top =452
    Right =1174
    Bottom =870
    Left =-1
    Top =-1
    Right =1077
    Bottom =119
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =38
        Top =6
        Right =292
        Bottom =199
        Top =0
        Name ="tblUnitSeriesCopy"
        Name =""
    End
    Begin
        Left =362
        Top =4
        Right =720
        Bottom =197
        Top =0
        Name ="tblNewPrductConstruction"
        Name =""
    End
End

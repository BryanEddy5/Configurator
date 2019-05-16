Operation =4
Option =0
Where ="(((tblUnitSeriesCopy.[TB Color Series])<>\"Filler\") And ((tblUnitSeriesCopy.[Un"
    "it ID]) Like \"*\" & Format(Forms!frmCreateItem!SubCount,\"00\") & \"*\"))"
Begin InputTables
    Name ="tblUnitSeriesCopy"
    Name ="tblNewPrductConstruction"
End
Begin OutputColumns
    Name ="tblUnitSeriesCopy.Unit Type"
    Expression ="[Item No]"
    Name ="tblUnitSeriesCopy.Fiber Code"
    Expression ="IIf(Forms!frmfibermode![OM#] Like \"*SM*\",\"SM\" & Forms!frmCreateItem!FiberTyp"
        "e,\"MM\" & Forms!frmCreateItem!FiberType)"
    Name ="tblUnitSeriesCopy.Print type"
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
    0xc9bca8a9e70d774083b6ec776a97f7f4
End
Begin
    Begin
        dbText "Name" ="tblUnitSeriesCopy.[Print type]"
        dbInteger "ColumnWidth" ="2760"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblUnitSeriesCopy.[Unit Type]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblUnitSeriesCopy.[TB Color Series]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblUnitSeriesCopy.[Unit ID]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblUnitSeriesCopy.OracleItem"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =172
    Top =244
    Right =1281
    Bottom =873
    Left =-1
    Top =-1
    Right =1077
    Bottom =136
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

Operation =4
Option =0
Where ="(((tblUnitSeriesCopy.[TB Color Series])<>\"filler\"))"
Begin InputTables
    Name ="tblUnitSeriesCopy"
    Name ="tblUnitSeriesSubunit"
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
    0xac25f5c13f2f1141a6c768c70f5bf781
End
Begin
    Begin
        dbText "Name" ="tblUnitSeriesCopy.[Unit Type]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblUnitSeriesCopy.[Fiber Code]"
        dbInteger "ColumnWidth" ="1425"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblUnitSeriesCopy.[Print type]"
        dbInteger "ColumnWidth" ="2850"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblUnitSeriesCopy.[TB Color Series]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblUnitSeriesSubunit.[New Oracle Part #]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblUnitSeriesCopy.OracleItem"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =138
    Top =74
    Right =1247
    Bottom =493
    Left =-1
    Top =-1
    Right =1077
    Bottom =190
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =503
        Top =12
        Right =788
        Bottom =189
        Top =0
        Name ="tblUnitSeriesCopy"
        Name =""
    End
    Begin
        Left =173
        Top =13
        Right =435
        Bottom =161
        Top =0
        Name ="tblUnitSeriesSubunit"
        Name =""
    End
End

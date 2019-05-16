Operation =4
Option =0
Where ="(((tblUnitSeriesCopy.[TB Color Series]) Not Like \"*filler*\") And ((tblUnitSeri"
    "esCopy.UnitNumber)=Forms!frmCreateItem!SubCount Or (tblUnitSeriesCopy.UnitNumber"
    ")=100))"
Begin InputTables
    Name ="tblUnitSeriesCopy"
    Name ="tblUnitSeriesSubunit"
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
    0x3560616a5563cc41816fa26843f5cc15
End
Begin
    Begin
        dbText "Name" ="tblUnitSeriesCopy.[Unit Type]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblUnitSeriesCopy.OracleItem"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblUnitSeriesCopy.[TB Color Series]"
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
        dbText "Name" ="tblUnitSeriesCopy.[Unit ID]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblUnitSeriesCopy.UnitNumber"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =19
    Top =177
    Right =1128
    Bottom =725
    Left =-1
    Top =-1
    Right =1077
    Bottom =103
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =503
        Top =12
        Right =764
        Bottom =160
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

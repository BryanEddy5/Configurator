Operation =3
Name ="Unit Series"
Option =0
Begin InputTables
    Name ="tblUnitSeriesCopy"
End
Begin OutputColumns
    Name ="Unit Color Series"
    Expression ="tblUnitSeriesCopy.[Unit Color Series]"
    Name ="Unit ID"
    Expression ="tblUnitSeriesCopy.[Unit ID]"
    Name ="Description"
    Expression ="tblUnitSeriesCopy.Description"
    Name ="TB Color Series"
    Expression ="tblUnitSeriesCopy.[TB Color Series]"
    Name ="Print type"
    Expression ="tblUnitSeriesCopy.[Print type]"
    Name ="Fiber Code"
    Expression ="tblUnitSeriesCopy.[Fiber Code]"
    Name ="Print Spacing"
    Expression ="tblUnitSeriesCopy.[Print Spacing]"
    Name ="verified"
    Expression ="tblUnitSeriesCopy.verified"
    Name ="OracleItem"
    Expression ="tblUnitSeriesCopy.OracleItem"
    Name ="UnitNumber"
    Expression ="tblUnitSeriesCopy.UnitNumber"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0xac547af0c9ae9b48a242527dcad5c6d8
End
Begin
    Begin
        dbText "Name" ="tblUnitSeriesCopy.[Unit Color Series]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblUnitSeriesCopy.OracleItem"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblUnitSeriesCopy.Description"
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
        dbText "Name" ="tblUnitSeriesCopy.[Unit Type]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblUnitSeriesCopy.[Print type]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblUnitSeriesCopy.[Fiber Code]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblUnitSeriesCopy.[Print Spacing]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblUnitSeriesCopy.verified"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblUnitSeriesCopy.UnitNumber"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1292
    Bottom =772
    Left =-1
    Top =-1
    Right =1260
    Bottom =97
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =38
        Top =6
        Right =362
        Bottom =124
        Top =0
        Name ="tblUnitSeriesCopy"
        Name =""
    End
End

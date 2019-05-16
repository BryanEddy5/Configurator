Operation =4
Option =0
Where ="(((tblItemCatalogInterface.ElementName)=\"Fiber Type\"))"
Begin InputTables
    Name ="tblItemCatalogInterface"
End
Begin OutputColumns
    Name ="tblItemCatalogInterface.ElementValue"
    Expression ="\"MIXED\""
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0xdcdc3ae89afc1447870846d0fa0accf9
End
Begin
    Begin
        dbText "Name" ="tblItemCatalogInterface.ElementValue"
        dbInteger "ColumnWidth" ="2955"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =76
    Top =376
    Right =1185
    Bottom =760
    Left =-1
    Top =-1
    Right =1077
    Bottom =170
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =38
        Top =6
        Right =270
        Bottom =154
        Top =0
        Name ="tblItemCatalogInterface"
        Name =""
    End
End

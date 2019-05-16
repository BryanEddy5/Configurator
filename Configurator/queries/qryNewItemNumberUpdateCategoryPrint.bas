Operation =4
Option =8
Where ="(((tblItemCategoryInterface.CategorySetName)=\"Design\"))"
Begin InputTables
    Name ="tblItemCategoryInterface"
End
Begin OutputColumns
    Name ="tblItemCategoryInterface.CategoryName"
    Expression ="\"PRINT\""
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x24de1f9217b67e459d37f9d1900701ef
End
Begin
    Begin
        dbText "Name" ="tblItemCategoryInterface.CategoryName"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =19
    Top =257
    Right =1128
    Bottom =741
    Left =-1
    Top =-1
    Right =1077
    Bottom =253
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =471
        Top =12
        Right =694
        Bottom =175
        Top =0
        Name ="tblItemCategoryInterface"
        Name =""
    End
End

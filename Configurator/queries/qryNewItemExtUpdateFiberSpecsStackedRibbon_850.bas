Operation =4
Option =0
Where ="(((tblQAItemSpecElements.ElementName)=\"FIBER1_850ATTN\"))"
Begin InputTables
    Name ="tblQAItemSpecElements"
End
Begin OutputColumns
    Name ="tblQAItemSpecElements.UpperSpecLimit"
    Expression ="\"4\""
    Name ="tblQAItemSpecElements.UpperUserDefinedLimit"
    Expression ="\"4\""
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0xfa4ce68ad1099c478f470ec840a386b6
End
Begin
    Begin
        dbText "Name" ="tblQAItemSpecElements.ElementName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblQAItemSpecElements.UpperSpecLimit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblQAItemSpecElements.UpperUserDefinedLimit"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =99
    Top =77
    Right =1208
    Bottom =706
    Left =-1
    Top =-1
    Right =1077
    Bottom =571
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =281
        Top =174
        Right =716
        Bottom =390
        Top =0
        Name ="tblQAItemSpecElements"
        Name =""
    End
End

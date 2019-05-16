Operation =1
Option =0
Having ="(((Sum(IIf([Description] Is Null,1,0)))<>0))"
Begin InputTables
    Name ="tblItemInterface"
End
Begin OutputColumns
    Alias ="CheckCode"
    Expression ="\"Item - Blank Item Description\""
    Alias ="DescriptionCheck"
    Expression ="Sum(IIf([Description] Is Null,1,0))"
    Alias ="StopLoad"
    Expression ="-1"
End
Begin Groups
    Expression ="\"Item - Blank Item Description\""
    GroupLevel =0
    Expression ="-1"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbBinary "GUID" = Begin
    0x929f971c38d19c43a3860f63c458093b
End
Begin
    Begin
        dbText "Name" ="CheckCode"
        dbInteger "ColumnWidth" ="5520"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x0b4af52958982b458633cb013bcc5a07
        End
    End
    Begin
        dbText "Name" ="DescriptionCheck"
        dbInteger "ColumnWidth" ="1710"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2ef87cbfa9427e40a6434f4ab0374105
        End
    End
    Begin
        dbText "Name" ="StopLoad"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd93272555bbf874c9411a79e9661f961
        End
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1560
    Bottom =854
    Left =-1
    Top =-1
    Right =1528
    Bottom =93
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =750
        Top =3
        Right =1017
        Bottom =117
        Top =0
        Name ="tblItemInterface"
        Name =""
    End
End

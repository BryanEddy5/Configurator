Operation =1
Option =0
Where ="(((AFLPRD_INVItmCatg_CAB.ItemNumber)=[Forms]![frmMainMenu]![ItemNumberSelect]))"
Begin InputTables
    Name ="AFLPRD_INVItmCatg_CAB"
End
Begin OutputColumns
    Expression ="AFLPRD_INVItmCatg_CAB.*"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x5631b06c0610e34e8fe9ce15db98fe86
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="AFLPRD_INVItmCatg_CAB.*"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AFLPRD_INVItmCatg_CAB.ItemNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AFLPRD_INVItmCatg_CAB.OrganizationCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AFLPRD_INVItmCatg_CAB.CategorySetName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AFLPRD_INVItmCatg_CAB.CategoryName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AFLPRD_INVItmCatg_CAB.AllowOverride"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AFLPRD_INVItmCatg_CAB.LimitToList"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AFLPRD_INVItmCatg_CAB.ItemCopied"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AFLPRD_INVItmCatg_CAB.Number"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =129
    Top =193
    Right =1638
    Bottom =979
    Left =-1
    Top =-1
    Right =1477
    Bottom =93
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =353
        Top =12
        Right =497
        Bottom =156
        Top =0
        Name ="AFLPRD_INVItmCatg_CAB"
        Name =""
    End
End

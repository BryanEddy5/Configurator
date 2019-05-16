Operation =1
Option =18
RowCount ="1,"
Where ="((([Basic Product Construction].[Item No]) Like Left([Unit Type],1) & \"*\") AND"
    " (([Basic Product Construction].[New Oracle Part #]) Like \"pe*\" Or ([Basic Pro"
    "duct Construction].[New Oracle Part #]) Like \"pt*\"))"
Begin InputTables
    Name ="Basic Product Construction"
    Name ="qryUnitSeriesCopy"
End
Begin OutputColumns
    Expression ="[Basic Product Construction].[Item No]"
    Expression ="[Basic Product Construction].[New Oracle Part #]"
    Alias ="NewOracle"
    Expression ="Left([New Oracle Part #],2) & Mid([New Oracle Part #],3,5)+8 & \"-00\""
End
Begin OrderBy
    Expression ="[Basic Product Construction].[New Oracle Part #]"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x6a25f131ffcbe34e89c6f30c5b248c9a
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="NewOracle"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc9edb86d9dc4a74fae88fcb0ff687d51
        End
    End
End
Begin
    State =0
    Left =157
    Top =276
    Right =1266
    Bottom =742
    Left =-1
    Top =-1
    Right =1077
    Bottom =235
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =401
        Top =4
        Right =715
        Bottom =212
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
    Begin
        Left =61
        Top =10
        Right =272
        Bottom =191
        Top =0
        Name ="qryUnitSeriesCopy"
        Name =""
    End
End

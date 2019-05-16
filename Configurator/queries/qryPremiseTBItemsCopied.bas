Operation =1
Option =0
Where ="(((tblItemLoadTempBOM.ComponentItemNumber) Like \"pt*\" Or (tblItemLoadTempBOM.C"
    "omponentItemNumber) Like \"rbn*\"))"
Begin InputTables
    Name ="tblItemLoadTempBOM"
    Name ="Basic Product Construction"
End
Begin OutputColumns
    Expression ="tblItemLoadTempBOM.ComponentItemNumber"
    Alias ="ItemNumberSuffix"
    Expression ="IIf([ComponentItemNumber] Like \"rbn*\",\"rbn\",(Mid([ComponentItemNumber],9,2))"
        ")"
    Expression ="[Basic Product Construction].[Item No]"
    Expression ="[Basic Product Construction].[Jacket Material]"
    Expression ="[Basic Product Construction].[Fiber Oracle item]"
    Expression ="[Basic Product Construction].[Color Chip ID]"
    Expression ="[Basic Product Construction].[EZ Strip]"
    Expression ="[Basic Product Construction].[TB Material]"
    Expression ="[Basic Product Construction].[TB Nominal OD]"
End
Begin Joins
    LeftTable ="tblItemLoadTempBOM"
    RightTable ="Basic Product Construction"
    Expression ="tblItemLoadTempBOM.ComponentItemNumber = [Basic Product Construction].[New Oracl"
        "e Part #]"
    Flag =1
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
    0xb4e224004c122500d08f8001ffffffff
End
Begin
    Begin
        dbText "Name" ="ItemNumberSuffix"
        dbInteger "ColumnWidth" ="2505"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x03092868245e5542a5f4ad3920832791
        End
    End
    Begin
        dbText "Name" ="tblItemLoadTempBOM.ComponentItemNumber"
        dbInteger "ColumnWidth" ="2640"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Fiber Oracle item]"
        dbInteger "ColumnWidth" ="1875"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[TB Material]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Item No]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[TB Nominal OD]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Jacket Material]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Color Chip ID]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[EZ Strip]"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =53
    Top =48
    Right =1162
    Bottom =572
    Left =-1
    Top =-1
    Right =1077
    Bottom =293
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =325
        Bottom =274
        Top =0
        Name ="tblItemLoadTempBOM"
        Name =""
    End
    Begin
        Left =363
        Top =6
        Right =751
        Bottom =289
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
End

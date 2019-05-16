Operation =1
Option =0
Where ="((([tblItemInterfaceCategoryValues].[CategorySetName]) Like \"*new product*\"))"
Begin InputTables
    Name ="tblItemInterfaceCategoryValues"
End
Begin OutputColumns
    Alias ="Expr1"
    Expression ="tblItemInterfaceCategoryValues.CategorySetName"
    Alias ="Expr2"
    Expression ="tblItemInterfaceCategoryValues.DesignTypeID"
    Alias ="Expr3"
    Expression ="tblItemInterfaceCategoryValues.CategoryValue"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xe1099d0946a9cb4a9be99a0e641fc964
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xed2c99601829b7489e6844fd5950ae2f
        End
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x8f5a613d1151d842a9da8ba2bea44b35
        End
    End
    Begin
        dbText "Name" ="Expr3"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xbbd0b3e99a0dd64ba5fdbc403aad2e93
        End
    End
End
Begin
    State =0
    Left =290
    Top =256
    Right =1903
    Bottom =742
    Left =-1
    Top =-1
    Right =1581
    Bottom =272
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =299
        Bottom =184
        Top =0
        Name ="tblItemInterfaceCategoryValues"
        Name =""
    End
End

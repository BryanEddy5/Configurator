Operation =1
Option =0
Begin InputTables
    Name ="qryProductTypesALL"
End
Begin OutputColumns
    Alias ="Expr1"
    Expression ="qryProductTypesALL.ProductType"
    Alias ="Expr2"
    Expression ="qryProductTypesALL.ProductTypeView"
    Alias ="Expr3"
    Expression ="qryProductTypesALL.CategorySetName"
End
Begin OrderBy
    Expression ="qryProductTypesALL.ProductTypeView"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x6f91cb1e4e09fe468d6010495377be52
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x03b082731e3639448fd89a4fe68e6d8b
        End
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x53cf35e6b5054b4b8fe90fd512a5f530
        End
    End
    Begin
        dbText "Name" ="Expr3"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf655d72473f3554993724f765cbf03e9
        End
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1109
    Bottom =669
    Left =-1
    Top =-1
    Right =1077
    Bottom =110
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =301
        Top =6
        Right =485
        Bottom =105
        Top =0
        Name ="qryProductTypesALL"
        Name =""
    End
End

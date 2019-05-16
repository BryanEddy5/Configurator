Operation =3
Name ="tblItemInterfaceCategoryValues"
Option =0
Begin InputTables
    Name ="qryfrmCategoryNew"
End
Begin OutputColumns
    Alias ="Expr1"
    Name ="CategorySetName"
    Expression ="qryfrmCategoryNew.CategorySetName"
    Alias ="Expr2"
    Name ="CategoryValue"
    Expression ="qryfrmCategoryNew.CategoryValue"
    Alias ="Expr3"
    Name ="DesignTypeID"
    Expression ="qryfrmCategoryNew.DesignTypeID"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0xdfaec7e588c0234783b8b6863d640b4c
End
Begin
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xdb329cb8596bb64aa848d3c0f6892f61
        End
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x0d6f073565c5d24583111222ebc5551e
        End
    End
    Begin
        dbText "Name" ="Expr3"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xfe63ddf9518c294ba073f7e63858e2d4
        End
    End
End
Begin
    State =0
    Left =263
    Top =311
    Right =1885
    Bottom =820
    Left =-1
    Top =-1
    Right =1590
    Bottom =180
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =38
        Top =6
        Right =377
        Bottom =209
        Top =0
        Name ="qryfrmCategoryNew"
        Name =""
    End
End

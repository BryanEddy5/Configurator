Operation =1
Option =0
Begin InputTables
    Name ="tblOracleDepartmentCodes"
End
Begin OutputColumns
    Alias ="Expr1"
    Expression ="tblOracleDepartmentCodes.DepartmentCode"
    Alias ="Expr2"
    Expression ="tblOracleDepartmentCodes.DepartmentCodeDescription"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xb6e9b2d2f3428742804093578d25e1c6
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc106bab341b25341b2a37942db9964d4
        End
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x936fca84a332714f8cae4f053b8f8ba8
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
        Left =38
        Top =6
        Right =196
        Bottom =105
        Top =0
        Name ="tblOracleDepartmentCodes"
        Name =""
    End
End

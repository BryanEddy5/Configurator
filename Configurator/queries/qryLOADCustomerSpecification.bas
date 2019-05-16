Operation =3
Name ="tblCustomerSpecs"
Option =0
Begin InputTables
    Name ="tblCustomersSpecifications"
End
Begin OutputColumns
    Alias ="Expr1"
    Name ="CustomerSpec"
    Expression ="tblCustomersSpecifications.CustomerSpec"
    Alias ="Expr2"
    Name ="CustomerPartNumber"
    Expression ="tblCustomersSpecifications.CustomerPartNumber"
    Alias ="Expr3"
    Name ="Customer"
    Expression ="tblCustomersSpecifications.Customer"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0xd0dafef08dbd464ba9f2e31f9ba8b833
End
Begin
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa2eee20714f1e74a96508b3b2963151e
        End
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x253c0d1b1dcdb2409e610424b503e6bf
        End
    End
    Begin
        dbText "Name" ="Expr3"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x07f1c541b18c9e47a2cc3bd2f53d7fd7
        End
    End
End
Begin
    State =0
    Left =70
    Top =378
    Right =1362
    Bottom =702
    Left =-1
    Top =-1
    Right =1260
    Bottom =93
    Left =1714
    Top =0
    ColumnsShown =651
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =109
        Top =0
        Name ="tblCustomersSpecifications"
        Name =""
    End
End

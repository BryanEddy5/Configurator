Operation =5
Option =0
Where ="(((tblCustomerSpecs.CustomerSpec)=Forms!frmCustomerSpecs!CustomerSpec) And ((tbl"
    "CustomerSpecs.CustomerPartNumber)=Forms!frmCustomerSpecs!CustomerPartNumber) And"
    " ((tblCustomerSpecs.Customer)=Forms!frmCustomerSpecs!Customer))"
Begin InputTables
    Name ="tblCustomerSpecs"
End
Begin OutputColumns
    Expression ="tblCustomerSpecs.CustomerSpec"
    Expression ="tblCustomerSpecs.CustomerPartNumber"
    Expression ="tblCustomerSpecs.Customer"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0xdc33aebe8cc0a443861006af9b2c4907
End
Begin
    Begin
        dbText "Name" ="tblCustomerSpecs.CustomerSpec"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCustomerSpecs.CustomerPartNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCustomerSpecs.Customer"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-5
    Top =237
    Right =1496
    Bottom =827
    Left =-1
    Top =-1
    Right =1469
    Bottom =258
    Left =0
    Top =0
    ColumnsShown =771
    Begin
        Left =48
        Top =12
        Right =287
        Bottom =156
        Top =0
        Name ="tblCustomerSpecs"
        Name =""
    End
End

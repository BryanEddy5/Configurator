Operation =1
Option =2
Where ="(((qryFiberItemCustomer.Customer) Not Like \"*CORE*\"))"
Begin InputTables
    Name ="qryFiberItemCustomer"
End
Begin OutputColumns
    Expression ="qryFiberItemCustomer.Customer"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xd8b639c1310ca944a1be1e141ea2cf59
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="qryFiberItemCustomer.Customer"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-17
    Top =181
    Right =1768
    Bottom =602
    Left =-1
    Top =-1
    Right =1753
    Bottom =173
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =49
        Top =18
        Right =331
        Bottom =164
        Top =0
        Name ="qryFiberItemCustomer"
        Name =""
    End
End

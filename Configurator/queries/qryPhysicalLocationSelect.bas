Operation =1
Option =0
Where ="((([tblManufacturingLocations].[Active])=-1))"
Begin InputTables
    Name ="tblManufacturingLocations"
End
Begin OutputColumns
    Alias ="Expr1"
    Expression ="tblManufacturingLocations.LocationID"
    Alias ="Expr2"
    Expression ="tblManufacturingLocations.Location"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xf21eaff153518944810395d4614364be
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xff1df3703b7a604da286fe80d1dbbd51
        End
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x8e68a6388522cc4492a970b37b110c89
        End
    End
End
Begin
    State =0
    Left =18
    Top =40
    Right =1127
    Bottom =360
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
        Right =134
        Bottom =90
        Top =0
        Name ="tblManufacturingLocations"
        Name =""
    End
End

Operation =1
Option =2
Where ="((([tblOracleBOMOpSequencesInterface].[DepartmentCode]) Like \"*dossert*\"))"
Begin InputTables
    Name ="tblOracleBOMOpSequencesInterface"
End
Begin OutputColumns
    Alias ="DossertItem"
    Expression ="tblOracleBOMOpSequencesInterface.AssemblyItemNumber"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x68ca14f5aa294f4eac4a15ba0da0cc13
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="DossertItem"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x643f626277cac347884c90ad7ef511e3
        End
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1560
    Bottom =854
    Left =-1
    Top =-1
    Right =1528
    Bottom =260
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =296
        Top =6
        Right =502
        Bottom =210
        Top =0
        Name ="tblOracleBOMOpSequencesInterface"
        Name =""
    End
End

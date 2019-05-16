Operation =3
Name ="tblCustomerSpec"
Option =0
Where ="((([Forms]![frmCreateItem]![PID])=[Forms]![frmCreateItem]![PID]))"
Begin InputTables
End
Begin OutputColumns
    Alias ="Spec"
    Name ="CustomerSpec"
    Expression ="[Forms]![frmCreateItem]![Spec]"
    Alias ="PID"
    Name ="CustomerPartNumber"
    Expression ="[Forms]![frmCreateItem]![PID]"
    Alias ="Customer"
    Name ="Customer"
    Expression ="[Forms]![frmCreateItem]![Customer]"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xa7428f893344dc41a64791ba12601138
End
dbBoolean "UseTransaction" ="-1"
Begin
    Begin
        dbText "Name" ="Spec"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4b3fe15f4343b64a9a9947d72dd2d428
        End
    End
    Begin
        dbText "Name" ="PID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3289bdab53fcac4f9eea8b84cb1a1e1c
        End
    End
    Begin
        dbText "Name" ="Customer"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x5c91da4b1bf66145b54586b49b8f486d
        End
    End
End
Begin
    State =0
    Left =104
    Top =72
    Right =1396
    Bottom =396
    Left =-1
    Top =-1
    Right =1260
    Bottom =93
    Left =1757
    Top =0
    ColumnsShown =651
End

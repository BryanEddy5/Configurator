Operation =1
Option =0
Where ="((([Basic Product Construction].[New Oracle Part #]) Not Like \"*test*\"))"
Begin InputTables
    Name ="Basic Product Construction"
End
Begin OutputColumns
    Alias ="Oracle"
    Expression ="[Basic Product Construction].[New Oracle Part #]"
    Expression ="[Basic Product Construction].[Item No]"
    Expression ="[Basic Product Construction].Base"
    Alias ="Label"
    Expression ="[Basic Product Construction].[LISTING COMPANY]"
    Expression ="[Basic Product Construction].Active"
    Expression ="[Basic Product Construction].Reason"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x6f09723dfd7baa4392c61ba481c655f8
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbByte "RecordsetType" ="0"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Oracle"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x998dd9c893095a4fa22e84fd5dbb600c
        End
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Item No]"
        dbInteger "ColumnWidth" ="3900"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Label"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x678520e8c85b80429a3d9a23a4425f49
        End
    End
    Begin
        dbText "Name" ="[Basic Product Construction].Base"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].Reason"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].Active"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =199
    Top =141
    Right =1439
    Bottom =831
    Left =-1
    Top =-1
    Right =1208
    Bottom =128
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =321
        Bottom =156
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
End

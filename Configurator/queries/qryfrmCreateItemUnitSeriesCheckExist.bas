Operation =1
Option =18
RowCount ="1"
Begin InputTables
    Name ="Unit Series"
    Name ="qryfrmCreateItemUnitSeriesCheck"
End
Begin OutputColumns
    Alias ="Existing"
    Expression ="IIf([UnitSeries] Is Null,0,-1)"
End
Begin Joins
    LeftTable ="Unit Series"
    RightTable ="qryfrmCreateItemUnitSeriesCheck"
    Expression ="[Unit Series].[Unit Color Series] = qryfrmCreateItemUnitSeriesCheck.UnitSeries"
    Flag =2
End
Begin OrderBy
    Expression ="IIf([UnitSeries] Is Null,0,-1)"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x4ecf6e82f9e54f49b4b288a833d46bcc
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Existing"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7a9162d13be1c04a839c864787ccbbb8
        End
    End
End
Begin
    State =0
    Left =195
    Top =306
    Right =1397
    Bottom =630
    Left =-1
    Top =-1
    Right =1170
    Bottom =93
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =172
        Top =6
        Right =360
        Bottom =124
        Top =0
        Name ="Unit Series"
        Name =""
    End
    Begin
        Left =821
        Top =25
        Right =917
        Bottom =98
        Top =0
        Name ="qryfrmCreateItemUnitSeriesCheck"
        Name =""
    End
End

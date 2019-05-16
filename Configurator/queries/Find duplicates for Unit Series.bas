Operation =1
Option =0
Where ="((([Unit Series].Number) In (SELECT [Number] FROM [Unit Series] As Tmp GROUP BY "
    "[Number] HAVING Count(*)>1 )))"
Begin InputTables
    Name ="Unit Series"
End
Begin OutputColumns
    Expression ="[Unit Series].Number"
    Expression ="[Unit Series].[Unit Color Series]"
    Expression ="[Unit Series].[Unit ID]"
    Expression ="[Unit Series].Description"
    Expression ="[Unit Series].[TB Color Series]"
    Expression ="[Unit Series].[Unit Type]"
    Expression ="[Unit Series].[Print type]"
    Expression ="[Unit Series].[Fiber Code]"
    Expression ="[Unit Series].[Print Spacing]"
    Expression ="[Unit Series].verified"
    Expression ="[Unit Series].Timestamp"
End
Begin OrderBy
    Expression ="[Unit Series].Number"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xc52aa3c6035f06428f16b7d73b8cf3c6
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
End
Begin
    State =0
    Left =86
    Top =142
    Right =1598
    Bottom =950
    Left =-1
    Top =-1
    Right =1480
    Bottom =220
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Unit Series"
        Name =""
    End
End

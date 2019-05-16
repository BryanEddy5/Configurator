Operation =1
Option =8
Where ="((([Basic Product Construction].Active)=True))"
Begin InputTables
    Name ="Basic Product Construction"
End
Begin OutputColumns
    Expression ="[Basic Product Construction].[Item No]"
End
Begin OrderBy
    Expression ="[Basic Product Construction].[Item No]"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x2961897ce9bd2f4f936495ac78f72dfd
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
End
Begin
    State =0
    Left =104
    Top =202
    Right =1898
    Bottom =789
    Left =-1
    Top =-1
    Right =1762
    Bottom =227
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =350
        Bottom =228
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
End

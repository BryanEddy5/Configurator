Operation =4
Option =0
Where ="((([Basic Product Construction].[New Oracle Part #])=[tblNewProductConstructionH"
    "old].[New Oracle Part #]))"
Begin InputTables
    Name ="Basic Product Construction"
    Name ="tblNewProductConstructionHold"
End
Begin OutputColumns
    Name ="[Basic Product Construction].Active"
    Expression ="-1"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x0a0f19fdc3541449a99e9595c45464af
End
Begin
End
Begin
    State =0
    Left =104
    Top =72
    Right =1898
    Bottom =636
    Left =-1
    Top =-1
    Right =1762
    Bottom =175
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =395
        Top =34
        Right =491
        Bottom =152
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
    Begin
        Left =801
        Top =6
        Right =1093
        Bottom =176
        Top =0
        Name ="tblNewProductConstructionHold"
        Name =""
    End
End

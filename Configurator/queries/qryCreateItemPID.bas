Operation =1
Option =2
Where ="((([Basic Product Construction].[Customer Part#]) Is Not Null And ([Basic Produc"
    "t Construction].[Customer Part#])<>\"\") And (([Basic Product Construction].Cust"
    "omer)=Forms!frmCreateItem!Customer))"
Begin InputTables
    Name ="Basic Product Construction"
End
Begin OutputColumns
    Expression ="[Basic Product Construction].[Customer Part#]"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x860b0b8a47e5684ab0de49d492048d84
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
End
Begin
    State =0
    Left =0
    Top =40
    Right =1579
    Bottom =881
    Left =-1
    Top =-1
    Right =1547
    Bottom =208
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =259
        Bottom =214
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
End

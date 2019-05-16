Operation =1
Option =0
Where ="(((tblItemInterfaceDataTypes.ItemType)=Forms!frmMainMenu!NewItemType) And ((tblI"
    "temInterfaceDataTypes.DataTypeName)=\"Locator\"))"
Begin InputTables
    Name ="tblItemInterfaceDataTypes"
End
Begin OutputColumns
    Expression ="tblItemInterfaceDataTypes.ItemType"
    Expression ="tblItemInterfaceDataTypes.DataTypeName"
    Alias ="LocationName"
    Expression ="tblItemInterfaceDataTypes.DataTypeValue"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbBinary "GUID" = Begin
    0xf43b20e6896b534daf2b12976c44ac8e
End
Begin
    Begin
        dbText "Name" ="LocationName"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2ce95b83557cde4687c657d9e484dc08
        End
    End
    Begin
        dbText "Name" ="tblItemInterfaceDataTypes.ItemType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblItemInterfaceDataTypes.DataTypeName"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =43
    Top =91
    Right =1497
    Bottom =1000
    Left =-1
    Top =-1
    Right =1422
    Bottom =93
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =259
        Bottom =120
        Top =0
        Name ="tblItemInterfaceDataTypes"
        Name =""
    End
End

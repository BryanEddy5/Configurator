Operation =1
Option =0
Where ="(((tblItemInterfaceDataTypes.ItemType)=Forms!frmMainMenu!NewItemType) And ((tblI"
    "temInterfaceDataTypes.DataTypeName)=\"Subinventory\") And ((tblItemInterfaceData"
    "Types.AppliesTo)=\"Item\"))"
Begin InputTables
    Name ="tblItemInterfaceDataTypes"
End
Begin OutputColumns
    Expression ="tblItemInterfaceDataTypes.ItemType"
    Expression ="tblItemInterfaceDataTypes.DataTypeName"
    Alias ="Subinventory"
    Expression ="tblItemInterfaceDataTypes.DataTypeValue"
    Expression ="tblItemInterfaceDataTypes.AppliesTo"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x9ef1284abd8a774c9d23ed7c399ef691
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Subinventory"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4edb34662891214fa7adbabbbd3222b2
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
    Begin
        dbText "Name" ="tblItemInterfaceDataTypes.AppliesTo"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =57
    Top =482
    Right =1540
    Bottom =1038
    Left =-1
    Top =-1
    Right =1451
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

Operation =1
Option =0
Where ="((([Basic Product Construction].[New Oracle Part #]) Like \"PT*\" Or ([Basic Pro"
    "duct Construction].[New Oracle Part #]) Like \"RBN*\") AND (([Basic Product Cons"
    "truction].Active)<>0))"
Begin InputTables
    Name ="Basic Product Construction"
End
Begin OutputColumns
    Expression ="[Basic Product Construction].[TB Material]"
    Expression ="[Basic Product Construction].[TB Nominal OD]"
    Expression ="[Basic Product Construction].[EZ Strip]"
    Alias ="Fiber"
    Expression ="[Basic Product Construction].[Fiber Oracle item]"
    Expression ="[Basic Product Construction].RibbonHighCure"
    Alias ="ComponentFiberCount"
    Expression ="Mid([Item No],3,3)"
    Alias ="TBLetter"
    Expression ="Mid([Item No],12,1)"
    Expression ="[Basic Product Construction].[New Oracle Part #]"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x68e950b74624ad4f81967906ffc63651
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="[Basic Product Construction].[TB Nominal OD]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[EZ Strip]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Fiber"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc0a24643baa82d47807a8d853000c3e8
        End
    End
    Begin
        dbText "Name" ="Component"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1740"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[TB Material]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].RibbonHighCure"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ComponentFiberCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe0e5531dfc967f49b01dbb7a08549282
        End
    End
    Begin
        dbText "Name" ="TBLetter"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc95efc42d528d545a0eba7c1e83f19ee
        End
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[New Oracle Part #]"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =58
    Top =239
    Right =1167
    Bottom =868
    Left =-1
    Top =-1
    Right =1077
    Bottom =193
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =268
        Top =19
        Right =713
        Bottom =330
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
End

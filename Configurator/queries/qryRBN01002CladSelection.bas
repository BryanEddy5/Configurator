Operation =1
Option =0
Where ="(((IIf([Clad Diameter] Is Null,0,CDbl([Clad Diameter])))<=124.9) And ((tblFiberE"
    "xtract.[Item Number])=\"fbr00388\"))"
Begin InputTables
    Name ="tblFiberExtract"
End
Begin OutputColumns
    Alias ="Expr1"
    Expression ="tblFiberExtract.[Box Number]"
    Alias ="Expr2"
    Expression ="tblFiberExtract.Locator"
    Alias ="Expr3"
    Expression ="tblFiberExtract.Name"
    Alias ="Expr4"
    Expression ="tblFiberExtract.Family"
    Alias ="Expr5"
    Expression ="tblFiberExtract.Manufacturer"
    Alias ="Expr6"
    Expression ="tblFiberExtract.[Item Number]"
    Alias ="Expr7"
    Expression ="tblFiberExtract.[Lot Number]"
    Alias ="Expr8"
    Expression ="tblFiberExtract.[Onhand Qty]"
    Alias ="Expr9"
    Expression ="tblFiberExtract.[Original Length]"
    Alias ="Expr10"
    Expression ="tblFiberExtract.Subinventory"
    Alias ="Expr11"
    Expression ="tblFiberExtract.Status"
    Alias ="Expr1"
    Expression ="IIf([GEO_CLAD_AVG_B] Is Null,0,CDbl([GEO_CLAD_AVG_B]))"
    Alias ="CladDia"
    Expression ="IIf([Clad Diameter] Is Null,0,CDbl([Clad Diameter]))"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xad8b23e599d5dd49a44431c046f183a9
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblFiberExtract.Family"
        dbInteger "ColumnWidth" ="2160"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xcc13e8e2021a7448842bf982c9eef189
        End
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4cd42ddf1ca5b84086965ae210b9068f
        End
    End
    Begin
        dbText "Name" ="Expr3"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xede0f94edf19ad4d81f8203a34057ce1
        End
    End
    Begin
        dbText "Name" ="Expr4"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x8404f6a6a5761846a765095a4970011c
        End
    End
    Begin
        dbText "Name" ="Expr5"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xcda6df8c443e7742ab54b11d20f96827
        End
    End
    Begin
        dbText "Name" ="Expr6"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x0f2ef80316846d4ba49ca3b3efb71cc2
        End
    End
    Begin
        dbText "Name" ="Expr7"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x03df9ef25120e142a8b5e4274888936e
        End
    End
    Begin
        dbText "Name" ="Expr8"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x73910b173348364481a878f57ddc0988
        End
    End
    Begin
        dbText "Name" ="Expr9"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf44ac3430b374f48a8f57fce9f7ebc64
        End
    End
    Begin
        dbText "Name" ="Expr10"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb4af683f99c90845a66423882507258b
        End
    End
    Begin
        dbText "Name" ="Expr11"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xcd128b544b04d946a71da164a56a798a
        End
    End
    Begin
        dbText "Name" ="CladDia"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xaea8c963e1078a4faa5566eda7093927
        End
    End
End
Begin
    State =0
    Left =3
    Top =2
    Right =1112
    Bottom =480
    Left =-1
    Top =-1
    Right =1077
    Bottom =247
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =497
        Bottom =244
        Top =0
        Name ="tblFiberExtract"
        Name =""
    End
End

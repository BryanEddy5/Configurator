Operation =1
Option =0
Where ="((([Basic Product Construction].[New Oracle Part #])=[Forms]![frmMainMenu]![NewP"
    "rintExtension]))"
Begin InputTables
    Name ="tblCableConstructions"
    Name ="tblCableConstructionReferences"
    Name ="vColor_Chip_Jacket_Material"
    Name ="Basic Product Construction"
End
Begin OutputColumns
    Alias ="Routing1"
    Expression ="Left([StandardOperation],4)"
    Alias ="Routing2"
    Expression ="Mid([StandardOperation],6,5)"
    Expression ="tblCableConstructions.JacketMaterial"
    Expression ="[Basic Product Construction].[New Oracle Part #]"
    Expression ="[Basic Product Construction].[Item No]"
    Expression ="[Basic Product Construction].Customer"
    Expression ="[Basic Product Construction].Active"
    Expression ="[Basic Product Construction].[Fiber Oracle item]"
    Expression ="[Basic Product Construction].FiberType2"
    Expression ="[Basic Product Construction].FiberType3"
    Expression ="[Basic Product Construction].[Customer Part#]"
    Expression ="[Basic Product Construction].[Jacket Color]"
    Expression ="[Basic Product Construction].Base"
    Expression ="[Basic Product Construction].[Unit Series]"
    Expression ="[Basic Product Construction].[Listing Company]"
    Expression ="[Basic Product Construction].CustomerRev"
    Expression ="[Basic Product Construction].Enumber"
    Expression ="vColor_Chip_Jacket_Material.ColorChip"
End
Begin Joins
    LeftTable ="tblCableConstructionReferences"
    RightTable ="tblCableConstructions"
    Expression ="tblCableConstructionReferences.BaseID = tblCableConstructions.BaseID"
    Flag =1
    LeftTable ="vColor_Chip_Jacket_Material"
    RightTable ="tblCableConstructions"
    Expression ="vColor_Chip_Jacket_Material.JacketMaterial = tblCableConstructions.JacketMateria"
        "l"
    Flag =1
    LeftTable ="Basic Product Construction"
    RightTable ="tblCableConstructionReferences"
    Expression ="[Basic Product Construction].Base = tblCableConstructionReferences.Base"
    Flag =1
    LeftTable ="Basic Product Construction"
    RightTable ="vColor_Chip_Jacket_Material"
    Expression ="[Basic Product Construction].[Jacket Color] = vColor_Chip_Jacket_Material.Color"
    Flag =1
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
    0x86b898f974503c4d95a747f954b1a816
End
Begin
    Begin
        dbText "Name" ="Routing1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6d0c9355cfd7ea48a0b24382e1b01bfb
        End
    End
    Begin
        dbText "Name" ="Routing2"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7111f52f2fe7d64fba18891a536415f5
        End
    End
    Begin
        dbText "Name" ="tblCableConstructions.JacketMaterial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vColor_Chip_Jacket_Material.ColorChip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].CustomerRev"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Fiber Oracle item]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Listing Company]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].FiberType2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].Base"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].Enumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].Customer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].FiberType3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Unit Series]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[New Oracle Part #]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Customer Part#]"
        dbInteger "ColumnWidth" ="2400"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].Active"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Item No]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Jacket Color]"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =73
    Top =100
    Right =1182
    Bottom =729
    Left =-1
    Top =-1
    Right =1077
    Bottom =161
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =824
        Top =10
        Right =1108
        Bottom =177
        Top =0
        Name ="tblCableConstructions"
        Name =""
    End
    Begin
        Left =479
        Top =10
        Right =694
        Bottom =154
        Top =0
        Name ="tblCableConstructionReferences"
        Name =""
    End
    Begin
        Left =1314
        Top =9
        Right =1458
        Bottom =153
        Top =0
        Name ="vColor_Chip_Jacket_Material"
        Name =""
    End
    Begin
        Left =1551
        Top =28
        Right =1695
        Bottom =172
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
End

Operation =1
Option =0
Where ="(((qryPremiseItemDetails_Union.OracleItemNumber)=[Forms]![frmMainMenu]![NewBaseI"
    "tem]))"
Begin InputTables
    Name ="qryPremiseItemDetails_Union"
    Name ="tblCableConstructions"
    Name ="tblCableConstructionReferences"
    Name ="vColor_Chip_Jacket_Material"
End
Begin OutputColumns
    Alias ="Routing1"
    Expression ="Left([StandardOperation],4)"
    Alias ="Routing2"
    Expression ="Mid([StandardOperation],6,5)"
    Expression ="tblCableConstructions.JacketMaterial"
    Expression ="qryPremiseItemDetails_Union.[Unit Series]"
    Expression ="qryPremiseItemDetails_Union.Enumber"
    Expression ="qryPremiseItemDetails_Union.Listing"
    Expression ="vColor_Chip_Jacket_Material.ColorChip"
    Expression ="qryPremiseItemDetails_Union.OracleItemNumber"
    Expression ="qryPremiseItemDetails_Union.DesignCode"
    Expression ="qryPremiseItemDetails_Union.Customer"
    Expression ="qryPremiseItemDetails_Union.Active"
    Expression ="qryPremiseItemDetails_Union.FiberType1"
    Expression ="qryPremiseItemDetails_Union.FiberType2"
    Expression ="qryPremiseItemDetails_Union.FiberType3"
    Expression ="qryPremiseItemDetails_Union.CustomerPartNumber"
    Expression ="qryPremiseItemDetails_Union.JacketColor"
    Expression ="qryPremiseItemDetails_Union.Base"
    Expression ="qryPremiseItemDetails_Union.UnitSeries"
    Expression ="qryPremiseItemDetails_Union.Listing"
    Expression ="qryPremiseItemDetails_Union.CustomerRev"
    Expression ="qryPremiseItemDetails_Union.Enumber"
    Expression ="qryPremiseItemDetails_Union.[Unit Series]"
    Expression ="vColor_Chip_Jacket_Material.Color_Chip_Concentration"
    Expression ="tblCableConstructions.CablePasses"
End
Begin Joins
    LeftTable ="qryPremiseItemDetails_Union"
    RightTable ="tblCableConstructionReferences"
    Expression ="qryPremiseItemDetails_Union.Base = tblCableConstructionReferences.Base"
    Flag =1
    LeftTable ="tblCableConstructionReferences"
    RightTable ="tblCableConstructions"
    Expression ="tblCableConstructionReferences.BaseID = tblCableConstructions.BaseID"
    Flag =1
    LeftTable ="tblCableConstructions"
    RightTable ="vColor_Chip_Jacket_Material"
    Expression ="tblCableConstructions.JacketMaterial = vColor_Chip_Jacket_Material.JacketMateria"
        "l"
    Flag =1
    LeftTable ="qryPremiseItemDetails_Union"
    RightTable ="vColor_Chip_Jacket_Material"
    Expression ="qryPremiseItemDetails_Union.JacketColor = vColor_Chip_Jacket_Material.Color"
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
    0xd72a3b4e1338034db93c3a71e03d3b8d
End
Begin
    Begin
        dbText "Name" ="Routing1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf68575d00c60464d8aa1cf6b4ea75878
        End
    End
    Begin
        dbText "Name" ="Routing2"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6cfde1f3de49b44e8ca97ea24c5fadba
        End
    End
    Begin
        dbText "Name" ="qryPremiseItemDetails_Union.OracleItemNumber"
        dbInteger "ColumnWidth" ="5820"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vColor_Chip_Jacket_Material.ColorChip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPremiseItemDetails_Union.[Unit Series]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1003"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPremiseItemDetails_Union.CustomerRev"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPremiseItemDetails_Union.Enumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.JacketMaterial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1004"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPremiseItemDetails_Union.DesignCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPremiseItemDetails_Union.Base"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1005"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPremiseItemDetails_Union.Customer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPremiseItemDetails_Union.JacketColor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPremiseItemDetails_Union.Active"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPremiseItemDetails_Union.UnitSeries"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPremiseItemDetails_Union.FiberType1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPremiseItemDetails_Union.FiberType2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPremiseItemDetails_Union.FiberType3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPremiseItemDetails_Union.CustomerPartNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryPremiseItemDetails_Union.Listing"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vColor_Chip_Jacket_Material.Color_Chip_Concentration"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.CablePasses"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =16
    Top =144
    Right =1125
    Bottom =773
    Left =-1
    Top =-1
    Right =1077
    Bottom =372
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =7
        Right =323
        Bottom =447
        Top =0
        Name ="qryPremiseItemDetails_Union"
        Name =""
    End
    Begin
        Left =351
        Top =22
        Right =542
        Bottom =425
        Top =0
        Name ="tblCableConstructions"
        Name =""
    End
    Begin
        Left =662
        Top =17
        Right =895
        Bottom =435
        Top =0
        Name ="tblCableConstructionReferences"
        Name =""
    End
    Begin
        Left =1078
        Top =49
        Right =1399
        Bottom =303
        Top =0
        Name ="vColor_Chip_Jacket_Material"
        Name =""
    End
End

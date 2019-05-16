Operation =1
Option =0
Where ="(((qryPremiseItemDetails_Union.OracleItemNumber)=[Forms]![frmMainMenu]![NewPrint"
    "Extension]))"
Begin InputTables
    Name ="qryPremiseItemDetails_Union"
    Name ="tblCableConstructions"
    Name ="tblCableConstructionReferences"
End
Begin OutputColumns
    Alias ="Routing1"
    Expression ="Left([StandardOperation],4)"
    Alias ="Routing2"
    Expression ="Mid([StandardOperation],7,5)"
    Expression ="qryPremiseItemDetails_Union.*"
    Expression ="qryPremiseItemDetails_Union.OracleItemNumber"
    Expression ="tblCableConstructions.JacketMaterial"
End
Begin Joins
    LeftTable ="qryPremiseItemDetails_Union"
    RightTable ="tblCableConstructionReferences"
    Expression ="qryPremiseItemDetails_Union.Base = tblCableConstructionReferences.Base"
    Flag =1
    LeftTable ="tblCableConstructions"
    RightTable ="tblCableConstructionReferences"
    Expression ="tblCableConstructions.BaseID = tblCableConstructionReferences.BaseID"
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
    0x2243bc4978e13942a469f9560fa93092
End
Begin
    Begin
        dbText "Name" ="Routing1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe3161e0a532c394fb263c9e4ba3777ff
        End
    End
    Begin
        dbText "Name" ="Routing2"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2ce78a1012992a49b44eab27416442c8
        End
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1109
    Bottom =669
    Left =-1
    Top =-1
    Right =1077
    Bottom =513
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="qryPremiseItemDetails_Union"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="tblCableConstructions"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =156
        Top =0
        Name ="tblCableConstructionReferences"
        Name =""
    End
End

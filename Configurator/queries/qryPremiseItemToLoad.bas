Operation =1
Option =0
Where ="(((qryPremiseItems.OracleItemNumber)=[Forms]![frmMainMenu]![NewItemNumber]))"
Begin InputTables
    Name ="qryPremiseItems"
    Name ="tblNetworkBOMBillOfMtlsInterface"
End
Begin OutputColumns
    Expression ="qryPremiseItems.*"
    Alias ="ExistingItem"
    Expression ="IIf([ItemNumber] Is Not Null,-1,0)"
End
Begin Joins
    LeftTable ="qryPremiseItems"
    RightTable ="tblNetworkBOMBillOfMtlsInterface"
    Expression ="qryPremiseItems.OracleItemNumber = tblNetworkBOMBillOfMtlsInterface.ItemNumber"
    Flag =2
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x5f3810a4c625c74cacb3c5b0e1188062
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="ExistingItem"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x8652ff56d409d5439ab9c20418d7c15a
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
    Bottom =110
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =245
        Bottom =109
        Top =0
        Name ="qryPremiseItems"
        Name =""
    End
    Begin
        Left =448
        Top =6
        Right =544
        Bottom =109
        Top =0
        Name ="tblNetworkBOMBillOfMtlsInterface"
        Name =""
    End
End

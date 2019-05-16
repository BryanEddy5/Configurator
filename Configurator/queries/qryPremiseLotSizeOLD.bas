Operation =1
Option =0
Where ="(((tblPremiseLotSizes.ItemPrefix)=Left([Forms]![frmMainMenu]![NewItemNumber],2))"
    ")"
Begin InputTables
    Name ="tblPremiseLotSizes"
End
Begin OutputColumns
    Expression ="tblPremiseLotSizes.ItemPrefix"
    Expression ="tblPremiseLotSizes.LotSize"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x819d598f7ebab14ab2d2415acae85a08
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
        Right =134
        Bottom =90
        Top =0
        Name ="tblPremiseLotSizes"
        Name =""
    End
End

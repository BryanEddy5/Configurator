Operation =1
Option =0
Where ="((([tblOracleBOMInventoryCompsInterface].[AssemblyItemNumber])=[Forms]![frmMainM"
    "enu]![NewItemNumber]))"
Begin InputTables
    Name ="tblOracleBOMInventoryCompsInterface"
End
Begin OutputColumns
    Alias ="ExistingBOMCheck"
    Expression ="Count(tblOracleBOMInventoryCompsInterface.AssemblyItemNumber)"
End
Begin Parameters
    Name ="[Forms]![frmMainMenu]![NewItemNumber]"
    Flag =10
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xd18d43c587e02b408535c085c05e20bd
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="ExistingBOMCheck"
        dbInteger "ColumnWidth" ="2910"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xda1ae8c7a716f14582f43dc2743083ec
        End
    End
End
Begin
    State =0
    Left =47
    Top =135
    Right =1389
    Bottom =455
    Left =-1
    Top =-1
    Right =1310
    Bottom =110
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =328
        Bottom =120
        Top =0
        Name ="tblOracleBOMInventoryCompsInterface"
        Name =""
    End
End

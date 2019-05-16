Operation =1
Option =0
Where ="((([tblOracleBOMOpSequencesInterface].[AssemblyItemNumber])=[Forms]![frmMainMenu"
    "]![NewItemNumber]))"
Begin InputTables
    Name ="tblOracleBOMOpSequencesInterface"
End
Begin OutputColumns
    Alias ="ExistingRoutingCheck"
    Expression ="Count(tblOracleBOMOpSequencesInterface.AssemblyItemNumber)"
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
    0x2386f930dc4c274fb51ddce48b4e64e3
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="ExistingRoutingCheck"
        dbInteger "ColumnWidth" ="2910"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xba3a3b2fa2a2ac4ca273a827eda5c5c9
        End
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1560
    Bottom =854
    Left =-1
    Top =-1
    Right =1528
    Bottom =110
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =285
        Bottom =120
        Top =0
        Name ="tblOracleBOMOpSequencesInterface"
        Name =""
    End
End

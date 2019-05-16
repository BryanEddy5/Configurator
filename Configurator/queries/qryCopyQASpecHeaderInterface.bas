Operation =3
Name ="tblQAItemSpecs"
Option =2
Begin InputTables
    Name ="tblItemInterface"
End
Begin OutputColumns
    Name ="SpecName"
    Expression ="tblItemInterface.ItemNumber"
    Name ="Item"
    Expression ="tblItemInterface.ItemNumber"
End
Begin Parameters
    Name ="[Forms]![frmMainMenu]![NewItemNumber]"
    Flag =10
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBinary "GUID" = Begin
    0x2209caf8b098514ca896fe97560aa957
End
Begin
    Begin
        dbText "Name" ="OrganizationCode"
        dbInteger "ColumnWidth" ="1725"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SpecName"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd43ca7323ed57d47b9ed6369468e747a
        End
    End
    Begin
        dbText "Name" ="Item"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x71e749956fe97542a77ccf4d8ea662a9
        End
    End
    Begin
        dbText "Name" ="tblItemInterface.ItemNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AFLPRD_INVSysItemSpec_CAB.ItemNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblItemInterface.Description"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =254
    Top =90
    Right =1151
    Bottom =667
    Left =-1
    Top =-1
    Right =865
    Bottom =278
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =389
        Top =52
        Right =688
        Bottom =245
        Top =0
        Name ="tblItemInterface"
        Name =""
    End
End

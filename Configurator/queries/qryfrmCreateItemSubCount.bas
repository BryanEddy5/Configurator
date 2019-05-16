Operation =3
Name ="tblfrmCreateItemSubCount"
Option =0
Where ="(((tblCableConstructionReferences.Base)=[forms]![frmCreateItem]![Base]))"
Begin InputTables
    Name ="tblCableConstructionReferences"
    Name ="Product Desrcriptions"
End
Begin OutputColumns
    Alias ="NumberSubs"
    Name ="NumSubs"
    Expression ="Sum([Product Desrcriptions].[NumSubs])"
    Alias ="Fillers"
    Name ="NumSubFillers"
    Expression ="Avg([NumSubFillers])"
End
Begin Joins
    LeftTable ="Product Desrcriptions"
    RightTable ="tblCableConstructionReferences"
    Expression ="[Product Desrcriptions].ProductID = tblCableConstructionReferences.SetupID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xacc488ad683dd44d98fe7c20f2eb91c2
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbBoolean "UseTransaction" ="-1"
Begin
    Begin
        dbText "Name" ="[Product Desrcriptions].NumSubs"
        dbInteger "ColumnWidth" ="1305"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.Base"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.NumSubFillers"
        dbInteger "ColumnWidth" ="2130"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Product Desrcriptions].[Level 1 Product]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbInteger "ColumnWidth" ="1305"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NumSubs"
        dbInteger "ColumnWidth" ="1305"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr2"
        dbInteger "ColumnWidth" ="2130"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NumberSubs"
        dbInteger "ColumnWidth" ="1305"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x00000000000000000000000001000000
        End
    End
    Begin
        dbText "Name" ="Fillers"
        dbInteger "ColumnWidth" ="2130"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x00000000000000000000000001000000
        End
    End
End
Begin
    State =0
    Left =-20
    Top =55
    Right =1593
    Bottom =763
    Left =-1
    Top =-1
    Right =1581
    Bottom =277
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =89
        Top =16
        Right =581
        Bottom =245
        Top =0
        Name ="tblCableConstructionReferences"
        Name =""
    End
    Begin
        Left =747
        Top =18
        Right =1318
        Bottom =292
        Top =0
        Name ="Product Desrcriptions"
        Name =""
    End
End

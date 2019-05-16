Operation =1
Option =1
Begin InputTables
    Name ="tblJacketMaterialColorChips"
End
Begin OutputColumns
    Expression ="tblJacketMaterialColorChips.*"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x0fdc5e0703ff164585544082612ef20d
End
dbMemo "Filter" ="((ColorLetter=\"6\"))"
dbMemo "OrderBy" ="Jacket"
dbByte "RecordsetType" ="0"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblJacketMaterialColorChips.Color"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblJacketMaterialColorChips.Jacket"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblJacketMaterialColorChips.ColorLetter"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblJacketMaterialColorChips.ColorChip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblJacketMaterialColorChips.TimeStamp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblJacketMaterialColorChips.DateCreated"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-187
    Top =170
    Right =1105
    Bottom =654
    Left =-1
    Top =-1
    Right =1260
    Bottom =175
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =110
        Top =14
        Right =348
        Bottom =158
        Top =0
        Name ="tblJacketMaterialColorChips"
        Name =""
    End
End

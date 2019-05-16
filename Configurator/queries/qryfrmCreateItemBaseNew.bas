Operation =1
Option =0
Where ="(((vCreateItemBase.Active)<>0) And ((vCreateItemBase.DesignCode) Like Left(Forms"
    "!frmMainMenu!ItemNo,5) & Mid(Forms!frmMainMenu!ItemNo,7,2)) And ((vCreateItemBas"
    "e.ReleasedDesign)<>0) And ((vCreateItemBase.TBLetter)=Mid(Forms!frmCreateItem!It"
    "emNo,12,1)))"
Begin InputTables
    Name ="vCreateItemBase"
End
Begin OutputColumns
    Expression ="vCreateItemBase.Base"
    Expression ="vCreateItemBase.ConstructionDescription"
    Expression ="vCreateItemBase.OD"
    Expression ="vCreateItemBase.NumPositions"
    Expression ="vCreateItemBase.NumFillers"
    Expression ="vCreateItemBase.JacketMaterial"
    Expression ="vCreateItemBase.CountOfBase"
    Expression ="vCreateItemBase.TightBuffered"
    Expression ="vCreateItemBase.Active"
    Expression ="vCreateItemBase.DesignCode"
    Expression ="vCreateItemBase.Location"
    Expression ="vCreateItemBase.ReleasedDesign"
    Expression ="vCreateItemBase.TBLetter"
    Expression ="vCreateItemBase.TBTypeID"
    Expression ="vCreateItemBase.DesignExtension"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x650a70e827721745bda01287e6f2000c
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="vCreateItemBase.OD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vCreateItemBase.Active"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vCreateItemBase.TBLetter"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vCreateItemBase.ConstructionDescription"
        dbInteger "ColumnWidth" ="2670"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vCreateItemBase.JacketMaterial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vCreateItemBase.ReleasedDesign"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vCreateItemBase.NumPositions"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vCreateItemBase.DesignCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vCreateItemBase.Base"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vCreateItemBase.NumFillers"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vCreateItemBase.Location"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vCreateItemBase.TBTypeID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vCreateItemBase.TightBuffered"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vCreateItemBase.DesignExtension"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1551
    Bottom =957
    Left =-1
    Top =-1
    Right =1519
    Bottom =298
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =533
        Top =1
        Right =809
        Bottom =320
        Top =0
        Name ="vCreateItemBase"
        Name =""
    End
End

Operation =1
Option =0
Where ="(((vCreateItemBaseFurcation.Active)<>0) And ((vCreateItemBaseFurcation.DesignCod"
    "e) Like Left(Forms!frmMainMenu!ItemNo,5) & Mid(Forms!frmMainMenu!ItemNo,7,2)) An"
    "d ((vCreateItemBaseFurcation.TBLetter)=Mid(Forms!frmCreateItem!ItemNo,12,1)) And"
    " ((vCreateItemBaseFurcation.FurcationLetter)=Mid(Forms!frmCreateItem!ItemNo,11,1"
    ")) And ((vCreateItemBaseFurcation.ReleasedDesign)<>0))"
Begin InputTables
    Name ="vCreateItemBaseFurcation"
End
Begin OutputColumns
    Expression ="vCreateItemBaseFurcation.Base"
    Expression ="vCreateItemBaseFurcation.ConstructionDescription"
    Expression ="vCreateItemBaseFurcation.OD"
    Expression ="vCreateItemBaseFurcation.NumPositions"
    Expression ="vCreateItemBaseFurcation.NumFillers"
    Expression ="vCreateItemBaseFurcation.JacketMaterial"
    Expression ="vCreateItemBaseFurcation.CountOfBase"
    Expression ="vCreateItemBaseFurcation.TightBuffered"
    Expression ="vCreateItemBaseFurcation.Active"
    Expression ="vCreateItemBaseFurcation.DesignCode"
    Expression ="vCreateItemBaseFurcation.Location"
    Expression ="vCreateItemBaseFurcation.ReleasedDesign"
    Expression ="vCreateItemBaseFurcation.TBLetter"
    Expression ="vCreateItemBaseFurcation.FurcationLetter"
    Expression ="vCreateItemBaseFurcation.TBTypeID"
    Expression ="vCreateItemBaseFurcation.ReleasedDesign"
    Expression ="vCreateItemBaseFurcation.DesignExtension"
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
    0xb912bfd4b0e1ac4c889ac8199109f18a
End
Begin
    Begin
        dbText "Name" ="vCreateItemBaseFurcation.NumPositions"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vCreateItemBaseFurcation.Active"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vCreateItemBaseFurcation.Base"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vCreateItemBaseFurcation.NumFillers"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vCreateItemBaseFurcation.DesignCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vCreateItemBaseFurcation.FurcationLetter"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vCreateItemBaseFurcation.ConstructionDescription"
        dbInteger "ColumnWidth" ="2670"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vCreateItemBaseFurcation.JacketMaterial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vCreateItemBaseFurcation.Location"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vCreateItemBaseFurcation.OD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vCreateItemBaseFurcation.CountOfBase"
        dbInteger "ColumnWidth" ="1635"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vCreateItemBaseFurcation.ReleasedDesign"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vCreateItemBaseFurcation.TBTypeID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vCreateItemBaseFurcation.TightBuffered"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vCreateItemBaseFurcation.DesignExtension"
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
    Bottom =245
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =150
        Top =22
        Right =370
        Bottom =197
        Top =0
        Name ="vCreateItemBaseFurcation"
        Name =""
    End
End

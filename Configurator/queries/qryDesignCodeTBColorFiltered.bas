Operation =1
Option =0
Where ="(((tblDesignCodeJacketColor.CablePasses)=Forms!frmDesignCode!Location) And ((tbl"
    "DesignCodeJacketColor.SubunitCableColorPostion)=True))"
Begin InputTables
    Name ="tblDesignCodeJacketColor"
End
Begin OutputColumns
    Expression ="tblDesignCodeJacketColor.ColorLetter"
    Expression ="tblDesignCodeJacketColor.Color"
    Expression ="tblDesignCodeJacketColor.CablePasses"
    Expression ="tblDesignCodeJacketColor.SubunitCableColorPostion"
End
Begin OrderBy
    Expression ="tblDesignCodeJacketColor.ColorLetter"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x6c5d8fe51831a74991914b5bc2372675
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblDesignCodeJacketColor.ColorLetter"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblDesignCodeJacketColor.Color"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblDesignCodeJacketColor.CablePasses"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblDesignCodeJacketColor.SubunitCableColorPostion"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-444
    Top =254
    Right =1377
    Bottom =1130
    Left =-1
    Top =-1
    Right =1789
    Bottom =121
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =81
        Top =-3
        Right =342
        Bottom =172
        Top =0
        Name ="tblDesignCodeJacketColor"
        Name =""
    End
End

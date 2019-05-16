Operation =1
Option =0
Where ="(((tblDesignCodeJacketColor.SubunitCableColorPostion)=True))"
Begin InputTables
    Name ="tblDesignCodeJacketColor"
End
Begin OutputColumns
    Expression ="tblDesignCodeJacketColor.ColorLetter"
    Expression ="tblDesignCodeJacketColor.Color"
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
    0x431170b4112a6e4f939925c3cc86bc28
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
        dbText "Name" ="tblDesignCodeJacketColor.SubunitCableColorPostion"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-252
    Top =167
    Right =1456
    Bottom =491
    Left =-1
    Top =-1
    Right =1676
    Bottom =76
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =324
        Bottom =128
        Top =0
        Name ="tblDesignCodeJacketColor"
        Name =""
    End
End

Operation =1
Option =2
Having ="(((tblDesignCodeFiberElementValue.FiberLetter)=\"@\"))"
Begin InputTables
    Name ="qryFiberFamily"
    Name ="tblDesignCodeFiberElementValue"
End
Begin OutputColumns
    Expression ="tblDesignCodeFiberElementValue.FiberLetter"
    Expression ="tblDesignCodeFiberElementValue.FiberValue"
    Expression ="tblDesignCodeFiberElementValue.PreferedCableColor"
End
Begin Joins
    LeftTable ="tblDesignCodeFiberElementValue"
    RightTable ="qryFiberFamily"
    Expression ="tblDesignCodeFiberElementValue.FiberLetter = qryFiberFamily.FiberLetter"
    Flag =2
End
Begin OrderBy
    Expression ="tblDesignCodeFiberElementValue.FiberValue"
    Flag =0
End
Begin Groups
    Expression ="tblDesignCodeFiberElementValue.FiberLetter"
    GroupLevel =0
    Expression ="tblDesignCodeFiberElementValue.FiberValue"
    GroupLevel =0
    Expression ="tblDesignCodeFiberElementValue.PreferedCableColor"
    GroupLevel =0
    Expression ="qryFiberFamily.FiberFamily"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x9769d47a749a924d9da504507330ad53
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblDesignCodeFiberElementValue.FiberLetter"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblDesignCodeFiberElementValue.FiberValue"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblDesignCodeFiberElementValue.PreferedCableColor"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-389
    Top =264
    Right =1153
    Bottom =648
    Left =-1
    Top =-1
    Right =1510
    Bottom =80
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =656
        Top =13
        Right =896
        Bottom =130
        Top =0
        Name ="qryFiberFamily"
        Name =""
    End
    Begin
        Left =227
        Top =-4
        Right =482
        Bottom =140
        Top =0
        Name ="tblDesignCodeFiberElementValue"
        Name =""
    End
End

Operation =1
Option =0
Begin InputTables
    Name ="tblVersionDesignSupport"
    Name ="tblVersionRevsDesignSupport"
End
Begin OutputColumns
    Expression ="tblVersionDesignSupport.*"
    Alias ="Expr1"
    Expression ="tblVersionRevsDesignSupport.Revision"
End
Begin OrderBy
    Expression ="tblVersionDesignSupport.RevDate"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x874be33a40124441bcb348775c95817b
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x671e9689c6b3e64e95065431024fe9c9
        End
    End
End
Begin
    State =0
    Left =18
    Top =40
    Right =1127
    Bottom =360
    Left =-1
    Top =-1
    Right =1077
    Bottom =110
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =45
        Top =0
        Name ="tblVersionDesignSupport"
        Name =""
    End
    Begin
        Left =309
        Top =4
        Right =405
        Bottom =43
        Top =0
        Name ="tblVersionRevsDesignSupport"
        Name =""
    End
End

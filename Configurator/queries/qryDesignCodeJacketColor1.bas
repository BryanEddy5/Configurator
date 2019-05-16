Operation =1
Option =0
Where ="(((tblDesignCodeJacketColor.ColorLetter)<>\"s\" And (tblDesignCodeJacketColor.Co"
    "lorLetter)<>\"y\" And (tblDesignCodeJacketColor.ColorLetter)<>\"x\" And (tblDesi"
    "gnCodeJacketColor.ColorLetter)<>\"z\"))"
Begin InputTables
    Name ="tblDesignCodeJacketColor"
End
Begin OutputColumns
    Alias ="Expr1"
    Expression ="tblDesignCodeJacketColor.ColorLetter"
    Alias ="Expr2"
    Expression ="tblDesignCodeJacketColor.Color"
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
    0x170d47fe88668b4fb2e3f0e57ef40a5c
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4632b530a30baa4fa8f413d9a6c87754
        End
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc2a18a7a02414242a91536b6b8e6702c
        End
    End
End
Begin
    State =0
    Left =260
    Top =185
    Right =1872
    Bottom =509
    Left =-1
    Top =-1
    Right =1580
    Bottom =93
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =94
        Top =0
        Name ="tblDesignCodeJacketColor"
        Name =""
    End
End

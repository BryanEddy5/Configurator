Operation =1
Option =0
Where ="(((tblDesignCodeJacketColor.ColorLetter)<>\"s\" And (tblDesignCodeJacketColor.Co"
    "lorLetter)<>\"y\" And (tblDesignCodeJacketColor.ColorLetter)<>\"x\" And (tblDesi"
    "gnCodeJacketColor.ColorLetter)<>\"z\") And ((tblDesignCodeJacketColor.CablePasse"
    "s)=Forms!frmDesignCode.location))"
Begin InputTables
    Name ="tblDesignCodeJacketColor"
End
Begin OutputColumns
    Expression ="tblDesignCodeJacketColor.ColorLetter"
    Expression ="tblDesignCodeJacketColor.Color"
    Expression ="tblDesignCodeJacketColor.CablePasses"
    Expression ="tblDesignCodeJacketColor.SubunitCableColorPostion"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xc24f47ffa658c347acf6ab8faaad3737
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
    Left =100
    Top =247
    Right =1765
    Bottom =819
    Left =-1
    Top =-1
    Right =1633
    Bottom =127
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =363
        Bottom =181
        Top =0
        Name ="tblDesignCodeJacketColor"
        Name =""
    End
End

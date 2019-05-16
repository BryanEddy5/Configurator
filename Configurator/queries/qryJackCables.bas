Operation =1
Option =2
Where ="((([Basic (Cable/Sheath)].Active)=-1) And (([Basic (Cable/Sheath)].[New Oracle P"
    "art Number]) Not Like \"*test*\") And (([Basic (Cable/Sheath)].[Item No]) Not Li"
    "ke \"*bif*\" And ([Basic (Cable/Sheath)].[Item No]) Not Like \"*crg*\"))"
Begin InputTables
    Name ="Basic (Cable/Sheath)"
End
Begin OutputColumns
    Alias ="Expr1"
    Expression ="[Basic (Cable/Sheath)].[Jacket Material]"
    Alias ="Jacket_Material_letter"
    Expression ="Mid([Item No],2,1)"
    Alias ="Jacket_Color_letter"
    Expression ="Mid([Item No],10,1)"
    Alias ="Expr2"
    Expression ="[Basic (Cable/Sheath)].Base"
    Alias ="AddOn"
    Expression ="Mid([Item No],14,10)"
    Alias ="Expr3"
    Expression ="[Basic (Cable/Sheath)].[Jacket Color]"
End
Begin OrderBy
    Expression ="Mid([Item No],2,1)"
    Flag =0
    Expression ="Mid([Item No],10,1)"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbMemo "OrderBy" ="[Jacket Material]"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xe2d4c30c02d1f64c9ecfb8da727a7aed
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Jacket_Material_letter"
        dbInteger "ColumnWidth" ="1560"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x04788d5b235deb4abb77498d1773408d
        End
    End
    Begin
        dbText "Name" ="Jacket_Color_letter"
        dbInteger "ColumnWidth" ="2160"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x78e6bd3cb4c89749a8e9cf84c6948ef8
        End
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9ca30a578c61ca44bac1f5919f3a5531
        End
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x023610eb6fb80347beadc8c5d8245c80
        End
    End
    Begin
        dbText "Name" ="AddOn"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd8ee56a5a3812a46942d7b3b4e1bf000
        End
    End
    Begin
        dbText "Name" ="Expr3"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd1f4f9dff8977248ad75759f42ce2d56
        End
    End
End
Begin
    State =0
    Left =18
    Top =14
    Right =1310
    Bottom =400
    Left =-1
    Top =-1
    Right =1260
    Bottom =155
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =172
        Top =6
        Right =366
        Bottom =169
        Top =0
        Name ="Basic (Cable/Sheath)"
        Name =""
    End
End

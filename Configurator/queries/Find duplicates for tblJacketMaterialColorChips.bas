Operation =1
Option =0
Where ="((([tblJacketMaterialColorChips].[Jacket]) In (SELECT [Jacket] FROM [tblJacketMa"
    "terialColorChips] As Tmp GROUP BY [Jacket],[ColorLetter] HAVING Count(*)>1  And "
    "[ColorLetter] = [tblJacketMaterialColorChips].[ColorLetter])))"
Begin InputTables
    Name ="tblJacketMaterialColorChips"
End
Begin OutputColumns
    Alias ="Expr1"
    Expression ="tblJacketMaterialColorChips.Jacket"
    Alias ="Expr2"
    Expression ="tblJacketMaterialColorChips.ColorLetter"
    Alias ="Expr3"
    Expression ="tblJacketMaterialColorChips.Color"
    Alias ="Expr4"
    Expression ="tblJacketMaterialColorChips.ColorChip"
    Alias ="Expr5"
    Expression ="tblJacketMaterialColorChips.TimeStamp"
    Alias ="Expr6"
    Expression ="tblJacketMaterialColorChips.DateCreated"
End
Begin OrderBy
    Expression ="tblJacketMaterialColorChips.Jacket"
    Flag =0
    Expression ="tblJacketMaterialColorChips.ColorLetter"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xfbcba8f9e7858248b97371cc19be6a98
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbMemo "OrderBy" ="[Find duplicates for tblJacketMaterialColorChips].[Color]"
Begin
    Begin
        dbText "Name" ="Jacket"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf252edd0ab620140a5ee59a88b753f83
        End
    End
    Begin
        dbText "Name" ="ColorLetter"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x109b3663137e384e808451b80b82235e
        End
    End
    Begin
        dbText "Name" ="Color"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4163712af114fa4f9d6efe5928e00137
        End
    End
    Begin
        dbText "Name" ="ColorChip"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x23cf4ac45d2ff444a8c28317bef1aa8e
        End
    End
    Begin
        dbText "Name" ="TimeStamp"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4a5132f9b620534294f0b90ad7893d4e
        End
    End
    Begin
        dbText "Name" ="DateCreated"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd0f136a98ad4214b96224172bc57079a
        End
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x864ea6583fe3c744a2f70417bc98b5c7
        End
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7bca95894d9dbf43b788f3ac1375dc19
        End
    End
    Begin
        dbText "Name" ="Expr3"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x5fe1d11148e2554ea737f35e6b825237
        End
    End
    Begin
        dbText "Name" ="Expr4"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa00b4f41700f834db19e65d0ff80606c
        End
    End
    Begin
        dbText "Name" ="Expr5"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xdf57e934bd0a1e4aacc7db03e8a3ab4e
        End
    End
    Begin
        dbText "Name" ="Expr6"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xcaec1f04f6e5d049a15d4933e6fb355c
        End
    End
End
Begin
    State =0
    Left =236
    Top =84
    Right =1565
    Bottom =785
    Left =-1
    Top =-1
    Right =1297
    Bottom =220
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="tblJacketMaterialColorChips"
        Name =""
    End
End

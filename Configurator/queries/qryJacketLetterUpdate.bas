Operation =4
Option =0
Where ="(((BasicCableSheath_LIVE.[Jacket Color]) Like \"*BLACK*\") And ((Mid([Item No],1"
    "0,2))=\"00\"))"
Begin InputTables
    Name ="BasicCableSheath_LIVE"
End
Begin OutputColumns
    Name ="BasicCableSheath_LIVE.Item No"
    Expression ="Left([Item No],9) & \"8\" & Mid([Item No],11,10)"
    Name ="BasicCableSheath_LIVE.Revision Date"
    Expression ="Date()"
    Name ="BasicCableSheath_LIVE.Revision Letter"
    Expression ="IIf([Revision Letter]='A','B',IIf([Revision Letter]='B','C',IIf([Revision Letter"
        "]='C','D',IIf([Revision Letter]='D','E',IIf([Revision Letter]='E','F',IIf([Revis"
        "ion Letter]='F','G',IIf([Revision Letter]='G','H',IIf([Revision Letter]='H','I',"
        "IIf([Revision Letter]='I','J',IIf([Revision Letter]='J','K',IIf([Revision Letter"
        "]='K','L',IIf([Revision Letter]='L','M',\"\"))))))))))))"
    Name ="BasicCableSheath_LIVE.Revision History"
    Expression ="Date() & \" - Updated Design Code Jacket Letter BCE;  \" & [Revision History]"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x1f4dd8b303cedb48b58e7116347ba60a
End
Begin
    Begin
        dbText "Name" ="BasicCableSheath_LIVE.[Revision History]"
        dbInteger "ColumnWidth" ="17715"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x01b254b055759f4682c1c6ca9351060f
        End
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x0a3646d8dc5a614e9c4ac540090dfa27
        End
    End
    Begin
        dbText "Name" ="Expr3"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf2bf1dc597a64849befe16daba22bfb7
        End
    End
    Begin
        dbText "Name" ="Expr4"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x770b0886c42c024f952b5ed32b7a63a0
        End
    End
End
Begin
    State =0
    Left =79
    Top =154
    Right =1371
    Bottom =758
    Left =-1
    Top =-1
    Right =1260
    Bottom =373
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =124
        Top =0
        Name ="BasicCableSheath_LIVE"
        Name =""
    End
End

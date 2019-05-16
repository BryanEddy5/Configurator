Operation =1
Option =0
Having ="(((qryBaseNomOD.Base) Not Like \"ia*\") And ((qryItemNoSearch.[Item No]) Like Fo"
    "rms!frmDesignCode!CableType & Forms!frmDesignCode!Jacket & Forms!frmDesignCode!F"
    "iberCount & \"*\") And ((qryItemNoSearch.Active)=True) And ((qryItemNoSearch.Loc"
    "ation) Like Forms!frmDesignCode!Passes & \"*\"))"
Begin InputTables
    Name ="qryBaseNomOD"
    Name ="qryItemNoSearch"
End
Begin OutputColumns
    Expression ="qryBaseNomOD.OD"
    Expression ="qryBaseNomOD.Base"
    Expression ="qryItemNoSearch.[Item No]"
    Alias ="Expr1"
    Expression ="qryItemNoSearch.Location"
    Expression ="qryBaseNomOD.Jacket"
    Alias ="TBLetter"
    Expression ="Mid([Item No],12,1)"
    Alias ="FiberCount"
    Expression ="Mid([Item No],3,3)"
    Alias ="Expr2"
    Expression ="qryItemNoSearch.Location"
End
Begin Joins
    LeftTable ="qryBaseNomOD"
    RightTable ="qryItemNoSearch"
    Expression ="qryBaseNomOD.Base = qryItemNoSearch.Base"
    Flag =1
End
Begin Groups
    Expression ="qryBaseNomOD.OD"
    GroupLevel =0
    Expression ="qryBaseNomOD.Base"
    GroupLevel =0
    Expression ="qryItemNoSearch.[Item No]"
    GroupLevel =0
    Expression ="qryBaseNomOD.Jacket"
    GroupLevel =0
    Expression ="Mid([Item No],12,1)"
    GroupLevel =0
    Expression ="Mid([Item No],3,3)"
    GroupLevel =0
    Expression ="qryItemNoSearch.Location"
    GroupLevel =0
    Expression ="qryItemNoSearch.Location"
    GroupLevel =0
    Expression ="qryItemNoSearch.Active"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x1597ce2ecd2299429925e437ddfbc6b9
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="TBLetter"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9da0226997bb8b4ebc85bcf2cff24720
        End
    End
    Begin
        dbText "Name" ="FiberCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x115f5b6682537f49a548df0b48d9638f
        End
    End
    Begin
        dbText "Name" ="qryBaseNomOD.OD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryBaseNomOD.Base"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryItemNoSearch.[Item No]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1003"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4f6f3ce3fadce64f81b396436c767da4
        End
    End
    Begin
        dbText "Name" ="qryBaseNomOD.Jacket"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryItemNoSearch.Location"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item No"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xaaea0680f0a0474c9c90cfe570c7f7bb
        End
    End
    Begin
        dbText "Name" ="Location"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf7606fcbdbc14d4195cdb818f7ea0f0f
        End
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xcbdf5b68fd60ff4c9cab1e8f4befd49f
        End
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x41d69d7c8b255f498c7667d6706c57ae
        End
    End
End
Begin
    State =0
    Left =-484
    Top =533
    Right =1128
    Bottom =987
    Left =-1
    Top =-1
    Right =1580
    Bottom =177
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =235
        Bottom =124
        Top =0
        Name ="qryBaseNomOD"
        Name =""
    End
    Begin
        Left =281
        Top =8
        Right =555
        Bottom =171
        Top =0
        Name ="qryItemNoSearch"
        Name =""
    End
End

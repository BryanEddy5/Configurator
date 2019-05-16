Operation =1
Option =2
Where ="(((qry_SearchByFiberType.Fiber) Not Like \"*MIXED*\") And ((qry_SearchByFiberTyp"
    "e.[TBF/Ribbon])=Mid(Forms!frmCreateItem!ItemNo,12,1)) And ((Mid(Forms!frmCreateI"
    "tem!ItemNo,13,30)) Like \"*\" & Mid([ITU],7,2) & \"*\") And ((Forms!frmCreateIte"
    "m!Fiber1Type)=qry_SearchByFiberType!Fiber_Type Or (Forms!frmCreateItem!Fiber1Typ"
    "e)=qry_SearchByFiberType!Alt1_Fiber_Type Or (Forms!frmCreateItem!Fiber1Type)=qry"
    "_SearchByFiberType!Alt2_Fiber_Type Or (Forms!frmCreateItem!Fiber1Type)=qry_Searc"
    "hByFiberType!Alt3_Fiber_Type Or (Forms!frmCreateItem!Fiber1Type)=qry_SearchByFib"
    "erType!Alt4_Fiber_Type))"
Begin InputTables
    Name ="qry_SearchByFiberType"
End
Begin OutputColumns
    Expression ="qry_SearchByFiberType.Fiber"
    Expression ="qry_SearchByFiberType.BIF"
    Expression ="qry_SearchByFiberType.Extension"
    Expression ="qry_SearchByFiberType.Vendor"
    Expression ="qry_SearchByFiberType.Cost"
    Expression ="qry_SearchByFiberType.SpecialNotes"
    Expression ="qry_SearchByFiberType.Description_Oracle"
    Expression ="qry_SearchByFiberType.[SM/MM]"
    Expression ="qry_SearchByFiberType.[OM#]"
    Alias ="Expr1"
    Expression ="Mid(Forms!frmCreateItem!ItemNo,13,30)"
    Expression ="qry_SearchByFiberType.Extension"
    Expression ="qry_SearchByFiberType.[TBF/Ribbon]"
End
Begin OrderBy
    Expression ="qry_SearchByFiberType.Cost"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x5d0cd6b73787bf44b38159db98ba115b
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="qry_SearchByFiberType.SpecialNotes"
        dbInteger "ColumnWidth" ="1320"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xfccebf86554c414da99a15c7f7a4ecda
        End
    End
    Begin
        dbText "Name" ="qry_SearchByFiberType.Fiber"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =75
    Top =237
    Right =1792
    Bottom =877
    Left =-1
    Top =-1
    Right =1685
    Bottom =283
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =331
        Bottom =244
        Top =0
        Name ="qry_SearchByFiberType"
        Name =""
    End
End

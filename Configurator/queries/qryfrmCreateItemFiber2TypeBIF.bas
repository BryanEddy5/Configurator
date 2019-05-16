Operation =1
Option =2
Where ="(((qry_SearchByFiberType.Fiber) Not Like \"*MIXED*\") And ((qry_SearchByFiberTyp"
    "e.BIF)=\"YES\") And ((qry_SearchByFiberType.[TBF/Ribbon])=Mid(Forms!frmCreateIte"
    "m!ItemNo,12,1)) And ((Forms!frmCreateItem!Fiber2Type)=qry_SearchByFiberType!Fibe"
    "r_Type Or (Forms!frmCreateItem!Fiber2Type)=qry_SearchByFiberType!Alt1_Fiber_Type"
    " Or (Forms!frmCreateItem!Fiber2Type)=qry_SearchByFiberType!Alt2_Fiber_Type Or (F"
    "orms!frmCreateItem!Fiber2Type)=qry_SearchByFiberType!Alt3_Fiber_Type Or (Forms!f"
    "rmCreateItem!Fiber2Type)=qry_SearchByFiberType!Alt4_Fiber_Type))"
Begin InputTables
    Name ="qry_SearchByFiberType"
End
Begin OutputColumns
    Expression ="qry_SearchByFiberType.Fiber"
    Expression ="qry_SearchByFiberType.BIF"
    Expression ="qry_SearchByFiberType.[OM#]"
    Expression ="qry_SearchByFiberType.Vendor"
    Expression ="qry_SearchByFiberType.Cost"
    Expression ="qry_SearchByFiberType.Description_Oracle"
    Expression ="qry_SearchByFiberType.[TBF/Ribbon]"
    Expression ="qry_SearchByFiberType.[SM/MM]"
    Alias ="Expr1"
    Expression ="Right(Left(Forms!frmCreateItem!ItemNo,6),1)"
    Expression ="qry_SearchByFiberType.SpecialNotes"
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
    0x5aebd0b0b430134cb5656d7314febea2
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="qry_SearchByFiberType.Fiber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_SearchByFiberType.Vendor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_SearchByFiberType.[OM#]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa0bd62b255329c45836d80cc147bd735
        End
    End
End
Begin
    State =0
    Left =258
    Top =346
    Right =1579
    Bottom =979
    Left =-1
    Top =-1
    Right =1289
    Bottom =266
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

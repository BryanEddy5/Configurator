Operation =1
Option =2
Where ="(((qry_SearchByFiberType.Fiber) Not Like \"*MIXED*\") And ((qry_SearchByFiberTyp"
    "e.[TBF/Ribbon])=Mid(Forms!frmCreateItem!ItemNo,12,1)) And ((qry_SearchByFiberTyp"
    "e.SpecialNotes)<>\"class 1\") And ((Forms!frmCreateItem!Fiber2Type)=qry_SearchBy"
    "FiberType!Fiber_Type Or (Forms!frmCreateItem!Fiber2Type)=qry_SearchByFiberType!A"
    "lt1_Fiber_Type Or (Forms!frmCreateItem!Fiber2Type)=qry_SearchByFiberType!Alt2_Fi"
    "ber_Type Or (Forms!frmCreateItem!Fiber2Type)=qry_SearchByFiberType!Alt3_Fiber_Ty"
    "pe Or (Forms!frmCreateItem!Fiber2Type)=qry_SearchByFiberType!Alt4_Fiber_Type))"
Begin InputTables
    Name ="qry_SearchByFiberType"
End
Begin OutputColumns
    Expression ="qry_SearchByFiberType.Fiber"
    Expression ="qry_SearchByFiberType.BIF"
    Expression ="qry_SearchByFiberType.Extension"
    Expression ="qry_SearchByFiberType.Vendor"
    Expression ="qry_SearchByFiberType.Cost"
    Expression ="qry_SearchByFiberType.Description_Oracle"
    Expression ="qry_SearchByFiberType.[TBF/Ribbon]"
    Expression ="qry_SearchByFiberType.[SM/MM]"
    Expression ="qry_SearchByFiberType.[OM#]"
    Expression ="qry_SearchByFiberType.SpecialNotes"
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
    0xc2b5245b88c13448838578a449da3cd9
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Expr1009"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x506afe3222ecd942b3878ce1aa6f8388
        End
    End
    Begin
        dbText "Name" ="qry_SearchByFiberType.Fiber"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =94
    Top =200
    Right =1931
    Bottom =818
    Left =-1
    Top =-1
    Right =1805
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

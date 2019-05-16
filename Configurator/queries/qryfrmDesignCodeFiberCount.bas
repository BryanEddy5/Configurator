Operation =1
Option =2
Where ="(((qryDesignCodeBreakdown.CableType)=Forms!frmDesignCode!CableType) And ((qryDes"
    "ignCodeBreakdown.Jacket)=Forms!frmDesignCode!Jacket) And ((qryDesignCodeBreakdow"
    "n.Location)=Forms!frmDesignCode!Passes))"
Begin InputTables
    Name ="qryDesignCodeBreakdown"
    Name ="tblFiberCount"
End
Begin OutputColumns
    Expression ="tblFiberCount.FiberCountString"
    Expression ="tblFiberCount.FiberCount"
    Expression ="qryDesignCodeBreakdown.CableType"
    Expression ="qryDesignCodeBreakdown.Jacket"
    Expression ="qryDesignCodeBreakdown.Location"
End
Begin Joins
    LeftTable ="qryDesignCodeBreakdown"
    RightTable ="tblFiberCount"
    Expression ="qryDesignCodeBreakdown.FiberCount = tblFiberCount.FiberCountString"
    Flag =1
End
Begin OrderBy
    Expression ="tblFiberCount.FiberCount"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xc654d1c51e7c2247b27930b9b3ae438d
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
End
Begin
    State =0
    Left =253
    Top =279
    Right =1865
    Bottom =576
    Left =-1
    Top =-1
    Right =1580
    Bottom =189
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =280
        Top =20
        Right =521
        Bottom =173
        Top =0
        Name ="qryDesignCodeBreakdown"
        Name =""
    End
    Begin
        Left =38
        Top =6
        Right =213
        Bottom =164
        Top =0
        Name ="tblFiberCount"
        Name =""
    End
End

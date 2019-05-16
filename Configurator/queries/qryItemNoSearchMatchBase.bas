Operation =1
Option =0
Where ="(((tblCableConstructionReferences.DesignCode)=Left(Forms!frmMainMenu!ItemNo,5) &"
    " Mid(Forms!frmMainMenu!ItemNo,7,2)) And ((tblCableConstructionReferences.Active)"
    "<>0) And ((tblCableTBType.TBLetter)=Mid(Forms!frmMainMenu!ItemNo,12,1)) And ((tb"
    "lCableConstructionReferences.ReleasedDesign)<>0))"
Begin InputTables
    Name ="tblCableTBType"
    Name ="tblCableConstructionReferences"
    Name ="tblCableConstructions"
    Name ="tblCableTightBufferReference"
End
Begin OutputColumns
    Expression ="tblCableConstructionReferences.Base"
    Expression ="tblCableConstructionReferences.DesignCode"
    Expression ="tblCableConstructionReferences.Active"
    Expression ="tblCableTBType.TBLetter"
    Expression ="tblCableTightBufferReference.TBType"
    Expression ="tblCableConstructionReferences.ReleasedDesign"
End
Begin Joins
    LeftTable ="tblCableConstructionReferences"
    RightTable ="tblCableConstructions"
    Expression ="tblCableConstructionReferences.BaseID = tblCableConstructions.BaseID"
    Flag =1
    LeftTable ="tblCableConstructionReferences"
    RightTable ="tblCableTightBufferReference"
    Expression ="tblCableConstructionReferences.TBType = tblCableTightBufferReference.TBType"
    Flag =1
    LeftTable ="tblCableTBType"
    RightTable ="tblCableTightBufferReference"
    Expression ="tblCableTBType.TBIndicatorID = tblCableTightBufferReference.TBTypeID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x9ae3b14e6d13ae4ea37357ab7028efb5
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblCableConstructionReferences.DesignCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.Active"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.Base"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableTBType.TBLetter"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableTightBufferReference.TBType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.ReleasedDesign"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-229
    Top =68
    Right =1063
    Bottom =800
    Left =-1
    Top =-1
    Right =1260
    Bottom =221
    Left =563
    Top =0
    ColumnsShown =539
    Begin
        Left =581
        Top =10
        Right =834
        Bottom =273
        Top =0
        Name ="tblCableTBType"
        Name =""
    End
    Begin
        Left =-47
        Top =19
        Right =245
        Bottom =230
        Top =0
        Name ="tblCableConstructionReferences"
        Name =""
    End
    Begin
        Left =-527
        Top =21
        Right =-239
        Bottom =169
        Top =0
        Name ="tblCableConstructions"
        Name =""
    End
    Begin
        Left =323
        Top =16
        Right =555
        Bottom =268
        Top =0
        Name ="tblCableTightBufferReference"
        Name =""
    End
End

Operation =1
Option =0
Where ="(((tblCableTightBufferReference.TBType)=Forms!frmdesigncode!tb) And ((tblCableTB"
    "Type.Active)<>0))"
Begin InputTables
    Name ="tblCableTBType"
    Name ="tblCableTightBufferReference"
End
Begin OutputColumns
    Expression ="tblCableTBType.TBLetter"
    Expression ="tblCableTBType.TBDescription"
    Expression ="tblCableTightBufferReference.TBType"
    Expression ="tblCableTBType.Active"
    Expression ="tblCableTBType.PreferenceOrder"
End
Begin Joins
    LeftTable ="tblCableTBType"
    RightTable ="tblCableTightBufferReference"
    Expression ="tblCableTBType.TBIndicatorID = tblCableTightBufferReference.TBTypeID"
    Flag =1
End
Begin OrderBy
    Expression ="tblCableTBType.PreferenceOrder"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x5a5156afe8403f468b4ca2b928aa021e
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblCableTBType.TBLetter"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableTBType.TBDescription"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableTBType.Active"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableTightBufferReference.[TBType]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableTBType.PreferenceOrder"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =25
    Top =143
    Right =1846
    Bottom =898
    Left =-1
    Top =-1
    Right =1789
    Bottom =199
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =628
        Top =12
        Right =966
        Bottom =255
        Top =0
        Name ="tblCableTBType"
        Name =""
    End
    Begin
        Left =48
        Top =12
        Right =374
        Bottom =238
        Top =0
        Name ="tblCableTightBufferReference"
        Name =""
    End
End

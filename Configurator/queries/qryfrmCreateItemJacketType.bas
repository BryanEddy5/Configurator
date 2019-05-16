Operation =1
Option =0
Having ="(((tblCableConstructionReferences.Base)=Forms!frmCreateItem!Base) And ((tblCable"
    "ConstructionReferences.Active)<>0))"
Begin InputTables
    Name ="tblCableConstructionReferences"
    Name ="tblCableConstructions"
End
Begin OutputColumns
    Alias ="Jacket"
    Expression ="tblCableConstructions.JacketMaterial"
    Expression ="tblCableConstructionReferences.Base"
    Alias ="JacketCount"
    Expression ="Count(tblCableConstructions.JacketMaterial)"
End
Begin Joins
    LeftTable ="tblCableConstructionReferences"
    RightTable ="tblCableConstructions"
    Expression ="tblCableConstructionReferences.BaseID = tblCableConstructions.BaseID"
    Flag =1
End
Begin OrderBy
    Expression ="Count(tblCableConstructions.JacketMaterial)"
    Flag =1
End
Begin Groups
    Expression ="tblCableConstructions.JacketMaterial"
    GroupLevel =0
    Expression ="tblCableConstructionReferences.Base"
    GroupLevel =0
    Expression ="tblCableConstructionReferences.Active"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xcf51d0247ec5814a96e8050d8a4629e5
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="JacketCount"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x64f6d360c046804d96de83f7769af3bb
        End
    End
    Begin
        dbText "Name" ="Jacket"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x443c2fa8c5f74c4cb27b15b04c75e905
        End
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.Base"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1579
    Bottom =881
    Left =-1
    Top =-1
    Right =1547
    Bottom =122
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =269
        Top =27
        Right =566
        Bottom =156
        Top =0
        Name ="tblCableConstructionReferences"
        Name =""
    End
    Begin
        Left =38
        Top =6
        Right =221
        Bottom =184
        Top =0
        Name ="tblCableConstructions"
        Name =""
    End
End

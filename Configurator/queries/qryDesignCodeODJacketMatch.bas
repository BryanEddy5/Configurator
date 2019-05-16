Operation =1
Option =16
RowCount ="1"
Having ="(((tblCableConstructionReferences.Base)=[Forms]![frmDesignCode]![OD]))"
Begin InputTables
    Name ="tblCableConstructionReferences"
    Name ="Product Desrcriptions"
    Name ="tblCableConstructions"
    Name ="tblCableTightBufferReference"
End
Begin OutputColumns
    Expression ="tblCableConstructionReferences.Base"
    Expression ="tblCableConstructions.NominalOD"
    Alias ="SubOD"
    Expression ="Mid([DesignCode],6,2)"
    Expression ="tblCableConstructions.CablePasses"
    Expression ="tblCableConstructions.JacketMaterial"
    Expression ="[Product Desrcriptions].Desrciprion"
    Expression ="[Product Desrcriptions].[Level 1 Desription]"
    Expression ="tblCableConstructionReferences.DesignCode"
    Expression ="tblCableTightBufferReference.TBType"
    Expression ="tblCableConstructionReferences.DesignExtension"
    Expression ="tblCableConstructionReferences.CableFamily"
    Expression ="tblCableConstructionReferences.DesignTypeID"
    Expression ="tblCableConstructionReferences.CableLevel1"
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
    LeftTable ="tblCableConstructionReferences"
    RightTable ="Product Desrcriptions"
    Expression ="tblCableConstructionReferences.SetupID = [Product Desrcriptions].ProductID"
    Flag =2
End
Begin Groups
    Expression ="tblCableConstructionReferences.Base"
    GroupLevel =0
    Expression ="tblCableConstructions.NominalOD"
    GroupLevel =0
    Expression ="Mid([DesignCode],6,2)"
    GroupLevel =0
    Expression ="tblCableConstructions.CablePasses"
    GroupLevel =0
    Expression ="tblCableConstructions.JacketMaterial"
    GroupLevel =0
    Expression ="[Product Desrcriptions].Desrciprion"
    GroupLevel =0
    Expression ="[Product Desrcriptions].[Level 1 Desription]"
    GroupLevel =0
    Expression ="tblCableConstructionReferences.DesignCode"
    GroupLevel =0
    Expression ="tblCableTightBufferReference.TBType"
    GroupLevel =0
    Expression ="tblCableConstructionReferences.DesignExtension"
    GroupLevel =0
    Expression ="tblCableConstructionReferences.CableFamily"
    GroupLevel =0
    Expression ="tblCableConstructionReferences.DesignTypeID"
    GroupLevel =0
    Expression ="tblCableConstructionReferences.CableLevel1"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x47e54d58cee66a4cb658b71b44a52dca
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
        dbText "Name" ="SubOD"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x425d3452f486af49b78ce2c3f1c2f8d4
        End
    End
    Begin
        dbText "Name" ="tblCableConstructions.NominalOD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.CablePasses"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Product Desrcriptions].[Level 1 Desription]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Product Desrcriptions].Desrciprion"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.Base"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.JacketMaterial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.DesignExtension"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.CableFamily"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.DesignTypeID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.CableLevel1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableTightBufferReference.TBType"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-235
    Top =77
    Right =1586
    Bottom =982
    Left =-1
    Top =-1
    Right =1789
    Bottom =292
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =371
        Top =16
        Right =592
        Bottom =187
        Top =0
        Name ="tblCableConstructionReferences"
        Name =""
    End
    Begin
        Left =645
        Top =12
        Right =967
        Bottom =205
        Top =0
        Name ="Product Desrcriptions"
        Name =""
    End
    Begin
        Left =1178
        Top =5
        Right =1361
        Bottom =279
        Top =0
        Name ="tblCableConstructions"
        Name =""
    End
    Begin
        Left =48
        Top =12
        Right =294
        Bottom =156
        Top =0
        Name ="tblCableTightBufferReference"
        Name =""
    End
End

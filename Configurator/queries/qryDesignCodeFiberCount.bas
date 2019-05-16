Operation =1
Option =2
Where ="(((tblCableConstructionReferences.Active)=True) And ((tblDesignTypeGroups.Design"
    "TypeGroupID)=Forms!frmdesigncode.CableType) And ((tblDesignTypes.SalesForceDescr"
    "iption)=forms!frmDesignCode!Jacket) And ((tblDesignCodeJacket.JacketDescription_"
    "SalesForce)=forms!frmDesignCode.JacketType))"
Begin InputTables
    Name ="tblDesignTypeGroups"
    Name ="tblCableConstructionReferences"
    Name ="tblCableConstructions"
    Name ="tblDesignTypes"
    Name ="tblDesignCodeJacket"
End
Begin OutputColumns
    Expression ="tblCableConstructionReferences.Fibercount"
    Expression ="tblCableConstructionReferences.Active"
    Expression ="tblDesignTypeGroups.DesignTypeGroupID"
    Expression ="tblDesignTypes.SalesForceDescription"
    Expression ="tblDesignCodeJacket.JacketDescription_SalesForce"
End
Begin Joins
    LeftTable ="tblCableConstructionReferences"
    RightTable ="tblCableConstructions"
    Expression ="tblCableConstructionReferences.BaseID = tblCableConstructions.BaseID"
    Flag =1
    LeftTable ="tblCableConstructionReferences"
    RightTable ="tblDesignTypes"
    Expression ="tblCableConstructionReferences.DesignTypeID = tblDesignTypes.DesignTypeID"
    Flag =1
    LeftTable ="tblDesignTypeGroups"
    RightTable ="tblDesignTypes"
    Expression ="tblDesignTypeGroups.DesignTypeGroupID = tblDesignTypes.DesignTypeGroupID"
    Flag =1
    LeftTable ="tblCableConstructionReferences"
    RightTable ="tblDesignCodeJacket"
    Expression ="tblCableConstructionReferences.CableLevel1 = tblDesignCodeJacket.JacketLetter"
    Flag =1
    LeftTable ="tblCableConstructionReferences"
    RightTable ="tblDesignCodeJacket"
    Expression ="tblCableConstructionReferences.CableFamily = tblDesignCodeJacket.CableLetter"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x8b386922dc601d4d97d489cae5e7f909
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblCableConstructionReferences.Active"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblDesignTypeGroups.DesignTypeGroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.Fibercount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblDesignTypes.SalesForceDescription"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblDesignCodeJacket.JacketDescription_SalesForce"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =14
    Top =60
    Right =1792
    Bottom =868
    Left =-1
    Top =-1
    Right =1746
    Bottom =91
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =1300
        Top =39
        Right =1500
        Bottom =196
        Top =0
        Name ="tblDesignTypeGroups"
        Name =""
    End
    Begin
        Left =401
        Top =33
        Right =641
        Bottom =214
        Top =0
        Name ="tblCableConstructionReferences"
        Name =""
    End
    Begin
        Left =139
        Top =27
        Right =323
        Bottom =205
        Top =0
        Name ="tblCableConstructions"
        Name =""
    End
    Begin
        Left =783
        Top =12
        Right =1000
        Bottom =156
        Top =0
        Name ="tblDesignTypes"
        Name =""
    End
    Begin
        Left =1048
        Top =12
        Right =1272
        Bottom =191
        Top =0
        Name ="tblDesignCodeJacket"
        Name =""
    End
End

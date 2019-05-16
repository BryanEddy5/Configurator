Operation =1
Option =2
Where ="(((tblDesignTypes.Active)<>0) And ((tblDesignTypeGroups.DesignTypeGroupID)=Forms"
    "!frmDesignCode!CableType))"
Begin InputTables
    Name ="tblDesignTypeGroups"
    Name ="tblDesignTypes"
End
Begin OutputColumns
    Expression ="tblDesignTypes.SalesForceDescription"
    Expression ="tblDesignTypes.DesignTypeID"
    Expression ="tblDesignTypes.Active"
    Expression ="tblDesignTypeGroups.DesignTypeGroupID"
    Expression ="tblDesignTypes.DesignTypeDesc"
End
Begin Joins
    LeftTable ="tblDesignTypeGroups"
    RightTable ="tblDesignTypes"
    Expression ="tblDesignTypeGroups.DesignTypeGroupID = tblDesignTypes.DesignTypeGroupID"
    Flag =1
End
Begin OrderBy
    Expression ="tblDesignTypes.DesignTypeDesc"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xd2c1479da4fd32409b2bcaeee0ba3277
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblDesignTypes.DesignTypeDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblDesignTypes.Active"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblDesignTypes.DesignTypeID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblDesignTypeGroups.DesignTypeGroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblDesignTypes.SalesForceDescription"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =86
    Top =142
    Right =1864
    Bottom =950
    Left =-1
    Top =-1
    Right =1746
    Bottom =222
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =258
        Top =38
        Right =465
        Bottom =222
        Top =0
        Name ="tblDesignTypeGroups"
        Name =""
    End
    Begin
        Left =703
        Top =18
        Right =992
        Bottom =258
        Top =0
        Name ="tblDesignTypes"
        Name =""
    End
End

Operation =1
Option =0
Begin InputTables
    Name ="dbo_tblDesignCode_CableImages"
    Name ="tblDesignTypes"
End
Begin OutputColumns
    Expression ="tblDesignTypes.SalesForceDescription"
    Expression ="dbo_tblDesignCode_CableImages.*"
End
Begin Joins
    LeftTable ="tblDesignTypes"
    RightTable ="dbo_tblDesignCode_CableImages"
    Expression ="tblDesignTypes.DesignTypeID = dbo_tblDesignCode_CableImages.DesignTypeID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xfd69a40ac9a61c48b2eb35262883e73a
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblDesignTypes.SalesForceDescription"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_tblDesignCode_CableImages.DesignTypeID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_tblDesignCode_CableImages.Armor_Type__c"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_tblDesignCode_CableImages.ImageURL__c"
        dbInteger "ColumnWidth" ="5955"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_tblDesignCode_CableImages.Packaging__c"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_tblDesignCode_CableImages.Product_Spec_link__c"
        dbInteger "ColumnWidth" ="5280"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1109
    Bottom =669
    Left =-1
    Top =-1
    Right =1077
    Bottom =564
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =732
        Top =214
        Right =996
        Bottom =460
        Top =0
        Name ="dbo_tblDesignCode_CableImages"
        Name =""
    End
    Begin
        Left =301
        Top =256
        Right =652
        Bottom =516
        Top =0
        Name ="tblDesignTypes"
        Name =""
    End
End

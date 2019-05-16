Operation =1
Option =2
Where ="(((dbo_vDesignCodeCableRating.SalesForceDescription)=[Forms]![frmdesigncode]![Ja"
    "cket]))"
Begin InputTables
    Name ="dbo_vDesignCodeCableRating"
End
Begin OutputColumns
    Expression ="dbo_vDesignCodeCableRating.Jacket"
    Expression ="dbo_vDesignCodeCableRating.SalesForceDescription"
    Expression ="dbo_vDesignCodeCableRating.DesignTypeID"
    Expression ="dbo_vDesignCodeCableRating.DesignTypeGroupID"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x203c0fcb862f4347b672dd53268a92fe
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="dbo_vDesignCodeCableRating.[DesignTypeGroupID]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_vDesignCodeCableRating.[DesignTypeID]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_vDesignCodeCableRating.DesignTypeID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_vDesignCodeCableRating.DesignTypeGroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_vDesignCodeCableRating.SalesForceDescription"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_vDesignCodeCableRating.Jacket"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =129
    Top =193
    Right =1864
    Bottom =979
    Left =-1
    Top =-1
    Right =1703
    Bottom =188
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =411
        Top =20
        Right =679
        Bottom =240
        Top =0
        Name ="dbo_vDesignCodeCableRating"
        Name =""
    End
End

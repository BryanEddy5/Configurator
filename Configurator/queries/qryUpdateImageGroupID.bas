Operation =1
Option =0
Where ="(((tblCableConstructionReferences.ImageGroupID)=GetCableImageid()))"
Begin InputTables
    Name ="tblCableConstructionReferences"
    Name ="tblBaseItems_local"
End
Begin OutputColumns
    Expression ="tblCableConstructionReferences.ImageGroupID"
End
Begin Joins
    LeftTable ="tblCableConstructionReferences"
    RightTable ="tblBaseItems_local"
    Expression ="tblCableConstructionReferences.SetupID = tblBaseItems_local.SetupID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xe20a02d6d72362488942fb380b58d6e7
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblCableConstructionReferences.ImageGroupID"
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
    Bottom =563
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =447
        Top =55
        Right =828
        Bottom =494
        Top =0
        Name ="tblCableConstructionReferences"
        Name =""
    End
    Begin
        Left =199
        Top =103
        Right =343
        Bottom =447
        Top =0
        Name ="tblBaseItems_local"
        Name =""
    End
End

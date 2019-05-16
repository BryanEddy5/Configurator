Operation =1
Option =0
Where ="(((tblProductCategorySegment1.Segment1Desc)=Forms!frmMainMenu!ItemSetupOrg) And "
    "((tblProductCategorySegment1.Active)=-1))"
Begin InputTables
    Name ="tblProductCategorySegment1"
End
Begin OutputColumns
    Expression ="tblProductCategorySegment1.*"
    Alias ="ProductType"
    Expression ="[Segment1ID] & \".0.0.0.0\""
    Alias ="ProductTypeView"
    Expression ="tblProductCategorySegment1.Segment1Desc"
End
Begin Parameters
    Name ="[Forms]![frmMainMenu]![ItemSetupOrg]"
    Flag =10
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x35f76a7ccd641e43b13c0337a782feb8
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblProductCategorySegment1.Segment1Desc"
        dbInteger "ColumnWidth" ="1560"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ProductTypeView"
        dbInteger "ColumnWidth" ="1830"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x00aefeee8259894ba7f33c6b28824624
        End
    End
    Begin
        dbText "Name" ="ProductType"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9819974dc9ec2a489affc5095f3ca528
        End
    End
End
Begin
    State =0
    Left =47
    Top =69
    Right =1156
    Bottom =494
    Left =-1
    Top =-1
    Right =1077
    Bottom =215
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =287
        Bottom =105
        Top =0
        Name ="tblProductCategorySegment1"
        Name =""
    End
End

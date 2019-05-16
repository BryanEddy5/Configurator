Operation =1
Option =0
Having ="(((Mid([Basic Product Construction].[Item No],12,1))<>\"s\" And (Mid([Basic Prod"
    "uct Construction].[Item No],12,1))<>\"c\") AND (([Basic Product Construction].[T"
    "B Material]) Not Like \"frc*\") AND (([Basic Product Construction].Active)=Yes))"
Begin InputTables
    Name ="Basic Product Construction"
    Name ="tblCableConstructionReferences"
    Name ="tblCableConstructions"
End
Begin OutputColumns
    Alias ="Family"
    Expression ="Left([Basic Product Construction].[Item No],2)"
    Alias ="TBType"
    Expression ="Mid([Basic Product Construction].[Item No],12,1)"
    Alias ="Jacket Material"
    Expression ="tblCableConstructions.JacketMaterial"
    Expression ="[Basic Product Construction].[TB Material]"
    Expression ="[Basic Product Construction].[TB Chips Type]"
    Alias ="CountOfTB Material"
    Expression ="Count([Basic Product Construction].[TB Material])"
    Expression ="[Basic Product Construction].[EZ Strip]"
    Expression ="[Basic Product Construction].[TB Nominal OD]"
End
Begin Joins
    LeftTable ="Basic Product Construction"
    RightTable ="tblCableConstructionReferences"
    Expression ="[Basic Product Construction].Base = tblCableConstructionReferences.Base"
    Flag =1
    LeftTable ="tblCableConstructionReferences"
    RightTable ="tblCableConstructions"
    Expression ="tblCableConstructionReferences.BaseID = tblCableConstructions.BaseID"
    Flag =1
End
Begin OrderBy
    Expression ="Left([Basic Product Construction].[Item No],2)"
    Flag =0
    Expression ="Mid([Basic Product Construction].[Item No],12,1)"
    Flag =0
    Expression ="tblCableConstructions.JacketMaterial"
    Flag =0
    Expression ="Count([Basic Product Construction].[TB Material])"
    Flag =1
End
Begin Groups
    Expression ="Left([Basic Product Construction].[Item No],2)"
    GroupLevel =0
    Expression ="Mid([Basic Product Construction].[Item No],12,1)"
    GroupLevel =0
    Expression ="tblCableConstructions.JacketMaterial"
    GroupLevel =0
    Expression ="[Basic Product Construction].[TB Material]"
    GroupLevel =0
    Expression ="[Basic Product Construction].[TB Chips Type]"
    GroupLevel =0
    Expression ="[Basic Product Construction].[EZ Strip]"
    GroupLevel =0
    Expression ="[Basic Product Construction].[TB Nominal OD]"
    GroupLevel =0
    Expression ="[Basic Product Construction].Active"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbInteger "RowHeight" ="270"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xf69fb67ba1cfb24bbeb9b061d4488b69
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="TBType"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x455a6d95ce63424785fb85db2f028e3e
        End
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Jacket Material]"
        dbInteger "ColumnOrder" ="3"
        dbInteger "ColumnWidth" ="1980"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[TB Material]"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[TB Chips Type]"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CountOfTB Material"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x47af6142a82a204794487db7d7876d0f
        End
    End
    Begin
        dbText "Name" ="Family"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x836355ac27d75348bba1d09e7ac63f9b
        End
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[EZ Strip]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[TB Nominal OD]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.[Jacket Material]"
        dbInteger "ColumnWidth" ="1980"
        dbInteger "ColumnOrder" ="3"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.JacketMaterial"
        dbInteger "ColumnWidth" ="1980"
        dbInteger "ColumnOrder" ="3"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Jacket Material"
        dbInteger "ColumnWidth" ="1980"
        dbInteger "ColumnOrder" ="3"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6628328863802e41a9223b35d2f5ed03
        End
    End
End
Begin
    State =0
    Left =146
    Top =347
    Right =1584
    Bottom =1169
    Left =-1
    Top =-1
    Right =1406
    Bottom =221
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =424
        Bottom =261
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
    Begin
        Left =657
        Top =24
        Right =951
        Bottom =169
        Top =0
        Name ="tblCableConstructionReferences"
        Name =""
    End
    Begin
        Left =1156
        Top =73
        Right =1300
        Bottom =217
        Top =0
        Name ="tblCableConstructions"
        Name =""
    End
End

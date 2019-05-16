Operation =1
Option =0
Having ="((([Basic Product Construction].[Special Instr Product1]) Not Like \"*box*\" And"
    " ([Basic Product Construction].[Special Instr Product1]) Not Like \"*afl*\") AND"
    " (([Basic Product Construction].Active)=True))"
Begin InputTables
    Name ="Basic Product Construction"
End
Begin OutputColumns
    Expression ="[Basic Product Construction].Base"
    Alias ="CountOfBase"
    Expression ="Count([Basic Product Construction].Base)"
    Expression ="[Basic Product Construction].[Standard Operation]"
    Expression ="[Basic Product Construction].[Jacket Material]"
    Expression ="[Basic Product Construction].[Nominal OD]"
    Expression ="[Basic Product Construction].[OD Tol  (+)]"
    Expression ="[Basic Product Construction].[OD Tol  (-)]"
    Expression ="[Basic Product Construction].Height"
    Expression ="[Basic Product Construction].[Height Tol (+)]"
    Expression ="[Basic Product Construction].[Height Tol (-)]"
    Expression ="[Basic Product Construction].[Nominal Wall]"
    Expression ="[Basic Product Construction].[Max Ave Wall]"
    Expression ="[Basic Product Construction].[Min Ave Wall]"
    Expression ="[Basic Product Construction].[Aramid Type]"
    Expression ="[Basic Product Construction].[Number of ends]"
    Expression ="[Basic Product Construction].[Aramid Type2]"
    Expression ="[Basic Product Construction].[Unit Chips Type]"
    Expression ="[Basic Product Construction].[Unit Ripcord]"
    Expression ="[Basic Product Construction].[Unit CM]"
    Expression ="[Basic Product Construction].[Lay Length]"
    Expression ="[Basic Product Construction].Talc"
    Expression ="[Basic Product Construction].[TB Chips Type]"
    Expression ="[Basic Product Construction].[TB Nominal OD]"
    Expression ="[Basic Product Construction].[TB OD Tol  (+)]"
    Expression ="[Basic Product Construction].[TB OD Tol  (-)]"
    Expression ="[Basic Product Construction].[Special Instr Product1]"
    Expression ="[Basic Product Construction].[zzSpecial Instr Product3]"
    Expression ="[Basic Product Construction].[zzSpecial Instr Product4]"
    Expression ="[Basic Product Construction].[PSS Document #]"
    Expression ="[Basic Product Construction].Active"
End
Begin OrderBy
    Expression ="[Basic Product Construction].Base"
    Flag =0
    Expression ="Count([Basic Product Construction].Base)"
    Flag =1
End
Begin Groups
    Expression ="[Basic Product Construction].Base"
    GroupLevel =0
    Expression ="[Basic Product Construction].[Standard Operation]"
    GroupLevel =0
    Expression ="[Basic Product Construction].[Jacket Material]"
    GroupLevel =0
    Expression ="[Basic Product Construction].[Nominal OD]"
    GroupLevel =0
    Expression ="[Basic Product Construction].[OD Tol  (+)]"
    GroupLevel =0
    Expression ="[Basic Product Construction].[OD Tol  (-)]"
    GroupLevel =0
    Expression ="[Basic Product Construction].Height"
    GroupLevel =0
    Expression ="[Basic Product Construction].[Height Tol (+)]"
    GroupLevel =0
    Expression ="[Basic Product Construction].[Height Tol (-)]"
    GroupLevel =0
    Expression ="[Basic Product Construction].[Nominal Wall]"
    GroupLevel =0
    Expression ="[Basic Product Construction].[Max Ave Wall]"
    GroupLevel =0
    Expression ="[Basic Product Construction].[Min Ave Wall]"
    GroupLevel =0
    Expression ="[Basic Product Construction].[Aramid Type]"
    GroupLevel =0
    Expression ="[Basic Product Construction].[Number of ends]"
    GroupLevel =0
    Expression ="[Basic Product Construction].[Aramid Type2]"
    GroupLevel =0
    Expression ="[Basic Product Construction].[Unit Chips Type]"
    GroupLevel =0
    Expression ="[Basic Product Construction].[Unit Ripcord]"
    GroupLevel =0
    Expression ="[Basic Product Construction].[Unit CM]"
    GroupLevel =0
    Expression ="[Basic Product Construction].[Lay Length]"
    GroupLevel =0
    Expression ="[Basic Product Construction].Talc"
    GroupLevel =0
    Expression ="[Basic Product Construction].[TB Chips Type]"
    GroupLevel =0
    Expression ="[Basic Product Construction].[TB Nominal OD]"
    GroupLevel =0
    Expression ="[Basic Product Construction].[TB OD Tol  (+)]"
    GroupLevel =0
    Expression ="[Basic Product Construction].[TB OD Tol  (-)]"
    GroupLevel =0
    Expression ="[Basic Product Construction].[Special Instr Product1]"
    GroupLevel =0
    Expression ="[Basic Product Construction].[zzSpecial Instr Product3]"
    GroupLevel =0
    Expression ="[Basic Product Construction].[zzSpecial Instr Product4]"
    GroupLevel =0
    Expression ="[Basic Product Construction].[PSS Document #]"
    GroupLevel =0
    Expression ="[Basic Product Construction].Active"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xf5217cd5eec67f47a0233340da8e2465
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="[Basic Product Construction].[Special Instr Product1]"
        dbInteger "ColumnWidth" ="16965"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CountOfBase"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9f94c77ec1200e409d89a730e9202073
        End
    End
End
Begin
    State =0
    Left =190
    Top =104
    Right =1898
    Bottom =540
    Left =-1
    Top =-1
    Right =1676
    Bottom =205
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =253
        Bottom =214
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
End

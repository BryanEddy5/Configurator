Operation =3
Name ="tblCutSheetApproval_Local"
Option =0
Where ="(((tblCutSheetApproval.Technical_Approval)=False) AND ((tblCutSheetApproval.Requ"
    "ested)=True))"
Begin InputTables
    Name ="tblCutSheetApproval"
    Name ="Basic Product Construction"
End
Begin OutputColumns
    Name ="Base"
    Expression ="tblCutSheetApproval.Base"
    Name ="Oracle"
    Expression ="tblCutSheetApproval.ItemRequested"
    Name ="ItemNo"
    Expression ="[Basic Product Construction].[Item No]"
    Name ="RequestedDate"
    Expression ="tblCutSheetApproval.RequestedDate"
End
Begin Joins
    LeftTable ="tblCutSheetApproval"
    RightTable ="Basic Product Construction"
    Expression ="tblCutSheetApproval.ItemRequested = [Basic Product Construction].[New Oracle Par"
        "t #]"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x730cfd2c42318a469b5806d1ded68e7a
End
Begin
    Begin
        dbText "Name" ="vCutSheet_AwaitingApproval.Technical_Approval"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vCutSheet_AwaitingApproval.Base"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vCutSheet_AwaitingApproval.Oracle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_CableUnion.ItemNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_CableUnion.Base"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vCutSheet_TechApproval.Oracle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vCutSheet_TechApproval.Base"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCutSheetApproval.Technical_Approval"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCutSheetApproval.ItemRequested"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCutSheetApproval.Base"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].ItemNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCutSheetApproval.Oracle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Item No]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCutSheetApproval.Technical_Approver"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCutSheetApproval.RequestedDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCutSheetApproval.Requested"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1551
    Bottom =957
    Left =-1
    Top =-1
    Right =1519
    Bottom =253
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =552
        Top =-2
        Right =815
        Bottom =240
        Top =0
        Name ="tblCutSheetApproval"
        Name =""
    End
    Begin
        Left =114
        Top =-2
        Right =336
        Bottom =246
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
End

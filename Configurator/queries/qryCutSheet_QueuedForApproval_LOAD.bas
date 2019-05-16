Operation =3
Name ="tblCutSheet_QueuedForApproval"
Option =0
Where ="(((tblCutSheetApproval_Local.Technical_Approval)<>0))"
Begin InputTables
    Name ="tblCutSheetApproval_Local"
End
Begin OutputColumns
    Name ="Base"
    Expression ="tblCutSheetApproval_Local.Base"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0xd4ebad51d7ccc1439ba6a4ac35b22176
End
Begin
    Begin
        dbText "Name" ="vCutSheet_AwaitingApproval.Base"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vCutSheet_AwaitingApproval.Technical_Approval"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCutSheetApproval_Local.Base"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCutSheetApproval_Local.Technical_Approval"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1603
    Bottom =950
    Left =-1
    Top =-1
    Right =1571
    Bottom =559
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =642
        Top =210
        Right =786
        Bottom =354
        Top =0
        Name ="tblCutSheetApproval_Local"
        Name =""
    End
End

Operation =1
Option =0
Having ="(((tblCableConstructions.CablePasses)=Forms!frmCreateItem!Location) And ((tblCab"
    "leConstructions.DesignCode) Like Left(Forms!frmMainMenu!ItemNo,5) & Mid(Forms!fr"
    "mMainMenu!ItemNo,7,2)) And ((tblCableConstructions.Active)<>0))"
Begin InputTables
    Name ="tblCableConstructions"
End
Begin OutputColumns
    Alias ="Expr1"
    Expression ="tblCableConstructions.Base"
    Alias ="Location"
    Expression ="tblCableConstructions.CablePasses"
    Alias ="Expr2"
    Expression ="tblCableConstructions.DesignCode"
    Alias ="Expr3"
    Expression ="tblCableConstructions.Active"
End
Begin Groups
    Expression ="tblCableConstructions.Base"
    GroupLevel =0
    Expression ="tblCableConstructions.CablePasses"
    GroupLevel =0
    Expression ="tblCableConstructions.DesignCode"
    GroupLevel =0
    Expression ="tblCableConstructions.Active"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xfdf9e4995406bb4db779e0734ecca191
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3cef0770068aaa4e85138c1257ca1fe0
        End
    End
    Begin
        dbText "Name" ="Location"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x1c2cc32cef28754695e2bb982e38417e
        End
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3e3133d4c355914ca65dabb3e31ed6d6
        End
    End
    Begin
        dbText "Name" ="Expr3"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa781f8d1be8bd547971b77a3da0110d2
        End
    End
End
Begin
    State =0
    Left =163
    Top =358
    Right =1455
    Bottom =740
    Left =-1
    Top =-1
    Right =1260
    Bottom =151
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =413
        Top =6
        Right =509
        Bottom =124
        Top =0
        Name ="tblCableConstructions"
        Name =""
    End
End

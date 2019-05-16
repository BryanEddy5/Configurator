Operation =3
Name ="tblSubUnitHold"
Option =0
Begin InputTables
    Name ="qrytblSubUnit_X_Hold"
    Name ="Basic Product Construction"
End
Begin OutputColumns
    Expression ="qrytblSubUnit_X_Hold.*"
End
Begin Joins
    LeftTable ="qrytblSubUnit_X_Hold"
    RightTable ="Basic Product Construction"
    Expression ="qrytblSubUnit_X_Hold.[New Oracle Part #] = [Basic Product Construction].[New Ora"
        "cle Part #]"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0xb49d7e56072c234d8560ca70e9572bc0
End
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="qrytblSubUnit_X_Hold.[Basic Product Construction_1].[Fiber Oracle item]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qrytblSubUnit_X_Hold.[Basic Product Construction_1].[Jacket Material]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qrytblSubUnit_X_Hold.[Basic Product Construction_1].OracleStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qrytblSubUnit_X_Hold.[Basic Product Construction_1].[New Oracle Part #]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qrytblSubUnit_X_Hold.[Basic Product Construction_1].[Item No]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qrytblSubUnit_X_Hold.Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qrytblSubUnit_X_Hold.[Basic Product Construction_1].Active"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qrytblSubUnit_X_Hold.TB_Colr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qrytblSubUnit_X_Hold.Expr1008"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qrytblSubUnit_X_Hold.[Product Desrcriptions].Product"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-45
    Top =54
    Right =1064
    Bottom =543
    Left =-1
    Top =-1
    Right =1077
    Bottom =224
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =157
        Top =11
        Right =556
        Bottom =174
        Top =0
        Name ="qrytblSubUnit_X_Hold"
        Name =""
    End
    Begin
        Left =684
        Top =15
        Right =1043
        Bottom =221
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
End

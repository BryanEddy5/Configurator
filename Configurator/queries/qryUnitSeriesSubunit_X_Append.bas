Operation =3
Name ="tblUnitSeriesSubunit"
Option =2
Where ="((([Basic Product Construction_1].[Fiber Oracle item])=Forms!frmCreateItem!Fiber"
    ") And (([Basic Product Construction_1].[New Oracle Part #]) Not Like \"*test*\")"
    " And (([Basic Product Construction_1].Active)=True) And ((Mid([Basic Product Con"
    "struction_1]![Item No],10,1))=Forms!frmCreateItem!Subcolor) And ((Mid([Basic Pro"
    "duct Construction_1]![Item No],11,2))=Forms!frmCreateItem!TBcolor & Forms!frmCre"
    "ateItem!TBtype) And (([Product Desrcriptions].Product)=Forms!frmCreateItem!Base)"
    ")"
Begin InputTables
    Name ="tblUnitSeriesCopy"
    Name ="Basic Product Construction"
    Name ="Basic Product Construction"
    Alias ="Basic Product Construction_1"
    Name ="Product Desrcriptions"
End
Begin OutputColumns
    Name ="Fiber Oracle item"
    Expression ="[Basic Product Construction_1].[Fiber Oracle item]"
    Alias ="Expr1"
    Name ="Expr1"
    Expression ="[Basic Product Construction_1]![Item No]"
    Name ="OracleStatus"
    Expression ="[Basic Product Construction_1].OracleStatus"
    Name ="New Oracle Part #"
    Expression ="[Basic Product Construction_1].[New Oracle Part #]"
    Name ="Active"
    Expression ="[Basic Product Construction_1].Active"
    Name ="Item No"
    Expression ="[Basic Product Construction_1].[Item No]"
End
Begin Joins
    LeftTable ="Basic Product Construction"
    RightTable ="Basic Product Construction_1"
    Expression ="[Basic Product Construction].[TB Material] = [Basic Product Construction_1].[TB "
        "Material]"
    Flag =1
    LeftTable ="Basic Product Construction"
    RightTable ="Basic Product Construction_1"
    Expression ="[Basic Product Construction].Base = [Basic Product Construction_1].Base"
    Flag =1
    LeftTable ="tblUnitSeriesCopy"
    RightTable ="Basic Product Construction"
    Expression ="tblUnitSeriesCopy.[Unit Type] = [Basic Product Construction].[Item No]"
    Flag =1
    LeftTable ="Basic Product Construction_1"
    RightTable ="Product Desrcriptions"
    Expression ="[Basic Product Construction_1].Base = [Product Desrcriptions].[Level 1 Product]"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0xab7f63358532c443b5ce42f903f6ee52
End
Begin
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6ce4b8ec3d09644ca5f8dad49d4450b1
        End
    End
    Begin
        dbText "Name" ="[Basic Product Construction_1].[Fiber Oracle item]"
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
    Bottom =224
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =38
        Top =6
        Right =195
        Bottom =184
        Top =0
        Name ="tblUnitSeriesCopy"
        Name =""
    End
    Begin
        Left =267
        Top =27
        Right =441
        Bottom =250
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
    Begin
        Left =562
        Top =7
        Right =764
        Bottom =260
        Top =0
        Name ="Basic Product Construction_1"
        Name =""
    End
    Begin
        Left =853
        Top =3
        Right =1112
        Bottom =181
        Top =0
        Name ="Product Desrcriptions"
        Name =""
    End
End

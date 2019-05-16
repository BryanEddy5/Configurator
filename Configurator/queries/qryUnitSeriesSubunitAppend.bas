Operation =3
Name ="tblUnitSeriesSubunit"
Option =2
Where ="((([Basic Product Construction_1].[Fiber Oracle item])=Forms!frmCreateItem!Fiber"
    ") And (([Basic Product Construction_1].[New Oracle Part #]) Not Like \"*test*\")"
    " And (([Basic Product Construction_1].Active)=True) And ((Mid([Basic Product Con"
    "struction_1]![Item No],10,1))=Forms!frmCreateItem!SubColor) And ((Mid([Basic Pro"
    "duct Construction_1]![Item No],11,2))=Forms!frmCreateItem!TBColor & Forms!frmCre"
    "ateItem!TBtype) And (([Product Desrcriptions].Product)=Forms!frmCreateItem!Base)"
    ")"
Begin InputTables
    Name ="Product Desrcriptions"
    Name ="tblUnitSeriesCopy"
    Name ="Basic Product Construction"
    Name ="Basic Product Construction"
    Alias ="Basic Product Construction_1"
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
    Expression ="[Basic Product Construction].Base = [Basic Product Construction_1].Base"
    Flag =1
    LeftTable ="Basic Product Construction"
    RightTable ="Basic Product Construction_1"
    Expression ="[Basic Product Construction].[TB Material] = [Basic Product Construction_1].[TB "
        "Material]"
    Flag =1
    LeftTable ="tblUnitSeriesCopy"
    RightTable ="Basic Product Construction"
    Expression ="tblUnitSeriesCopy.[Unit Type] = [Basic Product Construction].[Item No]"
    Flag =1
    LeftTable ="Product Desrcriptions"
    RightTable ="Basic Product Construction_1"
    Expression ="[Product Desrcriptions].[Level 1 Product] = [Basic Product Construction_1].Base"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0xf7d803115721304db6dad1dcec02496a
End
Begin
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe0c5338c99fbc94ab64d68ae472e3f86
        End
    End
End
Begin
    State =0
    Left =34
    Top =24
    Right =1143
    Bottom =509
    Left =-1
    Top =-1
    Right =1077
    Bottom =237
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =1040
        Top =34
        Right =1271
        Bottom =245
        Top =0
        Name ="Product Desrcriptions"
        Name =""
    End
    Begin
        Left =38
        Top =6
        Right =240
        Bottom =184
        Top =0
        Name ="tblUnitSeriesCopy"
        Name =""
    End
    Begin
        Left =321
        Top =0
        Right =460
        Bottom =283
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
    Begin
        Left =619
        Top =29
        Right =924
        Bottom =252
        Top =0
        Name ="Basic Product Construction_1"
        Name =""
    End
End

Operation =2
Name ="tblBaseConstructionSinglePass"
Option =0
Where ="(((qryBaseConstruction.[Standard Operation]) Is Not Null And (qryBaseConstructio"
    "n.[Standard Operation])<>\"\"))"
Begin InputTables
    Name ="qryBaseConstruction"
    Name ="qryBaseConstructionMax"
End
Begin OutputColumns
    Expression ="qryBaseConstruction.*"
End
Begin Joins
    LeftTable ="qryBaseConstruction"
    RightTable ="qryBaseConstructionMax"
    Expression ="qryBaseConstruction.CountOfBase = qryBaseConstructionMax.MaxOfCountOfBase"
    Flag =1
    LeftTable ="qryBaseConstruction"
    RightTable ="qryBaseConstructionMax"
    Expression ="qryBaseConstruction.Base = qryBaseConstructionMax.Base"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x815f234374da384399144878f1686414
End
Begin
End
Begin
    State =0
    Left =216
    Top =35
    Right =1881
    Bottom =877
    Left =-1
    Top =-1
    Right =1633
    Bottom =171
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =172
        Top =6
        Right =449
        Bottom =192
        Top =0
        Name ="qryBaseConstruction"
        Name =""
    End
    Begin
        Left =544
        Top =3
        Right =663
        Bottom =171
        Top =0
        Name ="qryBaseConstructionMax"
        Name =""
    End
End

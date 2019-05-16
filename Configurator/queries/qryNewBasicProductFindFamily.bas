Operation =1
Option =2
Where ="((([Basic Product Construction].[New Oracle Part #]) Not Like \"*test*\") And (("
    "[Basic Product Construction].Base)=tblNewProductConstructionHold!Base))"
Begin InputTables
    Name ="Basic Product Construction"
    Name ="tblNewProductConstructionHold"
End
Begin OutputColumns
    Expression ="[Basic Product Construction].*"
End
Begin OrderBy
    Expression ="[Basic Product Construction].[New Oracle Part #]"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x1a2c91eab0912e4082992abcb5403705
End
dbMemo "Filter" ="((qryNewBasicProductFindFamily.[New Oracle Part #] Like \"PX01856*\"))"
dbMemo "OrderBy" ="qryNewBasicProductFindFamily.[Revision Date]"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
End
Begin
    State =0
    Left =0
    Top =40
    Right =1292
    Bottom =772
    Left =-1
    Top =-1
    Right =1260
    Bottom =305
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =70
        Top =5
        Right =426
        Bottom =273
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
    Begin
        Left =464
        Top =6
        Right =842
        Bottom =214
        Top =0
        Name ="tblNewProductConstructionHold"
        Name =""
    End
End

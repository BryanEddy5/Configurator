Operation =1
Option =0
Where ="(((qryBasePrintLine_1_2.[New Oracle Part #])=[Forms]![frmFiberMatchExisting]![Or"
    "acle]))"
Begin InputTables
    Name ="qryBasePrintLine_1_2"
End
Begin OutputColumns
    Alias ="Oracle"
    Expression ="qryBasePrintLine_1_2.[New Oracle Part #]"
    Expression ="qryBasePrintLine_1_2.PrintLine1"
    Expression ="qryBasePrintLine_1_2.PrintLine2"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xec17baf0f8280c40b424ef23bd03c1b2
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Oracle"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xcefe9ff7823f7042846fc27b1cd01f78
        End
    End
End
Begin
    State =0
    Left =27
    Top =376
    Right =1959
    Bottom =891
    Left =-1
    Top =-1
    Right =1900
    Bottom =146
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =239
        Bottom =124
        Top =0
        Name ="qryBasePrintLine_1_2"
        Name =""
    End
End

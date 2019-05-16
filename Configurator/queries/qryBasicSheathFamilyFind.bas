Operation =1
Option =2
Where ="((([Basic (Cable/Sheath)].[New Oracle Part Number]) Not Like \"*test*\") And ((["
    "Basic (Cable/Sheath)].Base)=tblNewBasicCable!Base))"
Begin InputTables
    Name ="Basic (Cable/Sheath)"
    Name ="tblNewBasicCable"
End
Begin OutputColumns
    Expression ="[Basic (Cable/Sheath)].*"
End
Begin OrderBy
    Expression ="[Basic (Cable/Sheath)].[Revision Date]"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x2473d69577adb34ebb5c95ef889a8472
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
End
Begin
    State =0
    Left =0
    Top =40
    Right =1579
    Bottom =881
    Left =-1
    Top =-1
    Right =1547
    Bottom =285
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =557
        Top =6
        Right =815
        Bottom =259
        Top =0
        Name ="Basic (Cable/Sheath)"
        Name =""
    End
    Begin
        Left =208
        Top =39
        Right =304
        Bottom =157
        Top =0
        Name ="tblNewBasicCable"
        Name =""
    End
End

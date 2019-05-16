Operation =1
Option =0
Begin InputTables
    Name ="tblFiberElementValue"
    Name ="qryCableFamily"
End
Begin OutputColumns
    Expression ="qryCableFamily.CableFamily"
    Expression ="tblFiberElementValue.FiberLetter"
    Alias ="CountOfCableFamily"
    Expression ="Count(qryCableFamily.CableFamily)"
    Expression ="tblFiberElementValue.FiberValue"
End
Begin Groups
    Expression ="qryCableFamily.CableFamily"
    GroupLevel =0
    Expression ="tblFiberElementValue.FiberLetter"
    GroupLevel =0
    Expression ="tblFiberElementValue.FiberValue"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xf7013d18c6b34b4785459e3485bae2df
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="CountOfCableFamily"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x5189e0aa204e344e8cbf4fe356f52364
        End
    End
End
Begin
    State =0
    Left =246
    Top =257
    Right =1665
    Bottom =581
    Left =-1
    Top =-1
    Right =1387
    Bottom =93
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =109
        Top =0
        Name ="tblFiberElementValue"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =94
        Top =0
        Name ="qryCableFamily"
        Name =""
    End
End

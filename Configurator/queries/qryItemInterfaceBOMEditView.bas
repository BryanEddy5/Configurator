Operation =1
Option =0
Begin InputTables
    Name ="tblItemLoadTempBOM"
End
Begin OutputColumns
    Expression ="tblItemLoadTempBOM.*"
    Alias ="PTCheck"
    Expression ="IIf(Left([ComponentItemNumber],2)=\"PT\",1,0)"
End
Begin OrderBy
    Expression ="tblItemLoadTempBOM.OperationSeqNumber"
    Flag =0
    Expression ="tblItemLoadTempBOM.ItemSeqNumber"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xa9e1156e729065488510bdd02fdae650
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="PTCheck"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc97ddafc84a3f64ebfe9ec7fe1619b60
        End
    End
    Begin
        dbText "Name" ="tblItemLoadTempBOM.ItemSeqNumber"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1560
    Bottom =854
    Left =-1
    Top =-1
    Right =1528
    Bottom =247
    Left =1529
    Top =0
    ColumnsShown =539
    Begin
        Left =-1145
        Top =6
        Right =-934
        Bottom =109
        Top =0
        Name ="tblItemLoadTempBOM"
        Name =""
    End
End

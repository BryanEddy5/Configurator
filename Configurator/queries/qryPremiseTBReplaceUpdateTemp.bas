Operation =4
Option =0
Where ="(((Left([ComponentItemNumber],7))=Left([Forms]![frmPremiseTBReplace]![ComponentI"
    "temNumber],7)))"
Begin InputTables
    Name ="tblItemLoadTempBOM"
End
Begin OutputColumns
    Name ="tblItemLoadTempBOM.ComponentItemNumber"
    Expression ="IIf(Left([Forms]![frmPremiseTBReplace]![NewComponentItem],3)=\"rbn\" Or Left([Fo"
        "rms]![frmPremiseTBReplace]![NewComponentItem],3)=\"swr\",[Forms]![frmPremiseTBRe"
        "place]![NewComponentItem],Left([Forms]![frmPremiseTBReplace]![cboBaseSelect],7) "
        "& \"-\" & Right([ComponentItemNumber],2))"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbBinary "GUID" = Begin
    0x6d8a28ec0b6fb5439d1440f103c242e9
End
Begin
    Begin
        dbText "Name" ="tblItemLoadTempBOM.ComponentItemNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Left([ComponentItemNumber],7)"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-425
    Top =32
    Right =684
    Bottom =356
    Left =-1
    Top =-1
    Right =1077
    Bottom =93
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =188
        Top =6
        Right =385
        Bottom =124
        Top =0
        Name ="tblItemLoadTempBOM"
        Name =""
    End
End

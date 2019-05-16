Operation =4
Option =0
Begin InputTables
    Name ="tblNewPrductConstruction"
End
Begin OutputColumns
    Name ="tblNewPrductConstruction.[New Oracle Part #]"
    Expression ="Forms!frmsubunitnextitemno!NewOracle"
    Name ="tblNewPrductConstruction.[Item No]"
    Expression ="Left([Item No],5) & Forms!frmCreateItem!FiberType & Mid([Item No],7,3) & Forms!f"
        "rmCreateItem!SubColor & Forms!frmCreateItem!TBColor & Forms!frmCreateItem!TBType"
        " & Mid([Item No],13,7)"
    Name ="tblNewPrductConstruction.Active"
    Expression ="-1"
    Name ="tblNewPrductConstruction.Reason"
    Expression ="\"\""
    Name ="tblNewPrductConstruction.[Revision Date]"
    Expression ="Date()"
    Name ="tblNewPrductConstruction.[Revision Letter]"
    Expression ="\"1\""
    Name ="tblNewPrductConstruction.[Revision History]"
    Expression ="Date() & \" - INITIAL ISSUE \" & Forms!frmSwitchboard!Initials"
    Name ="tblNewPrductConstruction.[Jacket Color]"
    Expression ="Forms!frmSubUnitJacketColor!JacketColor"
    Name ="tblNewPrductConstruction.[Color Chip ID]"
    Expression ="Forms!frmSubUnitJacketColor!ColorChip"
    Name ="tblNewPrductConstruction.[Fiber Type]"
    Expression ="Forms!frmFiberSpec!FiberType"
    Name ="tblNewPrductConstruction.[Fiber Oracle item]"
    Expression ="Forms!frmFiberSpec!Fiber"
    Name ="tblNewPrductConstruction.[1st Req Freq]"
    Expression ="Forms!frmFiberSpec![1stReqFreq]"
    Name ="tblNewPrductConstruction.[1st Max Atten]"
    Expression ="Forms!frmFiberSpec![1stMaxAtten]"
    Name ="tblNewPrductConstruction.[1st Min BandW]"
    Expression ="Forms!frmFiberSpec![1stMinBandW]"
    Name ="tblNewPrductConstruction.[2nd Req Freq]"
    Expression ="Forms!frmFiberSpec![2ndReqFreq]"
    Name ="tblNewPrductConstruction.[2nd Max Atten]"
    Expression ="Forms!frmFiberSpec![2ndMaxAtten]"
    Name ="tblNewPrductConstruction.[2nd Min BandW]"
    Expression ="Forms!frmFiberSpec![2ndMinBandW]"
    Name ="tblNewPrductConstruction.[SM 1300 Max Atten]"
    Expression ="\"\""
    Name ="tblNewPrductConstruction.[SM 1550 Max Atten]"
    Expression ="\"\""
    Name ="tblNewPrductConstruction.OracleStatus"
    Expression ="\"Acitve\""
    Name ="tblNewPrductConstruction.FiberType2"
    Expression ="\"NONE\""
    Name ="tblNewPrductConstruction.FiberType3"
    Expression ="\"NONE\""
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xda2a340311e2de43b32fc80a149a7074
End
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
Begin
    Begin
        dbText "Name" ="tblNewPrductConstruction.[New Oracle Part #]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.FiberType2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.FiberType3"
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
    Bottom =159
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =42
        Top =0
        Right =361
        Bottom =193
        Top =0
        Name ="tblNewPrductConstruction"
        Name =""
    End
End

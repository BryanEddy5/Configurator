Operation =4
Option =0
Begin InputTables
    Name ="tblNewPrductConstruction"
End
Begin OutputColumns
    Name ="tblNewPrductConstruction.[TB Material]"
    Expression ="Forms!frmTBMaterial!TBMaterial"
    Name ="tblNewPrductConstruction.[TB Chips Type]"
    Expression ="Forms!frmTBMaterial!TBChips"
    Name ="tblNewPrductConstruction.[EZ Strip]"
    Expression ="Forms!frmTBMaterial!EZStrip"
    Name ="tblNewPrductConstruction.[TB Nominal OD]"
    Expression ="Forms!frmTBMaterial!TBOD"
    Name ="tblNewPrductConstruction.[TB OD Tol  (+)]"
    Expression ="IIf([Forms]![frmTBMaterial]![TBMaterial]<>\"N/A\",0.05,\"N/A\")"
    Name ="tblNewPrductConstruction.[TB OD Tol  (-)]"
    Expression ="IIf([Forms]![frmTBMaterial]![TBMaterial]<>\"N/A\",0.05,\"N/A\")"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0x8a4321b5d0ac614ba189507969832783
End
Begin
    Begin
        dbText "Name" ="tblNewPrductConstruction.[TB Material]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.[EZ Strip]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewProductConstructionHold.[TB OD Tol  (+)]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.[TB OD Tol  (+)]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewProductConstructionHold.[TB Material]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewProductConstructionHold.[TB OD Tol  (-)]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.[TB OD Tol  (-)]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.[TB Chips Type]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewProductConstructionHold.[TB Chips Type]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewProductConstructionHold.[EZ Strip]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewPrductConstruction.[TB Nominal OD]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblNewProductConstructionHold.[TB Nominal OD]"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =9
    Top =286
    Right =1459
    Bottom =723
    Left =-1
    Top =-1
    Right =1418
    Bottom =172
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="tblNewPrductConstruction"
        Name =""
    End
End

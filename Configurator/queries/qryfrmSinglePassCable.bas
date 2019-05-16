Operation =1
Option =0
Begin InputTables
    Name ="tblCableConstructionReferences"
    Name ="Basic Product Construction"
End
Begin OutputColumns
    Alias ="Oracle"
    Expression ="[Basic Product Construction].[New Oracle Part #]"
    Expression ="[Basic Product Construction].[Item No]"
    Expression ="[Basic Product Construction].Customer"
    Expression ="[Basic Product Construction].[Customer Part#]"
    Expression ="[Basic Product Construction].CustomerRev"
    Expression ="[Basic Product Construction].Active"
    Expression ="[Basic Product Construction].Reason"
    Expression ="[Basic Product Construction].[calculate bill?]"
    Expression ="[Basic Product Construction].Base"
    Expression ="[Basic Product Construction].[Cable Type]"
    Expression ="[Basic Product Construction].[Revision Date]"
    Expression ="[Basic Product Construction].[Revision Letter]"
    Expression ="[Basic Product Construction].[Revision History]"
    Expression ="[Basic Product Construction].[Standard Operation]"
    Expression ="[Basic Product Construction].[Jacket Material]"
    Expression ="[Basic Product Construction].[Jacket Color]"
    Expression ="[Basic Product Construction].[Color Chip ID]"
    Expression ="[Basic Product Construction].[Nominal OD]"
    Expression ="[Basic Product Construction].[OD Tol  (+)]"
    Expression ="[Basic Product Construction].[OD Tol  (-)]"
    Expression ="[Basic Product Construction].Height"
    Expression ="[Basic Product Construction].[Height Tol (+)]"
    Expression ="[Basic Product Construction].[Height Tol (-)]"
    Expression ="[Basic Product Construction].[Nominal Wall]"
    Expression ="[Basic Product Construction].[Max Ave Wall]"
    Expression ="[Basic Product Construction].[Min Ave Wall]"
    Expression ="[Basic Product Construction].[Min Spot Wall]"
    Expression ="[Basic Product Construction].[Print Reel No]"
    Expression ="[Basic Product Construction].[Print Item No]"
    Expression ="[Basic Product Construction].[Std Subunit Print]"
    Expression ="[Basic Product Construction].[Print Type (base)]"
    Expression ="[Basic Product Construction].[Print Line 1]"
    Expression ="[Basic Product Construction].[Print Line 2]"
    Expression ="[Basic Product Construction].[Print Line 3]"
    Expression ="[Basic Product Construction].[Print Line 4]"
    Expression ="[Basic Product Construction].NewPrintLine4"
    Expression ="[Basic Product Construction].[Print Spacing]"
    Expression ="[Basic Product Construction].PrintNotes"
    Expression ="[Basic Product Construction].[Listing Company]"
    Expression ="[Basic Product Construction].[UL-ETL-Listing]"
    Expression ="[Basic Product Construction].[UL-ETL-Const]"
    Expression ="[Basic Product Construction].[UL-ETL-Section]"
    Expression ="[Basic Product Construction].ENumber"
    Expression ="[Basic Product Construction].[Label Type]"
    Expression ="[Basic Product Construction].[Aramid Type]"
    Expression ="[Basic Product Construction].[Number of ends]"
    Expression ="[Basic Product Construction].[Aramid Type2]"
    Expression ="[Basic Product Construction].[Number of ends2]"
    Expression ="[Basic Product Construction].[Unit Chips Type]"
    Expression ="[Basic Product Construction].[Unit Ripcord]"
    Expression ="[Basic Product Construction].[Unit CM]"
    Expression ="[Basic Product Construction].[Lay Length]"
    Expression ="[Basic Product Construction].Talc"
    Expression ="[Basic Product Construction].[EZ Strip]"
    Expression ="[Basic Product Construction].[TB Material]"
    Expression ="[Basic Product Construction].[TB Chips Type]"
    Expression ="[Basic Product Construction].[TB Nominal OD]"
    Expression ="[Basic Product Construction].[TB OD Tol  (+)]"
    Expression ="[Basic Product Construction].[TB OD Tol  (-)]"
    Expression ="[Basic Product Construction].[TB Color Series]"
    Expression ="[Basic Product Construction].[Fiber Type]"
    Expression ="[Basic Product Construction].[Fiber Oracle item]"
    Expression ="[Basic Product Construction].FiberType2"
    Expression ="[Basic Product Construction].FiberType3"
    Expression ="[Basic Product Construction].[1st Req Freq]"
    Expression ="[Basic Product Construction].[1st Max Atten]"
    Expression ="[Basic Product Construction].[1st Min BandW]"
    Expression ="[Basic Product Construction].[2nd Req Freq]"
    Expression ="[Basic Product Construction].[2nd Max Atten]"
    Expression ="[Basic Product Construction].[2nd Min BandW]"
    Expression ="[Basic Product Construction].[SM 1300 Max Atten]"
    Expression ="[Basic Product Construction].[SM 1550 Max Atten]"
    Expression ="[Basic Product Construction].[Special Instr Product1]"
    Expression ="[Basic Product Construction].[zzSpecial Instr Product3]"
    Expression ="[Basic Product Construction].[zzSpecial Instr Product4]"
    Expression ="[Basic Product Construction].[zzOracle Part#]"
    Expression ="[Basic Product Construction].OracleStatus"
    Expression ="[Basic Product Construction].[PSS Document #]"
    Expression ="[Basic Product Construction].DateCreated"
    Expression ="[Basic Product Construction].CreatedBy"
    Expression ="[Basic Product Construction].RevisedBy"
    Expression ="[Basic Product Construction].RevisedDate"
    Expression ="tblCableConstructionReferences.NumSubFillers"
    Alias ="Location"
    Expression ="1"
    Expression ="tblCableConstructionReferences.BaseID"
    Expression ="[Basic Product Construction].Base"
    Expression ="tblCableConstructionReferences.Active"
    Expression ="tblCableConstructionReferences.PrefixID"
    Expression ="tblCableConstructionReferences.ProductID"
    Expression ="[Basic Product Construction].RibbonHighCure"
    Expression ="tblCableConstructionReferences.CableFamily"
    Expression ="tblCableConstructionReferences.FamilyID"
    Expression ="tblCableConstructionReferences.ReleasedDesign"
    Expression ="[Basic Product Construction].[Unit Series]"
    Expression ="[Basic Product Construction].[Print Height]"
    Expression ="[Basic Product Construction].[1st Req Freq - B]"
    Expression ="[Basic Product Construction].[1st Max Atten - B]"
    Expression ="[Basic Product Construction].[1st Min BandW - B]"
    Expression ="[Basic Product Construction].[2nd Req Freq - B]"
    Expression ="[Basic Product Construction].[2nd Max Atten - B]"
    Expression ="[Basic Product Construction].[2nd Min BandW - B]"
    Expression ="[Basic Product Construction].[1st Req Freq - C]"
    Expression ="[Basic Product Construction].[1st Max Atten - C]"
    Expression ="[Basic Product Construction].[1st Min BandW - C]"
    Expression ="[Basic Product Construction].[2nd Req Freq - C]"
    Expression ="[Basic Product Construction].[2nd Max Atten - C]"
    Expression ="[Basic Product Construction].[2nd Min BandW - C]"
    Expression ="[Basic Product Construction].Active"
    Alias ="ItemID"
    Expression ="[Basic Product Construction].ID"
    Expression ="[Basic Product Construction].ItemEngineeringAssist"
    Expression ="[Basic Product Construction].ItemEngineeringAssistReason"
End
Begin Joins
    LeftTable ="tblCableConstructionReferences"
    RightTable ="Basic Product Construction"
    Expression ="tblCableConstructionReferences.Base = [Basic Product Construction].Base"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xbb20474349c5f047b2e85d077e0c4e17
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Oracle"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x58b243473b975141870588ae921c9cbe
        End
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Unit Series]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Print Height]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].Active"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[2nd Min BandW - C]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[1st Req Freq - C]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Location"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x683bfc15cb6f7b4cb892980ac483c6f7
        End
    End
    Begin
        dbText "Name" ="ItemID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x550fcbd29547e345957458360a1fde15
        End
    End
    Begin
        dbText "Name" ="[Basic Product Construction].ItemEngineeringAssist"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].ItemEngineeringAssistReason"
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
    Bottom =384
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="tblCableConstructionReferences"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =632
        Bottom =364
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
End

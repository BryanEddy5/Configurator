Operation =3
Name ="tblNewPrductConstruction"
Option =16
RowCount ="24"
Where ="((([Basic Product Construction].[New Oracle Part #]) Like [Forms]![frmCreateItem"
    "]![strTB] & \"*\" And ([Basic Product Construction].[New Oracle Part #]) Not Lik"
    "e \"*test*\" And ([Basic Product Construction].[New Oracle Part #]) Not Like \"*"
    "-??N*\"))"
Begin InputTables
    Name ="Basic Product Construction"
End
Begin OutputColumns
    Name ="New Oracle Part #"
    Expression ="[Basic Product Construction].[New Oracle Part #]"
    Name ="Item No"
    Expression ="[Basic Product Construction].[Item No]"
    Name ="Customer"
    Expression ="[Basic Product Construction].Customer"
    Name ="Customer Part#"
    Expression ="[Basic Product Construction].[Customer Part#]"
    Name ="CustomerRev"
    Expression ="[Basic Product Construction].CustomerRev"
    Name ="Active"
    Expression ="[Basic Product Construction].Active"
    Name ="Reason"
    Expression ="[Basic Product Construction].Reason"
    Name ="calculate bill?"
    Expression ="[Basic Product Construction].[calculate bill?]"
    Name ="Base"
    Expression ="[Basic Product Construction].Base"
    Name ="Cable Type"
    Expression ="[Basic Product Construction].[Cable Type]"
    Name ="Revision Date"
    Expression ="[Basic Product Construction].[Revision Date]"
    Name ="Revision Letter"
    Expression ="[Basic Product Construction].[Revision Letter]"
    Name ="Revision History"
    Expression ="[Basic Product Construction].[Revision History]"
    Name ="Standard Operation"
    Expression ="[Basic Product Construction].[Standard Operation]"
    Name ="Jacket Material"
    Expression ="[Basic Product Construction].[Jacket Material]"
    Name ="Jacket Color"
    Expression ="[Basic Product Construction].[Jacket Color]"
    Alias ="ColorChip"
    Name ="Color Chip ID"
    Expression ="Null"
    Name ="Nominal OD"
    Expression ="[Basic Product Construction].[Nominal OD]"
    Name ="OD Tol  (+)"
    Expression ="[Basic Product Construction].[OD Tol  (+)]"
    Name ="OD Tol  (-)"
    Expression ="[Basic Product Construction].[OD Tol  (-)]"
    Name ="Height"
    Expression ="[Basic Product Construction].Height"
    Name ="Height Tol (+)"
    Expression ="[Basic Product Construction].[Height Tol (+)]"
    Name ="Height Tol (-)"
    Expression ="[Basic Product Construction].[Height Tol (-)]"
    Name ="Nominal Wall"
    Expression ="[Basic Product Construction].[Nominal Wall]"
    Name ="Max Ave Wall"
    Expression ="[Basic Product Construction].[Max Ave Wall]"
    Name ="Min Ave Wall"
    Expression ="[Basic Product Construction].[Min Ave Wall]"
    Name ="Min Spot Wall"
    Expression ="[Basic Product Construction].[Min Spot Wall]"
    Name ="Print Reel No"
    Expression ="[Basic Product Construction].[Print Reel No]"
    Name ="Print Item No"
    Expression ="[Basic Product Construction].[Print Item No]"
    Name ="Std Subunit Print"
    Expression ="[Basic Product Construction].[Std Subunit Print]"
    Name ="Print Type (base)"
    Expression ="[Basic Product Construction].[Print Type (base)]"
    Name ="Print Line 1"
    Expression ="[Basic Product Construction].[Print Line 1]"
    Name ="Print Line 2"
    Expression ="[Basic Product Construction].[Print Line 2]"
    Name ="Print Line 3"
    Expression ="[Basic Product Construction].[Print Line 3]"
    Name ="Print Line 4"
    Expression ="[Basic Product Construction].[Print Line 4]"
    Name ="NewPrintLine4"
    Expression ="[Basic Product Construction].NewPrintLine4"
    Name ="Print Spacing"
    Expression ="[Basic Product Construction].[Print Spacing]"
    Name ="PrintNotes"
    Expression ="[Basic Product Construction].PrintNotes"
    Name ="Listing Company"
    Expression ="[Basic Product Construction].[Listing Company]"
    Name ="UL-ETL-Listing"
    Expression ="[Basic Product Construction].[UL-ETL-Listing]"
    Name ="UL-ETL-Const"
    Expression ="[Basic Product Construction].[UL-ETL-Const]"
    Name ="UL-ETL-Section"
    Expression ="[Basic Product Construction].[UL-ETL-Section]"
    Name ="ENumber"
    Expression ="[Basic Product Construction].ENumber"
    Name ="Label Type"
    Expression ="[Basic Product Construction].[Label Type]"
    Name ="Aramid Type"
    Expression ="[Basic Product Construction].[Aramid Type]"
    Name ="Number of ends"
    Expression ="[Basic Product Construction].[Number of ends]"
    Name ="Aramid Type2"
    Expression ="[Basic Product Construction].[Aramid Type2]"
    Name ="Number of ends2"
    Expression ="[Basic Product Construction].[Number of ends2]"
    Name ="Unit Chips Type"
    Expression ="[Basic Product Construction].[Unit Chips Type]"
    Name ="Unit Ripcord"
    Expression ="[Basic Product Construction].[Unit Ripcord]"
    Name ="Unit CM"
    Expression ="[Basic Product Construction].[Unit CM]"
    Name ="Lay Length"
    Expression ="[Basic Product Construction].[Lay Length]"
    Name ="Talc"
    Expression ="[Basic Product Construction].Talc"
    Name ="EZ Strip"
    Expression ="[Basic Product Construction].[EZ Strip]"
    Name ="TB Material"
    Expression ="[Basic Product Construction].[TB Material]"
    Name ="TB Chips Type"
    Expression ="[Basic Product Construction].[TB Chips Type]"
    Name ="TB Nominal OD"
    Expression ="[Basic Product Construction].[TB Nominal OD]"
    Name ="TB OD Tol  (+)"
    Expression ="[Basic Product Construction].[TB OD Tol  (+)]"
    Name ="TB OD Tol  (-)"
    Expression ="[Basic Product Construction].[TB OD Tol  (-)]"
    Name ="TB Color Series"
    Expression ="[Basic Product Construction].[TB Color Series]"
    Name ="Fiber Type"
    Expression ="[Basic Product Construction].[Fiber Type]"
    Name ="Fiber Oracle item"
    Expression ="[Basic Product Construction].[Fiber Oracle item]"
    Name ="FiberType2"
    Expression ="[Basic Product Construction].FiberType2"
    Name ="FiberType3"
    Expression ="[Basic Product Construction].FiberType3"
    Name ="1st Req Freq"
    Expression ="[Basic Product Construction].[1st Req Freq]"
    Name ="1st Max Atten"
    Expression ="[Basic Product Construction].[1st Max Atten]"
    Name ="1st Min BandW"
    Expression ="[Basic Product Construction].[1st Min BandW]"
    Name ="2nd Req Freq"
    Expression ="[Basic Product Construction].[2nd Req Freq]"
    Name ="2nd Max Atten"
    Expression ="[Basic Product Construction].[2nd Max Atten]"
    Name ="2nd Min BandW"
    Expression ="[Basic Product Construction].[2nd Min BandW]"
    Name ="SM 1300 Max Atten"
    Expression ="[Basic Product Construction].[SM 1300 Max Atten]"
    Name ="SM 1550 Max Atten"
    Expression ="[Basic Product Construction].[SM 1550 Max Atten]"
    Name ="Special Instr Product1"
    Expression ="[Basic Product Construction].[Special Instr Product1]"
    Name ="zzSpecial Instr Product3"
    Expression ="[Basic Product Construction].[zzSpecial Instr Product3]"
    Name ="zzSpecial Instr Product4"
    Expression ="[Basic Product Construction].[zzSpecial Instr Product4]"
    Name ="zzOracle Part#"
    Expression ="[Basic Product Construction].[zzOracle Part#]"
    Name ="OracleStatus"
    Expression ="[Basic Product Construction].OracleStatus"
    Name ="PSS Document #"
    Expression ="[Basic Product Construction].[PSS Document #]"
End
Begin OrderBy
    Expression ="[Basic Product Construction].[New Oracle Part #]"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0xa578ff49d3154c4ba393d6ca03f5a166
End
Begin
    Begin
        dbText "Name" ="Basic Product Construction.New Oracle Part #"
        dbInteger "ColumnWidth" ="2145"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.CustomerRev"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Customer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Active"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Print Type (base)"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Item No"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Height Tol (+)"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Print Line 1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.NewPrintLine4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Print Line 2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Print Line 3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Print Line 4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Customer Part#"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Reason"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Nominal Wall"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Print Reel No"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Max Ave Wall"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Print Item No"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Min Ave Wall"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Std Subunit Print"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Min Spot Wall"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.calculate bill?"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.RevisedBy"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Revision History"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Jacket Color"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Color Chip ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Base"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Standard Operation"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Cable Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.BaseID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Revision Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Revision Letter"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Jacket Material"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Nominal OD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.OD Tol  (+)"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Height Tol (-)"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.OD Tol  (-)"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Height"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Print Spacing"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.PrintNotes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.zzOracle Part#"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Listing Company"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Unit Chips Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Talc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Fiber Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Lay Length"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Fiber Oracle item"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.UL-ETL-Listing"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.TB Chips Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.PSS Document #"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Unit Ripcord"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.1st Max Atten"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.2nd Min BandW"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Unit CM"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.FiberType2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.1st Min BandW"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.SM 1300 Max Atten"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.TB Material"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.OracleStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.TB Nominal OD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.TB OD Tol  (+)"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.FiberType3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.2nd Req Freq"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.zzSpecial Instr Product4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.1st Req Freq"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.2nd Max Atten"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.UL-ETL-Const"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Aramid Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.DateCreated"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.UL-ETL-Section"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Number of ends"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.EZ Strip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.CreatedBy"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Label Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.TB OD Tol  (-)"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.zzSpecial Instr Product3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.ENumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Aramid Type2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Number of ends2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.TB Color Series"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.SM 1550 Max Atten"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Special Instr Product1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.RevisedDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[New Oracle Part #]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].BaseID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Color Chip ID]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ColorChip"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x00000000000000000000000001000000
        End
    End
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
    Bottom =154
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =38
        Top =6
        Right =427
        Bottom =167
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
End

Operation =1
Option =0
Where ="((([Basic Product Construction].[New Oracle Part #])=\"pa00416-10\"))"
Begin InputTables
    Name ="Basic Product Construction"
End
Begin OutputColumns
    Expression ="[Basic Product Construction].[New Oracle Part #]"
    Expression ="[Basic Product Construction].[Item No]"
    Expression ="[Basic Product Construction].Customer"
    Expression ="[Basic Product Construction].[Customer Part#]"
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
    Alias ="Expr1"
    Expression ="[Basic Product Construction].[\"E\" Number]"
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
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xe92d22fa0ef1f942bc0d6f2be7e0a9d1
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2ac3797b90d4a44b921344d1cbabc7ea
        End
    End
End
Begin
    State =0
    Left =106
    Top =101
    Right =1900
    Bottom =503
    Left =-1
    Top =-1
    Right =1762
    Bottom =171
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =262
        Bottom =169
        Top =0
        Name ="Basic Product Construction"
        Name =""
    End
End

Operation =1
Option =18
RowCount ="1"
Where ="((([Basic Product Construction_1].[Fiber Oracle item])=Forms!frmCreateItem!Fiber"
    ") And (([Basic Product Construction_1].[New Oracle Part #]) Not Like \"*test*\")"
    " And (([Basic Product Construction_1].Active)<>0) And ((Mid([Basic Product Const"
    "ruction_1]![Item No],10,1))=Forms!frmCreateItem!SubColor) And ((Mid([Basic Produ"
    "ct Construction_1]![Item No],11,2))=Forms!frmCreateItem!TBColor & Forms!frmCreat"
    "eItem!TBType) And (([Product Desrcriptions].Product)=Forms!frmCreateItem!Base An"
    "d ([Product Desrcriptions].Product)=Forms!frmCreateItem!Base))"
Begin InputTables
    Name ="Product Desrcriptions"
    Name ="Basic Product Construction"
    Alias ="Basic Product Construction_1"
End
Begin OutputColumns
    Expression ="[Basic Product Construction_1].[Fiber Oracle item]"
    Expression ="[Basic Product Construction_1].OracleStatus"
    Expression ="[Basic Product Construction_1].[New Oracle Part #]"
    Expression ="[Basic Product Construction_1].Active"
    Expression ="[Basic Product Construction_1].[Item No]"
    Expression ="[Basic Product Construction_1].[Jacket Material]"
    Alias ="Expr1"
    Expression ="Mid([Basic Product Construction_1]![Item No],10,1)"
    Alias ="TB_Colr"
    Expression ="Mid([Basic Product Construction_1]![Item No],11,2)"
    Expression ="[Product Desrcriptions].Product"
    Expression ="[Product Desrcriptions].Product"
End
Begin Joins
    LeftTable ="Product Desrcriptions"
    RightTable ="Basic Product Construction_1"
    Expression ="[Product Desrcriptions].[Level 1 Product] = [Basic Product Construction_1].Base"
    Flag =1
End
Begin OrderBy
    Expression ="[Basic Product Construction_1].[Fiber Oracle item]"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xbac3e638c629644296d715daacef8904
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Basic Product Construction_1.New Oracle Part #"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.Item No"
        dbInteger "ColumnOrder" ="78"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.Active"
        dbInteger "ColumnWidth" ="2325"
        dbInteger "ColumnOrder" ="4"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.Fiber Oracle item"
        dbInteger "ColumnOrder" ="60"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.OracleStatus"
        dbInteger "ColumnOrder" ="73"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction_1].[Fiber Oracle item]"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction_1].[Item No]"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction_1].OracleStatus"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction_1].[New Oracle Part #]"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb5362fd73e908e44bec3e699e8e046e0
        End
    End
    Begin
        dbText "Name" ="TB_Colr"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd9e24f0bc7f27e4db025c565f45129fe
        End
    End
    Begin
        dbText "Name" ="[Product Desrcriptions].Product"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction_1].[Jacket Material]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1008"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.Jacket Material"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction_1].Active"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.Height Tol (-)"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.CustomerRev"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.Revision Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.Reason"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.Revision History"
        dbInteger "ColumnWidth" ="17775"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.Customer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.NewPrintLine4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.Revision Letter"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.Base"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.UL-ETL-Listing"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.Height"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.Print Line 3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.Height Tol (+)"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.calculate bill?"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.Jacket Color"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.Nominal OD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.PrintNotes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.Print Line 2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.Cable Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.Standard Operation"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.OD Tol  (+)"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.OD Tol  (-)"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.Aramid Type2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.Customer Part#"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.1st Min BandW"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.UL-ETL-Const"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.zzSpecial Instr Product4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.Print Item No"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.Number of ends"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.Unit Chips Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.Color Chip ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.Number of ends2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.Print Spacing"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.TB OD Tol  (+)"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.Nominal Wall"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.SM 1300 Max Atten"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.UL-ETL-Section"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.1st Req Freq"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.Std Subunit Print"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.EZ Strip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.Talc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.Print Line 4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.TB OD Tol  (-)"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.Max Ave Wall"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.FiberType3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.Min Spot Wall"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.Label Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.Print Type (base)"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.TB Color Series"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.Fiber Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.Listing Company"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.FiberType2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.Min Ave Wall"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.2nd Req Freq"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.Print Reel No"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.Aramid Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.ENumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.Print Line 1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.Lay Length"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.Unit Ripcord"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.Unit CM"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.TB Material"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.TB Chips Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.TB Nominal OD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.2nd Max Atten"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.1st Max Atten"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.2nd Min BandW"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.SM 1550 Max Atten"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.Special Instr Product1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.zzSpecial Instr Product3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.zzOracle Part#"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction_1.PSS Document #"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =76
    Top =434
    Right =1185
    Bottom =1063
    Left =-1
    Top =-1
    Right =1077
    Bottom =225
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =1106
        Top =44
        Right =1335
        Bottom =252
        Top =0
        Name ="Product Desrcriptions"
        Name =""
    End
    Begin
        Left =464
        Top =9
        Right =751
        Bottom =277
        Top =0
        Name ="Basic Product Construction_1"
        Name =""
    End
End

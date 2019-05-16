Operation =3
Name ="tblNewPrductConstruction"
Option =2
Where ="(((tblCableConstructionReferences_1.Base)=[Forms]![frmCreateItem]![cbSubunit]))"
Begin InputTables
    Name ="tblCableConstructionReferences"
    Alias ="tblCableConstructionReferences_1"
    Name ="tblCableConstructions"
    Alias ="tblCableConstructions_1"
End
Begin OutputColumns
    Name ="Jacket Material"
    Expression ="tblCableConstructions_1.JacketMaterial"
    Name ="Base"
    Expression ="tblCableConstructionReferences_1.Base"
End
Begin Joins
    LeftTable ="tblCableConstructions_1"
    RightTable ="tblCableConstructionReferences_1"
    Expression ="tblCableConstructions_1.BaseID = tblCableConstructionReferences_1.BaseID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x1354cbbc2bab6c4f91fdeff4d13b90af
End
Begin
    Begin
        dbText "Name" ="Basic Product Construction.New Oracle Part #"
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
        dbText "Name" ="Basic Product Construction.Cable Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Print Line 2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Revision Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Print Line 3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Revision Letter"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Print Line 4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Jacket Material"
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
        dbText "Name" ="Basic Product Construction.Customer Part#"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Reason"
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
        dbText "Name" ="Basic Product Construction.Revision History"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Jacket Color"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.OD Tol  (-)"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Color Chip ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Height"
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
        dbText "Name" ="Basic Product Construction.Print Spacing"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.PrintNotes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Listing Company"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.UL-ETL-Listing"
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
        dbText "Name" ="Basic Product Construction.UL-ETL-Section"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Number of ends"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Label Type"
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
        dbText "Name" ="Basic Product Construction.Unit Chips Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Talc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Lay Length"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Unit Ripcord"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Unit CM"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.EZ Strip"
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
        dbText "Name" ="Basic Product Construction.TB Material"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.OracleStatus"
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
        dbText "Name" ="Basic Product Construction.TB Nominal OD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.TB OD Tol  (+)"
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
        dbText "Name" ="Basic Product Construction.TB Color Series"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Fiber Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Fiber Oracle item"
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
        dbText "Name" ="Basic Product Construction.SM 1550 Max Atten"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.Special Instr Product1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Basic Product Construction.zzOracle Part#"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Product Desrcriptions].Product"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Item No]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].Active"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.Product"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences.Base"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Product Desrcriptions].[Level 1 Product]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Basic Product Construction].[Jacket Color]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions.JacketMaterial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructions_1.JacketMaterial"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2745"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="[Product Desrcriptions].SubFiller"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCableConstructionReferences_1.Base"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =52
    Top =356
    Right =1161
    Bottom =978
    Left =-1
    Top =-1
    Right =1077
    Bottom =276
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =594
        Top =26
        Right =832
        Bottom =319
        Top =0
        Name ="tblCableConstructionReferences_1"
        Name =""
    End
    Begin
        Left =69
        Top =102
        Right =315
        Bottom =284
        Top =0
        Name ="tblCableConstructions_1"
        Name =""
    End
End

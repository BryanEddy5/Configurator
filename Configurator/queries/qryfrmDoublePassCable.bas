dbMemo "SQL" ="SELECT Left([Revision Date],10) AS RevDate, [Basic (Cable/Sheath)].[New Oracle P"
    "art Number] AS Oracle, [Basic (Cable/Sheath)].[Item No], [Basic (Cable/Sheath)]."
    "Customer, [Basic (Cable/Sheath)].[Customer Part Num], [Basic (Cable/Sheath)].Cus"
    "tomerRev, [Basic (Cable/Sheath)].Active, [Basic (Cable/Sheath)].Reason, [Basic ("
    "Cable/Sheath)].[Revision Date], [Basic (Cable/Sheath)].[Revision Letter], [Basic"
    " (Cable/Sheath)].[Revision History], [Basic (Cable/Sheath)].[Print Item No], [Ba"
    "sic (Cable/Sheath)].[Print Reel No], [Basic (Cable/Sheath)].[Print Height], [Bas"
    "ic (Cable/Sheath)].[Print Type (base)], [Basic (Cable/Sheath)].[Print Line 1], ["
    "Basic (Cable/Sheath)].[Print Line 2], [Basic (Cable/Sheath)].[Print Line 3], [Ba"
    "sic (Cable/Sheath)].[Print Line 4], [Basic (Cable/Sheath)].NewPrintLine4, [Basic"
    " (Cable/Sheath)].PrintSpacing, [Basic (Cable/Sheath)].PrintNotes, [Basic (Cable/"
    "Sheath)].[Listing Company], [Basic (Cable/Sheath)].[UL-ETL-Listing], [Basic (Cab"
    "le/Sheath)].[UL-ETL-Const], [Basic (Cable/Sheath)].[UL-ETL-Section], [Basic (Cab"
    "le/Sheath)].[Unit Series], [Basic (Cable/Sheath)].[Jacket Color], [Basic (Cable/"
    "Sheath)].[Fiber Type A], [Basic (Cable/Sheath)].[1st Req Freq - A], [Basic (Cabl"
    "e/Sheath)].[1st Max Atten - A], [Basic (Cable/Sheath)].[1st Min BandW - A], [Bas"
    "ic (Cable/Sheath)].[2nd Req Freq - A], [Basic (Cable/Sheath)].[2nd Max Atten - A"
    "], [Basic (Cable/Sheath)].[2nd Min BandW - A], [Basic (Cable/Sheath)].[Fiber Typ"
    "e B], [Basic (Cable/Sheath)].[1st Req Freq - B], [Basic (Cable/Sheath)].[1st Max"
    " Atten - B], [Basic (Cable/Sheath)].[1st Min BandW - B], [Basic (Cable/Sheath)]."
    "[2nd Req Freq - B], [Basic (Cable/Sheath)].[2nd Max Atten - B], [Basic (Cable/Sh"
    "eath)].[2nd Min BandW - B], [Basic (Cable/Sheath)].[Fiber Type C], [Basic (Cable"
    "/Sheath)].[1st Req Freq - C], [Basic (Cable/Sheath)].[1st Max Atten - C], [Basic"
    " (Cable/Sheath)].[1st Min BandW - C], [Basic (Cable/Sheath)].[2nd Req Freq - C],"
    " [Basic (Cable/Sheath)].[2nd Max Atten - C], [Basic (Cable/Sheath)].[2nd Min Ban"
    "dW - C], [Basic (Cable/Sheath)].[Special Instr Product1], [Basic (Cable/Sheath)]"
    ".[Special Instr Product2], [Basic (Cable/Sheath)].[Special Instr Product3], [Bas"
    "ic (Cable/Sheath)].[Special Instr Product4], [Basic (Cable/Sheath)].[Label Type]"
    ", [Basic (Cable/Sheath)].ENumber, [Basic (Cable/Sheath)].[load into oracle 11?],"
    " [Basic (Cable/Sheath)].[calculate bill?], [Basic (Cable/Sheath)].[Standard Oper"
    "ation], [Basic (Cable/Sheath)].OracleStatus, [Basic (Cable/Sheath)].ColorChip, ["
    "Basic (Cable/Sheath)].ENumber, tblCableConstructionReferences.BaseID, tblCableCo"
    "nstructionReferences.Active, tblCableConstructionReferences.PrefixID, tblCableCo"
    "nstructionReferences.ProductID, [Basic (Cable/Sheath)].Base, tblCableConstructio"
    "nReferences.FamilyID, tblCableConstructionReferences.ReleasedDesign, tblCableCon"
    "structionReferences.DateCreated\015\012FROM [Basic (Cable/Sheath)] LEFT JOIN tbl"
    "CableConstructionReferences ON [Basic (Cable/Sheath)].Base = tblCableConstructio"
    "nReferences.Base;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x5b1558ba5497604880c912877e96222f
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="RevDate"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6a485dff8dca5949ac807a4ef717a480
        End
    End
    Begin
        dbText "Name" ="Oracle"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x79d848945d4c9e488abe70bb011d0450
        End
    End
End

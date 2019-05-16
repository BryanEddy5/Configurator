dbMemo "SQL" ="SELECT qryFiberOracle.Oracle AS ItemNumber, qryFiberOracle.Oracle AS SpecName, \""
    "E NUMBER\" AS SpecificationElement, \"\" AS UOM, [Enumber] AS TargetValue\015\012"
    "FROM qryFiberOracle LEFT JOIN qrySpecwithLabels ON qryFiberOracle.Oracle=qrySpec"
    "withLabels.ItemNumber\015\012WHERE (((qryFiberOracle.Oracle) Not Like \"*-00*\")"
    " AND ((qrySpecwithLabels.ItemNumber) Is Null) AND ((qryFiberOracle.ListingCompan"
    "y)=\"etl\" Or (qryFiberOracle.ListingCompany)=\"ul\"));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x033a1cf2e75f7d498b3382d513945e39
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="ItemNumber"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x0e8f9f2e865e184181f69bbde08f6d1b
        End
    End
    Begin
        dbText "Name" ="SpecName"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x1ab1881bdd6db84c9280c7cfc236f87e
        End
    End
    Begin
        dbText "Name" ="SpecificationElement"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x00ccd206d4772a45a516faa234cb4928
        End
    End
    Begin
        dbText "Name" ="UOM"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2c3487561e692c40882e6e2d0f5e5be2
        End
    End
    Begin
        dbText "Name" ="TargetValue"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6f9cf2e85994b1498b290429ceadc376
        End
    End
End

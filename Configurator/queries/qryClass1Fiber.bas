Operation =1
Option =0
Where ="(((dbo_fi_fiber_length_extract_AFLPRD.[Item Number]) Like \"fbr00396*\" Or (dbo_"
    "fi_fiber_length_extract_AFLPRD.[Item Number]) Like \"fbr00393*\") AND ((IIf([GEO"
    "_CORE_AVG_T] Is Null,0,CDbl([GEO_CORE_AVG_T])))>=49 And (IIf([GEO_CORE_AVG_T] Is"
    " Null,0,CDbl([GEO_CORE_AVG_T])))<=51) AND ((IIf([GEO_CORE_AVG_B] Is Null,0,CDbl("
    "[GEO_CORE_AVG_B])))>=49 And (IIf([GEO_CORE_AVG_B] Is Null,0,CDbl([GEO_CORE_AVG_B"
    "])))<=51) AND ((IIf([GEO_CLAD_CONC_B] Is Null,0,CDbl([GEO_CLAD_CONC_B])))<=1) AN"
    "D ((IIf([GEO_CLAD_CONC_T] Is Null,0,CDbl([GEO_CLAD_CONC_T])))<=1) AND ((IIf([GEO"
    "_CLAD_AVG_B] Is Null,0,CDbl([GEO_CLAD_AVG_B])))>=124.6 And (IIf([GEO_CLAD_AVG_B]"
    " Is Null,0,CDbl([GEO_CLAD_AVG_B])))<=125.15) AND ((IIf([CONC_2T] Is Null,0,CDbl("
    "[CONC_2T])))<=10) AND ((IIf([CONC_2B] Is Null,0,CDbl([CONC_2B])))<=10))"
Begin InputTables
    Name ="dbo_fi_fiber_length_extract_AFLPRD"
End
Begin OutputColumns
    Expression ="dbo_fi_fiber_length_extract_AFLPRD.*"
    Expression ="dbo_fi_fiber_length_extract_AFLPRD.[Item Number]"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x26d89d9a443a1b46a4e2f06a6d5497af
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbMemo "Filter" ="([dbo_fi_fiber_length_extract_AFLPRD].[Item Number] Like \"*FBR00393*\")"
dbMemo "OrderBy" ="[qryClass1Fiber].[Locator], [dbo_fi_fiber_length_extract_AFLPRD].[Item Number]"
Begin
    Begin
        dbText "Name" ="dbo_fi_fiber_length_extract_AFLPRD.[Item Number]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_fi_fiber_length_extract_AFLPRD.Org"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_fi_fiber_length_extract_AFLPRD.Item Number"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x55724479b3da5d42a13a60676e5ebfaa
        End
    End
    Begin
        dbText "Name" ="dbo_fi_fiber_length_extract_AFLPRD.Lot Number"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_fi_fiber_length_extract_AFLPRD.Locator"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_fi_fiber_length_extract_AFLPRD.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_fi_fiber_length_extract_AFLPRD.Onhand Qty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_fi_fiber_length_extract_AFLPRD.Material Cost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_fi_fiber_length_extract_AFLPRD.Resource Cost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_fi_fiber_length_extract_AFLPRD.Overhead Cost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_fi_fiber_length_extract_AFLPRD.Last Transacted"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_fi_fiber_length_extract_AFLPRD.Subinventory"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_fi_fiber_length_extract_AFLPRD.1310 ATTN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_fi_fiber_length_extract_AFLPRD.850 BAND"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_fi_fiber_length_extract_AFLPRD.Manufacturer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_fi_fiber_length_extract_AFLPRD.Family"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_fi_fiber_length_extract_AFLPRD.Vendor"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2535"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="dbo_fi_fiber_length_extract_AFLPRD.1550 ATTN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_fi_fiber_length_extract_AFLPRD.1300 BAND"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_fi_fiber_length_extract_AFLPRD.Coating Diameter"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_fi_fiber_length_extract_AFLPRD.Name"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1665"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="dbo_fi_fiber_length_extract_AFLPRD.Data File ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_fi_fiber_length_extract_AFLPRD.Color"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_fi_fiber_length_extract_AFLPRD.850 ATTN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_fi_fiber_length_extract_AFLPRD.MFD 1550"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_fi_fiber_length_extract_AFLPRD.Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_fi_fiber_length_extract_AFLPRD.1300 ATTN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_fi_fiber_length_extract_AFLPRD.Extract Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_fi_fiber_length_extract_AFLPRD.Cable Cutoff"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_fi_fiber_length_extract_AFLPRD.Clad Diameter"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_fi_fiber_length_extract_AFLPRD.Clad Non-Circularity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_fi_fiber_length_extract_AFLPRD.Core clad Concentricity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_fi_fiber_length_extract_AFLPRD.Original Length"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_fi_fiber_length_extract_AFLPRD.POReceiptDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_fi_fiber_length_extract_AFLPRD.WIPIssueDT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_fi_fiber_length_extract_AFLPRD.P O  Number"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_fi_fiber_length_extract_AFLPRD.Box Number"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_fi_fiber_length_extract_AFLPRD.Vendor Ship Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_fi_fiber_length_extract_AFLPRD.GEO_CORE_AVG_T"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_fi_fiber_length_extract_AFLPRD.GEO_CLAD_AVG_T"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_fi_fiber_length_extract_AFLPRD.Consigned"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_fi_fiber_length_extract_AFLPRD.Fiber Cutoff"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_fi_fiber_length_extract_AFLPRD.MFD 1310nm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_fi_fiber_length_extract_AFLPRD.Numerical Aperture"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_fi_fiber_length_extract_AFLPRD.GEO_CLAD_CONC_B"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_fi_fiber_length_extract_AFLPRD.Core Diameter"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_fi_fiber_length_extract_AFLPRD.MAC Value"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_fi_fiber_length_extract_AFLPRD.1625 ATTN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_fi_fiber_length_extract_AFLPRD.Original Item Received As"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_fi_fiber_length_extract_AFLPRD.GEO_CORE_AVG_B"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_fi_fiber_length_extract_AFLPRD.GEO_CLAD_AVG_B"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_fi_fiber_length_extract_AFLPRD.CONC_2B"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_fi_fiber_length_extract_AFLPRD.GEO_CLAD_CONC_T"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_fi_fiber_length_extract_AFLPRD.CONC_2T"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_fi_fiber_length_extract_AFLPRD.Number"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Org"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7b3df9872af7f641a1d30094b6ef673e
        End
    End
    Begin
        dbText "Name" ="Manufacturer"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x0594e514916a564b93a14c5f171775a2
        End
    End
    Begin
        dbText "Name" ="Family"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x73d6994bc60a3440a3fc973b5c1c8d26
        End
    End
    Begin
        dbText "Name" ="Name"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3a452811df41694b996a3b334e77e682
        End
    End
    Begin
        dbText "Name" ="Lot Number"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x77c5d22817d05044b847db06817b6f14
        End
    End
    Begin
        dbText "Name" ="Locator"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2dcab44f614f4f44a475ed64b65ae17e
        End
    End
    Begin
        dbText "Name" ="Vendor"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x172fdbe289554541ba7f8398a6c47b14
        End
    End
    Begin
        dbText "Name" ="Data File ID"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xfede61693160af4fb12cd842ea255647
        End
    End
    Begin
        dbText "Name" ="Color"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa9e9f8ba4151104cba6a03a1995494a6
        End
    End
    Begin
        dbText "Name" ="Description"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x18c90b697bf7da47b2528f3dc33ec182
        End
    End
    Begin
        dbText "Name" ="Onhand Qty"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7d172bb85cc7774c8228c79bb5500ea6
        End
    End
    Begin
        dbText "Name" ="Material Cost"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd076ee78f147534c8329085ce2386f76
        End
    End
    Begin
        dbText "Name" ="Resource Cost"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7d6adf093fa6774ba3187e1631f22945
        End
    End
    Begin
        dbText "Name" ="Overhead Cost"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x18a06e7abf28cf499cfd7492577b0457
        End
    End
    Begin
        dbText "Name" ="Last Transacted"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa56f505106c1a34f91ee6c43f29e836a
        End
    End
    Begin
        dbText "Name" ="Status"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe2cd62e7839119428c513ae0ffc3908f
        End
    End
    Begin
        dbText "Name" ="Subinventory"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4d64f3232654784a9f82cc9b2a081802
        End
    End
    Begin
        dbText "Name" ="1310 ATTN"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa2be8bfbeaf61a46a7ab09659e513280
        End
    End
    Begin
        dbText "Name" ="1550 ATTN"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf33b7f0d193425439bd63b9be613c621
        End
    End
    Begin
        dbText "Name" ="850 ATTN"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x55af12e93915614287cbd55166d40b1e
        End
    End
    Begin
        dbText "Name" ="1300 ATTN"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x8b34c72cc42a644282eacae51aed2db4
        End
    End
    Begin
        dbText "Name" ="850 BAND"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2ac083886d9c964d9b36755308d3ba62
        End
    End
    Begin
        dbText "Name" ="1300 BAND"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4f39708fd07ab54aa4d3108a607ffa98
        End
    End
    Begin
        dbText "Name" ="MFD 1550"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x21b5d498bd5fce4b805605a6778e1ff1
        End
    End
    Begin
        dbText "Name" ="Extract Date"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x1d666e30c1e7c949a89fe67a6326b548
        End
    End
    Begin
        dbText "Name" ="Cable Cutoff"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x358b01c8f5c79446bdc99586658aafb4
        End
    End
    Begin
        dbText "Name" ="Clad Diameter"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x5e1103be35633743863fdb205b9d1edd
        End
    End
    Begin
        dbText "Name" ="Clad Non-Circularity"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb421475cbb59e2488aed75371e605b8a
        End
    End
    Begin
        dbText "Name" ="Coating Diameter"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x1f8d4f835b59cd4dabf546077eb088fe
        End
    End
    Begin
        dbText "Name" ="Core clad Concentricity"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb2572d8e9c55ce499de196754c823878
        End
    End
    Begin
        dbText "Name" ="Fiber Cutoff"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x29d59a18544a404bb99390577fbfa78d
        End
    End
    Begin
        dbText "Name" ="MFD 1310nm"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x40d54f0f16c06746af4a64ed1600ce4d
        End
    End
    Begin
        dbText "Name" ="Numerical Aperture"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3d39e8e834764543a75003bdae223483
        End
    End
    Begin
        dbText "Name" ="Core Diameter"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x59e6afb2fb7fd249b20f9bc402d0b0f3
        End
    End
    Begin
        dbText "Name" ="MAC Value"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xac84855875f14d43baed3390f9270748
        End
    End
    Begin
        dbText "Name" ="Original Length"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf5626f31ea8add4998ce2a65f3d9576b
        End
    End
    Begin
        dbText "Name" ="1625 ATTN"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe6817f6ea2640142a7ce3021fdfa94ce
        End
    End
    Begin
        dbText "Name" ="POReceiptDt"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe8d21f4e447a3244bbb37769f3206ea4
        End
    End
    Begin
        dbText "Name" ="WIPIssueDT"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x89f7bcf615702a4189da882e93f6c8a9
        End
    End
    Begin
        dbText "Name" ="P O  Number"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7495495dfac5d14b9e61b13b877b192a
        End
    End
    Begin
        dbText "Name" ="Box Number"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x672f1c057e3a9d43a4a23692cd036738
        End
    End
    Begin
        dbText "Name" ="Vendor Ship Date"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x629fbc8e28534247bd5e0bd9a0be2344
        End
    End
    Begin
        dbText "Name" ="Original Item Received As"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x65afdf4a2ac52c428401a3c801bc83e1
        End
    End
    Begin
        dbText "Name" ="GEO_CORE_AVG_T"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x92cb0bf7085afa4a9c9c229d7634dfe6
        End
    End
    Begin
        dbText "Name" ="GEO_CORE_AVG_B"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3dbb98007da46f4ba0b186cd94852499
        End
    End
    Begin
        dbText "Name" ="GEO_CLAD_CONC_T"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4d7ff1d2a4f71d44af6c21ad431417a0
        End
    End
    Begin
        dbText "Name" ="GEO_CLAD_CONC_B"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xfa0e295b63499641b8835da8b2d2cab0
        End
    End
    Begin
        dbText "Name" ="GEO_CLAD_AVG_T"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x1efcca8be742ff4cb65a196211ae945e
        End
    End
    Begin
        dbText "Name" ="GEO_CLAD_AVG_B"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4e3b54dde0540448b98164623c83ce58
        End
    End
    Begin
        dbText "Name" ="CONC_2T"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x2ddb1d8a16f2cc40a87a6c768886536d
        End
    End
    Begin
        dbText "Name" ="CONC_2B"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x99a4dfe52847d2479edee5801987f67b
        End
    End
    Begin
        dbText "Name" ="Consigned"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x8e24a805a75ab448b939fd328f094314
        End
    End
    Begin
        dbText "Name" ="Number"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x9047c60756136746ad16d8e6148e92fa
        End
    End
    Begin
        dbText "Name" ="DateInserted"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xca4bd99b8d76ae4ea29bb513a8aa370c
        End
    End
    Begin
        dbText "Name" ="Pallet_Number"
        dbBinary "GUID" = Begin
            0x9dc2650f254cbd43b18c93c5bae70795
        End
    End
End
Begin
    State =0
    Left =43
    Top =91
    Right =1603
    Bottom =967
    Left =-1
    Top =-1
    Right =1528
    Bottom =145
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =285
        Bottom =197
        Top =0
        Name ="dbo_fi_fiber_length_extract_AFLPRD"
        Name =""
    End
End

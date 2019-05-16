Operation =3
Name ="BasicProductConstruction_LIVE"
Option =0
Begin InputTables
    Name ="tabl_FiberConversion_NewPremisData"
End
Begin OutputColumns
    Alias ="Expr1"
    Name ="New Oracle Part #"
    Expression ="tabl_FiberConversion_NewPremisData.NewOracle"
    Alias ="Expr2"
    Name ="Item No"
    Expression ="tabl_FiberConversion_NewPremisData.[Item No]"
    Alias ="Expr3"
    Name ="Customer"
    Expression ="tabl_FiberConversion_NewPremisData.Customer"
    Alias ="Expr4"
    Name ="Customer Part#"
    Expression ="tabl_FiberConversion_NewPremisData.[Customer Part#]"
    Alias ="Expr5"
    Name ="Active"
    Expression ="tabl_FiberConversion_NewPremisData.Active"
    Alias ="Expr6"
    Name ="Reason"
    Expression ="tabl_FiberConversion_NewPremisData.Reason"
    Alias ="Expr7"
    Name ="calculate bill?"
    Expression ="tabl_FiberConversion_NewPremisData.[calculate bill?]"
    Alias ="Expr8"
    Name ="Base"
    Expression ="tabl_FiberConversion_NewPremisData.Base"
    Alias ="Expr9"
    Name ="Cable Type"
    Expression ="tabl_FiberConversion_NewPremisData.[Cable Type]"
    Alias ="Expr10"
    Name ="Revision Date"
    Expression ="tabl_FiberConversion_NewPremisData.[Revision Date]"
    Alias ="Expr11"
    Name ="Revision Letter"
    Expression ="tabl_FiberConversion_NewPremisData.[Revision Letter]"
    Alias ="Expr12"
    Name ="Revision History"
    Expression ="tabl_FiberConversion_NewPremisData.[Revision History]"
    Alias ="Expr13"
    Name ="Jacket Material"
    Expression ="tabl_FiberConversion_NewPremisData.[Jacket Material]"
    Alias ="Expr14"
    Name ="Standard Operation"
    Expression ="tabl_FiberConversion_NewPremisData.[Standard Operation]"
    Alias ="Expr15"
    Name ="Jacket Color"
    Expression ="tabl_FiberConversion_NewPremisData.[Jacket Color]"
    Alias ="Expr16"
    Name ="Color Chip ID"
    Expression ="tabl_FiberConversion_NewPremisData.[Color Chip ID]"
    Alias ="Expr17"
    Name ="Nominal OD"
    Expression ="tabl_FiberConversion_NewPremisData.[Nominal OD]"
    Alias ="Expr18"
    Name ="OD Tol  (+)"
    Expression ="tabl_FiberConversion_NewPremisData.[OD Tol  (+)]"
    Alias ="Expr19"
    Name ="OD Tol  (-)"
    Expression ="tabl_FiberConversion_NewPremisData.[OD Tol  (-)]"
    Alias ="Expr20"
    Name ="Height"
    Expression ="tabl_FiberConversion_NewPremisData.Height"
    Alias ="Expr21"
    Name ="Height Tol (+)"
    Expression ="tabl_FiberConversion_NewPremisData.[Height Tol (+)]"
    Alias ="Expr22"
    Name ="Height Tol (-)"
    Expression ="tabl_FiberConversion_NewPremisData.[Height Tol (-)]"
    Alias ="Expr23"
    Name ="Nominal Wall"
    Expression ="tabl_FiberConversion_NewPremisData.[Nominal Wall]"
    Alias ="Expr24"
    Name ="Max Ave Wall"
    Expression ="tabl_FiberConversion_NewPremisData.[Max Ave Wall]"
    Alias ="Expr25"
    Name ="Min Ave Wall"
    Expression ="tabl_FiberConversion_NewPremisData.[Min Ave Wall]"
    Alias ="Expr26"
    Name ="Min Spot Wall"
    Expression ="tabl_FiberConversion_NewPremisData.[Min Spot Wall]"
    Alias ="Expr27"
    Name ="Print Reel No"
    Expression ="tabl_FiberConversion_NewPremisData.[Print Reel No]"
    Alias ="Expr28"
    Name ="Print Item No"
    Expression ="tabl_FiberConversion_NewPremisData.[Print Item No]"
    Alias ="Expr29"
    Name ="Std Subunit Print"
    Expression ="tabl_FiberConversion_NewPremisData.[Std Subunit Print]"
    Alias ="Expr30"
    Name ="Print Type (base)"
    Expression ="tabl_FiberConversion_NewPremisData.[Print Type (base)]"
    Alias ="Expr31"
    Name ="Print Line 1"
    Expression ="tabl_FiberConversion_NewPremisData.[Print Line 1]"
    Alias ="Expr32"
    Name ="Print Line 2"
    Expression ="tabl_FiberConversion_NewPremisData.[Print Line 2]"
    Alias ="Expr33"
    Name ="Print Line 3"
    Expression ="tabl_FiberConversion_NewPremisData.[Print Line 3]"
    Alias ="Expr34"
    Name ="Print Line 4"
    Expression ="tabl_FiberConversion_NewPremisData.[Print Line 4]"
    Alias ="Expr35"
    Name ="NewPrintLine4"
    Expression ="tabl_FiberConversion_NewPremisData.NewPrintLine4"
    Alias ="Expr36"
    Name ="Print Spacing"
    Expression ="tabl_FiberConversion_NewPremisData.[Print Spacing]"
    Alias ="Expr37"
    Name ="PrintNotes"
    Expression ="tabl_FiberConversion_NewPremisData.PrintNotes"
    Alias ="Expr38"
    Name ="Listing Company"
    Expression ="tabl_FiberConversion_NewPremisData.[Listing Company]"
    Alias ="Expr39"
    Name ="UL-ETL-Listing"
    Expression ="tabl_FiberConversion_NewPremisData.[UL-ETL-Listing]"
    Alias ="Expr40"
    Name ="UL-ETL-Const"
    Expression ="tabl_FiberConversion_NewPremisData.[UL-ETL-Const]"
    Alias ="Expr41"
    Name ="UL-ETL-Section"
    Expression ="tabl_FiberConversion_NewPremisData.[UL-ETL-Section]"
    Alias ="Expr42"
    Name ="\"E\" Number"
    Expression ="tabl_FiberConversion_NewPremisData.[\"E\" Number]"
    Alias ="Expr43"
    Name ="Label Type"
    Expression ="tabl_FiberConversion_NewPremisData.[Label Type]"
    Alias ="Expr44"
    Name ="Aramid Type"
    Expression ="tabl_FiberConversion_NewPremisData.[Aramid Type]"
    Alias ="Expr45"
    Name ="Number of ends"
    Expression ="tabl_FiberConversion_NewPremisData.[Number of ends]"
    Alias ="Expr46"
    Name ="Aramid Type2"
    Expression ="tabl_FiberConversion_NewPremisData.[Aramid Type2]"
    Alias ="Expr47"
    Name ="Unit Chips Type"
    Expression ="tabl_FiberConversion_NewPremisData.[Unit Chips Type]"
    Alias ="Expr48"
    Name ="Number of ends2"
    Expression ="tabl_FiberConversion_NewPremisData.[Number of ends2]"
    Alias ="Expr49"
    Name ="Unit Ripcord"
    Expression ="tabl_FiberConversion_NewPremisData.[Unit Ripcord]"
    Alias ="Expr50"
    Name ="Unit CM"
    Expression ="tabl_FiberConversion_NewPremisData.[Unit CM]"
    Alias ="Expr51"
    Name ="Lay Length"
    Expression ="tabl_FiberConversion_NewPremisData.[Lay Length]"
    Alias ="Expr52"
    Name ="Talc"
    Expression ="tabl_FiberConversion_NewPremisData.Talc"
    Alias ="Expr53"
    Name ="EZ Strip"
    Expression ="tabl_FiberConversion_NewPremisData.[EZ Strip]"
    Alias ="Expr54"
    Name ="TB Material"
    Expression ="tabl_FiberConversion_NewPremisData.[TB Material]"
    Alias ="Expr55"
    Name ="TB Chips Type"
    Expression ="tabl_FiberConversion_NewPremisData.[TB Chips Type]"
    Alias ="Expr56"
    Name ="TB OD Tol  (+)"
    Expression ="tabl_FiberConversion_NewPremisData.[TB OD Tol  (+)]"
    Alias ="Expr57"
    Name ="TB Nominal OD"
    Expression ="tabl_FiberConversion_NewPremisData.[TB Nominal OD]"
    Alias ="Expr58"
    Name ="TB OD Tol  (-)"
    Expression ="tabl_FiberConversion_NewPremisData.[TB OD Tol  (-)]"
    Alias ="Expr59"
    Name ="TB Color Series"
    Expression ="tabl_FiberConversion_NewPremisData.[TB Color Series]"
    Alias ="Expr60"
    Name ="Fiber Type"
    Expression ="tabl_FiberConversion_NewPremisData.[Fiber Type]"
    Alias ="Expr61"
    Name ="Fiber Oracle item"
    Expression ="tabl_FiberConversion_NewPremisData.[Fiber Oracle item]"
    Alias ="Expr62"
    Name ="1st Req Freq"
    Expression ="tabl_FiberConversion_NewPremisData.[1st Req Freq]"
    Alias ="Expr63"
    Name ="1st Max Atten"
    Expression ="tabl_FiberConversion_NewPremisData.[1st Max Atten]"
    Alias ="Expr64"
    Name ="1st Min BandW"
    Expression ="tabl_FiberConversion_NewPremisData.[1st Min BandW]"
    Alias ="Expr65"
    Name ="2nd Req Freq"
    Expression ="tabl_FiberConversion_NewPremisData.[2nd Req Freq]"
    Alias ="Expr66"
    Name ="2nd Max Atten"
    Expression ="tabl_FiberConversion_NewPremisData.[2nd Max Atten]"
    Alias ="Expr67"
    Name ="2nd Min BandW"
    Expression ="tabl_FiberConversion_NewPremisData.[2nd Min BandW]"
    Alias ="Expr68"
    Name ="SM 1300 Max Atten"
    Expression ="tabl_FiberConversion_NewPremisData.[SM 1300 Max Atten]"
    Alias ="Expr69"
    Name ="SM 1550 Max Atten"
    Expression ="tabl_FiberConversion_NewPremisData.[SM 1550 Max Atten]"
    Alias ="Expr70"
    Name ="Special Instr Product1"
    Expression ="tabl_FiberConversion_NewPremisData.[Special Instr Product1]"
    Alias ="Expr71"
    Name ="zzSpecial Instr Product3"
    Expression ="tabl_FiberConversion_NewPremisData.[zzSpecial Instr Product3]"
    Alias ="Expr72"
    Name ="zzSpecial Instr Product4"
    Expression ="tabl_FiberConversion_NewPremisData.[zzSpecial Instr Product4]"
    Alias ="Expr73"
    Name ="zzOracle Part#"
    Expression ="tabl_FiberConversion_NewPremisData.[zzOracle Part#]"
    Alias ="Expr74"
    Name ="OracleStatus"
    Expression ="tabl_FiberConversion_NewPremisData.OracleStatus"
    Alias ="Expr75"
    Name ="PSS Document #"
    Expression ="tabl_FiberConversion_NewPremisData.[PSS Document #]"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbBinary "GUID" = Begin
    0xbd47d5a310493d48827b24dc277aa7d4
End
Begin
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa8cb3c542f22414e88bb6dac88eb84ac
        End
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4fb4a372d96d7642833e0158dd0e78b1
        End
    End
    Begin
        dbText "Name" ="Expr3"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x0a38351e9234ce4286b0877e32c1e73d
        End
    End
    Begin
        dbText "Name" ="Expr4"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc729328783201d42a8a4dae4ee4c61ca
        End
    End
    Begin
        dbText "Name" ="Expr5"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3b46888dbeb7854fa18fe5ec5ccd3d2b
        End
    End
    Begin
        dbText "Name" ="Expr6"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xadb5bc25aaf9d546816f6b7806bfcb9e
        End
    End
    Begin
        dbText "Name" ="Expr7"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xfb44c0133282dc4f9e655e6dafa6e6fb
        End
    End
    Begin
        dbText "Name" ="Expr8"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x030edd43c75fed419efb451f8018adc2
        End
    End
    Begin
        dbText "Name" ="Expr9"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x8c6e4cdf22ac3b4bb873b5f1d630baba
        End
    End
    Begin
        dbText "Name" ="Expr10"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x1cb8fa8839f69049912e02e77c361dd4
        End
    End
    Begin
        dbText "Name" ="Expr11"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6debbf6f467a0f44badcd73b9bd0023b
        End
    End
    Begin
        dbText "Name" ="Expr12"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x648dcdc8553c0140bdfcdf971ad20b3d
        End
    End
    Begin
        dbText "Name" ="Expr13"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x515d02a3aa8c744ba9f0af69d3530ff4
        End
    End
    Begin
        dbText "Name" ="Expr14"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd3f93df8e4caba4ba30df008aa3ac957
        End
    End
    Begin
        dbText "Name" ="Expr15"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x80801fae1965ad4a988f5fd547d4f276
        End
    End
    Begin
        dbText "Name" ="Expr16"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd417c6919f2d6e47bb4eb417f796619b
        End
    End
    Begin
        dbText "Name" ="Expr17"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x95069753f84da94fb8226a0f53a96fec
        End
    End
    Begin
        dbText "Name" ="Expr18"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x8d293bb2d91e5346883681cc63e948a5
        End
    End
    Begin
        dbText "Name" ="Expr19"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd3f7d4b32a6bb94a98cd7d9320fc2cdd
        End
    End
    Begin
        dbText "Name" ="Expr20"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x22e2e3ef213de64fafed4667efe2a3e4
        End
    End
    Begin
        dbText "Name" ="Expr21"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x0e5115aaf7944c41b8495666afff6fae
        End
    End
    Begin
        dbText "Name" ="Expr22"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7a4dd0824294c2468212c2cb8616ac2f
        End
    End
    Begin
        dbText "Name" ="Expr23"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa6c604b3ca781945bc556edbbcd014d7
        End
    End
    Begin
        dbText "Name" ="Expr24"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x04d4cb27bb2f7443b8c4af19e5824545
        End
    End
    Begin
        dbText "Name" ="Expr25"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc4cefe414ee283418872eb35199a1947
        End
    End
    Begin
        dbText "Name" ="Expr26"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe76ad22c2e98d747971a457c71328991
        End
    End
    Begin
        dbText "Name" ="Expr27"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xcfa06dff9932f4429da6a1030366fdd3
        End
    End
    Begin
        dbText "Name" ="Expr28"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x38380fa492a9454c9fd21b07d66a0459
        End
    End
    Begin
        dbText "Name" ="Expr29"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xfb83e6b3f0144a4d89731f203a68dbdd
        End
    End
    Begin
        dbText "Name" ="Expr30"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe25e12693034ec42abe4afd222c1f4f7
        End
    End
    Begin
        dbText "Name" ="Expr31"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd022bdcdd1baa04c98227bb9140a2af7
        End
    End
    Begin
        dbText "Name" ="Expr32"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xdcf2e81dbf67134899c88cb2050645bc
        End
    End
    Begin
        dbText "Name" ="Expr33"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe51ea2fca32df447900a7a61b2f17497
        End
    End
    Begin
        dbText "Name" ="Expr34"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa17002ff81ed4f43bf217dc899071c3a
        End
    End
    Begin
        dbText "Name" ="Expr35"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc4f7709bea7745489385ffbb3ff53dcd
        End
    End
    Begin
        dbText "Name" ="Expr36"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x5a88879d6b261a4ea4f53cdac1988d5a
        End
    End
    Begin
        dbText "Name" ="Expr37"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x5af75d911ad102459c757caa1fa17f79
        End
    End
    Begin
        dbText "Name" ="Expr38"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb02cbd8da33f4c47b46ad51fb5d5c351
        End
    End
    Begin
        dbText "Name" ="Expr39"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xc4e3a2252549554b9c84e808134ff58c
        End
    End
    Begin
        dbText "Name" ="Expr40"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3c1ce8a06ae29041a8fba7acfbed714a
        End
    End
    Begin
        dbText "Name" ="Expr41"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xde1688817241c241b565852e22aab5b9
        End
    End
    Begin
        dbText "Name" ="Expr42"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf24a2d866179284394d237a5ab295ef5
        End
    End
    Begin
        dbText "Name" ="Expr43"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x042ed2603eda594aa479bdf15dd6c4bb
        End
    End
    Begin
        dbText "Name" ="Expr44"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x6b2ad488e0a9c24cbf178e3571ea514d
        End
    End
    Begin
        dbText "Name" ="Expr45"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa3cae5b9ab773d49bf08f7d4d1f5eb0d
        End
    End
    Begin
        dbText "Name" ="Expr46"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x09479456e111194da034a4fd3aef2cc6
        End
    End
    Begin
        dbText "Name" ="Expr47"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x73f19ee71a2a9c4f9fbda0d5efe48f41
        End
    End
    Begin
        dbText "Name" ="Expr48"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xe7ebd84323c2e843a7e5445e3ffbd60c
        End
    End
    Begin
        dbText "Name" ="Expr49"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa97c01aaeb39044398ba6f2bcfcd365a
        End
    End
    Begin
        dbText "Name" ="Expr50"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x0c3e45e08ccb0347a39d14facbc791a6
        End
    End
    Begin
        dbText "Name" ="Expr51"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xfdcf6166cda89546801a4462ba2e3679
        End
    End
    Begin
        dbText "Name" ="Expr52"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x31ee040f1be8494084878e9ccdf55623
        End
    End
    Begin
        dbText "Name" ="Expr53"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf130c70397539c45ba09d9c6f723e428
        End
    End
    Begin
        dbText "Name" ="Expr54"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x31c223f5ea29dd4dbbdb411bb288b340
        End
    End
    Begin
        dbText "Name" ="Expr55"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd8d76b765d372e45be9b952d7a7fd66a
        End
    End
    Begin
        dbText "Name" ="Expr56"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xbadee20668d1c74097f24b501a9123f9
        End
    End
    Begin
        dbText "Name" ="Expr57"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x980aa39145e5c348938aea3369b17c36
        End
    End
    Begin
        dbText "Name" ="Expr58"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xafc1c66e290c0f46a96a680e8b3e5ed0
        End
    End
    Begin
        dbText "Name" ="Expr59"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x301f99372c5d1c4d8369931b358f000a
        End
    End
    Begin
        dbText "Name" ="Expr60"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x03ae70dcbd26c741b3ae736dd8bc0ba8
        End
    End
    Begin
        dbText "Name" ="Expr61"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x149113fee2d1994ab1c4a0f8fe8d1023
        End
    End
    Begin
        dbText "Name" ="Expr62"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x71d5da0182e1c74b9bb3edf4a1e105c0
        End
    End
    Begin
        dbText "Name" ="Expr63"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x1138d5b234f6a6468aa4853922e16be5
        End
    End
    Begin
        dbText "Name" ="Expr64"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x042cd33cde02ce428083f718b9bf41fb
        End
    End
    Begin
        dbText "Name" ="Expr65"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4d39263050b7b44ea6a64d554d058146
        End
    End
    Begin
        dbText "Name" ="Expr66"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x7bcff81cdb467c43a8045ec2354fa731
        End
    End
    Begin
        dbText "Name" ="Expr67"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xf7e3f003e24158449c52b472c44bbe00
        End
    End
    Begin
        dbText "Name" ="Expr68"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x18a2891767b2df44870036cbc7d641e6
        End
    End
    Begin
        dbText "Name" ="Expr69"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x8245744c369cba4f94d26303db068b2d
        End
    End
    Begin
        dbText "Name" ="Expr70"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x60f8778d845a274884af964eb62eea38
        End
    End
    Begin
        dbText "Name" ="Expr71"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x4c8b1ff065aad54786b40e147da9920e
        End
    End
    Begin
        dbText "Name" ="Expr72"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb1f241c39ca0e1408abb1dc163e51bae
        End
    End
    Begin
        dbText "Name" ="Expr73"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xee6adafb562a4a4c8026d97af0b43c4f
        End
    End
    Begin
        dbText "Name" ="Expr74"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x594de3b9946d01438b153fb2a5a41344
        End
    End
    Begin
        dbText "Name" ="Expr75"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x69c059741185e944bde2b600807ed667
        End
    End
End
Begin
    State =0
    Left =142
    Top =59
    Right =1434
    Bottom =758
    Left =-1
    Top =-1
    Right =1260
    Bottom =382
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =38
        Top =6
        Right =430
        Bottom =394
        Top =0
        Name ="tabl_FiberConversion_NewPremisData"
        Name =""
    End
End

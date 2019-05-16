Version =20
VersionRequired =20
Begin Report
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    FilterOn = NotDefault
    AllowDesignChanges = NotDefault
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =9180
    DatasheetFontHeight =10
    ItemSuffix =31
    Left =5235
    Top =1635
    DatasheetGridlinesColor =12632256
    Filter ="Oracle ='PX03384-01'"
    RecSrcDt = Begin
        0x69186580acfee440
    End
    RecordSource ="SELECT [Basic Product Construction].[New Oracle Part #] AS Oracle, [Basic Produc"
        "t Construction].[Item No], [Basic Product Construction].Customer, [Basic Product"
        " Construction].[Customer Part#], [Basic Product Construction].CustomerRev, [Basi"
        "c Product Construction].Active, [Basic Product Construction].Reason, [Basic Prod"
        "uct Construction].[calculate bill?], [Basic Product Construction].Base, [Basic P"
        "roduct Construction].[Cable Type], [Basic Product Construction].[Revision Date],"
        " [Basic Product Construction].[Revision Letter], [Basic Product Construction].[R"
        "evision History], [Basic Product Construction].[Standard Operation], [Basic Prod"
        "uct Construction].[Jacket Material], [Basic Product Construction].[Jacket Color]"
        ", [Basic Product Construction].[Color Chip ID], [Basic Product Construction].[No"
        "minal OD], [Basic Product Construction].[OD Tol  (+)], [Basic Product Constructi"
        "on].[OD Tol  (-)], [Basic Product Construction].Height, [Basic Product Construct"
        "ion].[Height Tol (+)], [Basic Product Construction].[Height Tol (-)], [Basic Pro"
        "duct Construction].[Nominal Wall], [Basic Product Construction].[Max Ave Wall], "
        "[Basic Product Construction].[Min Ave Wall], [Basic Product Construction].[Min S"
        "pot Wall], [Basic Product Construction].[Print Reel No], [Basic Product Construc"
        "tion].[Print Item No], [Basic Product Construction].[Std Subunit Print], [Basic "
        "Product Construction].[Print Type (base)], [Basic Product Construction].[Print L"
        "ine 1], [Basic Product Construction].[Print Line 2], [Basic Product Construction"
        "].[Print Line 3], [Basic Product Construction].[Print Line 4], [Basic Product Co"
        "nstruction].NewPrintLine4, [Basic Product Construction].[Print Spacing], [Basic "
        "Product Construction].PrintNotes, [Basic Product Construction].[Listing Company]"
        ", [Basic Product Construction].[UL-ETL-Listing], [Basic Product Construction].[U"
        "L-ETL-Const], [Basic Product Construction].[UL-ETL-Section], [Basic Product Cons"
        "truction].ENumber, [Basic Product Construction].[Label Type], [Basic Product Con"
        "struction].[Aramid Type], [Basic Product Construction].[Number of ends], [Basic "
        "Product Construction].[Aramid Type2], [Basic Product Construction].[Number of en"
        "ds2], [Basic Product Construction].[Unit Chips Type], [Basic Product Constructio"
        "n].[Unit Ripcord], [Basic Product Construction].[Unit CM], [Basic Product Constr"
        "uction].[Lay Length], [Basic Product Construction].Talc, [Basic Product Construc"
        "tion].[EZ Strip], [Basic Product Construction].[TB Material], [Basic Product Con"
        "struction].[TB Chips Type], [Basic Product Construction].[TB Nominal OD], [Basic"
        " Product Construction].[TB OD Tol  (+)], [Basic Product Construction].[TB OD Tol"
        "  (-)], [Basic Product Construction].[TB Color Series], [Basic Product Construct"
        "ion].[Fiber Type], [Basic Product Construction].[Fiber Oracle item], [Basic Prod"
        "uct Construction].FiberType2, [Basic Product Construction].FiberType3, [Basic Pr"
        "oduct Construction].[1st Req Freq], [Basic Product Construction].[1st Max Atten]"
        ", [Basic Product Construction].[1st Min BandW], [Basic Product Construction].[2n"
        "d Req Freq], [Basic Product Construction].[2nd Max Atten], [Basic Product Constr"
        "uction].[2nd Min BandW], [Basic Product Construction].[SM 1300 Max Atten], [Basi"
        "c Product Construction].[SM 1550 Max Atten], [Basic Product Construction].[Speci"
        "al Instr Product1], [Basic Product Construction].[zzSpecial Instr Product3], [Ba"
        "sic Product Construction].[zzSpecial Instr Product4], [Basic Product Constructio"
        "n].[zzOracle Part#], [Basic Product Construction].OracleStatus, [Basic Product C"
        "onstruction].[PSS Document #], [Basic Product Construction].DateCreated, [Basic "
        "Product Construction].CreatedBy, [Basic Product Construction].RevisedBy, [Basic "
        "Product Construction].RevisedDate, tblCableConstructionReferences.NumSubFillers,"
        " 1 AS Location, tblCableConstructionReferences.BaseID, [Basic Product Constructi"
        "on].Base FROM tblCableConstructionReferences INNER JOIN [Basic Product Construct"
        "ion] ON tblCableConstructionReferences.Base = [Basic Product Construction].Base;"
        " "
    Caption ="Cable Construction"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000dc230000b838000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    RecordSource ="SELECT [Basic Product Construction].[New Oracle Part #] AS Oracle, [Basic Produc"
        "t Construction].[Item No], [Basic Product Construction].Customer, [Basic Product"
        " Construction].[Customer Part#], [Basic Product Construction].CustomerRev, [Basi"
        "c Product Construction].Active, [Basic Product Construction].Reason, [Basic Prod"
        "uct Construction].[calculate bill?], [Basic Product Construction].Base, [Basic P"
        "roduct Construction].[Cable Type], [Basic Product Construction].[Revision Date],"
        " [Basic Product Construction].[Revision Letter], [Basic Product Construction].[R"
        "evision History], [Basic Product Construction].[Standard Operation], [Basic Prod"
        "uct Construction].[Jacket Material], [Basic Product Construction].[Jacket Color]"
        ", [Basic Product Construction].[Color Chip ID], [Basic Product Construction].[No"
        "minal OD], [Basic Product Construction].[OD Tol  (+)], [Basic Product Constructi"
        "on].[OD Tol  (-)], [Basic Product Construction].Height, [Basic Product Construct"
        "ion].[Height Tol (+)], [Basic Product Construction].[Height Tol (-)], [Basic Pro"
        "duct Construction].[Nominal Wall], [Basic Product Construction].[Max Ave Wall], "
        "[Basic Product Construction].[Min Ave Wall], [Basic Product Construction].[Min S"
        "pot Wall], [Basic Product Construction].[Print Reel No], [Basic Product Construc"
        "tion].[Print Item No], [Basic Product Construction].[Std Subunit Print], [Basic "
        "Product Construction].[Print Type (base)], [Basic Product Construction].[Print L"
        "ine 1], [Basic Product Construction].[Print Line 2], [Basic Product Construction"
        "].[Print Line 3], [Basic Product Construction].[Print Line 4], [Basic Product Co"
        "nstruction].NewPrintLine4, [Basic Product Construction].[Print Spacing], [Basic "
        "Product Construction].PrintNotes, [Basic Product Construction].[Listing Company]"
        ", [Basic Product Construction].[UL-ETL-Listing], [Basic Product Construction].[U"
        "L-ETL-Const], [Basic Product Construction].[UL-ETL-Section], [Basic Product Cons"
        "truction].ENumber, [Basic Product Construction].[Label Type], [Basic Product Con"
        "struction].[Aramid Type], [Basic Product Construction].[Number of ends], [Basic "
        "Product Construction].[Aramid Type2], [Basic Product Construction].[Number of en"
        "ds2], [Basic Product Construction].[Unit Chips Type], [Basic Product Constructio"
        "n].[Unit Ripcord], [Basic Product Construction].[Unit CM], [Basic Product Constr"
        "uction].[Lay Length], [Basic Product Construction].Talc, [Basic Product Construc"
        "tion].[EZ Strip], [Basic Product Construction].[TB Material], [Basic Product Con"
        "struction].[TB Chips Type], [Basic Product Construction].[TB Nominal OD], [Basic"
        " Product Construction].[TB OD Tol  (+)], [Basic Product Construction].[TB OD Tol"
        "  (-)], [Basic Product Construction].[TB Color Series], [Basic Product Construct"
        "ion].[Fiber Type], [Basic Product Construction].[Fiber Oracle item], [Basic Prod"
        "uct Construction].FiberType2, [Basic Product Construction].FiberType3, [Basic Pr"
        "oduct Construction].[1st Req Freq], [Basic Product Construction].[1st Max Atten]"
        ", [Basic Product Construction].[1st Min BandW], [Basic Product Construction].[2n"
        "d Req Freq], [Basic Product Construction].[2nd Max Atten], [Basic Product Constr"
        "uction].[2nd Min BandW], [Basic Product Construction].[SM 1300 Max Atten], [Basi"
        "c Product Construction].[SM 1550 Max Atten], [Basic Product Construction].[Speci"
        "al Instr Product1], [Basic Product Construction].[zzSpecial Instr Product3], [Ba"
        "sic Product Construction].[zzSpecial Instr Product4], [Basic Product Constructio"
        "n].[zzOracle Part#], [Basic Product Construction].OracleStatus, [Basic Product C"
        "onstruction].[PSS Document #], [Basic Product Construction].DateCreated, [Basic "
        "Product Construction].CreatedBy, [Basic Product Construction].RevisedBy, [Basic "
        "Product Construction].RevisedDate, tblCableConstructionReferences.NumSubFillers,"
        " 1 AS Location, tblCableConstructionReferences.BaseID, [Basic Product Constructi"
        "on].Base FROM tblCableConstructionReferences INNER JOIN [Basic Product Construct"
        "ion] ON tblCableConstructionReferences.Base = [Basic Product Construction].Base;"
        " "
    FilterOnLoad =255
    FitToPage =1
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin Line
            BorderLineStyle =0
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            FontName ="Tahoma"
        End
        Begin ComboBox
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =3
            BorderShade =90.0
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Subform
            BorderLineStyle =0
            BorderThemeColorIndex =3
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BorderShade =90.0
            ShowPageHeaderAndPageFooter =1
        End
        Begin Section
            CanGrow = NotDefault
            Height =14520
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Subform
                    Top =60
                    Width =9180
                    Height =14340
                    BorderColor =15916467
                    Name ="Child29"
                    SourceObject ="Form.frmSinglePassCable"
                    LinkChildFields ="Oracle"
                    LinkMasterFields ="Oracle"
                    GridlineColor =10921638

                    LayoutCachedTop =60
                    LayoutCachedWidth =9180
                    LayoutCachedHeight =14400
                    Begin
                        Begin Label
                            Width =660
                            Height =240
                            Name ="Label30"
                            Caption ="Child29:"
                            LayoutCachedWidth =660
                            LayoutCachedHeight =240
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
